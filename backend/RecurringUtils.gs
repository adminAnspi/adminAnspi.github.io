// =====================================================
// RECURRING UTILS - VERSIONE CORRETTA v3.3
// Fix: buffer 30 minuti post evento straordinario in checkTimeConflict
// Helper condivisi: findNextAvailableDate, createVirtualEvent
// =====================================================

function disableAutoArchive() {
  try {
    var cache = CacheService.getScriptCache();
    var disableUntil = new Date();
    disableUntil.setHours(disableUntil.getHours() + 2);
    cache.put('auto_archive_disabled', disableUntil.toISOString(), 7200);
    console.log('🚫 Archiviazione disabilitata fino alle ' + disableUntil.toLocaleString('it-IT'));
    return { status: 'ok', message: 'Archiviazione automatica disabilitata fino alle ' + disableUntil.toLocaleString('it-IT'), data: { disabledUntil: disableUntil.toISOString() } };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

function enableAutoArchive() {
  try {
    CacheService.getScriptCache().remove('auto_archive_disabled');
    return { status: 'ok', message: 'Archiviazione automatica riabilitata' };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

function checkAndGenerateRecurringEvents() {
  console.log('Avvio processo ricorrenze e archiviazione');

  var cache = CacheService.getScriptCache();
  var disabledUntil = cache.get('auto_archive_disabled');
  if (disabledUntil) {
    var disableTime = new Date(disabledUntil);
    if (new Date() < disableTime) {
      return { status: 'skipped', message: 'Archiviazione automatica disabilitata fino alle ' + disableTime.toLocaleString('it-IT') };
    } else {
      cache.remove('auto_archive_disabled');
    }
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    return { status: 'skipped', message: 'Processo già attivo' };
  }

  try {
    var sheet = getSheet();
    var data  = sheet.getDataRange().getValues();
    var headers = data[0];

    var cols = {
      id:         headers.findIndex(function(h) { return h.toLowerCase().includes('id'); }),
      nome:       headers.findIndex(function(h) { return h.toLowerCase().includes('nome'); }),
      data:       headers.findIndex(function(h) { return h.toLowerCase().includes('data'); }),
      ora:        headers.findIndex(function(h) { return h.toLowerCase().includes('ora'); }),
      evento:     headers.findIndex(function(h) { return h.toLowerCase().includes('evento'); }),
      campo:      headers.findIndex(function(h) { return h.toLowerCase().includes('campo'); }),
      ricorrente: headers.findIndex(function(h) { return h.toLowerCase().includes('ricorrente'); }),
      email:      headers.findIndex(function(h) { return h.toLowerCase().includes('email'); }),
      telefono:   headers.findIndex(function(h) { return h.toLowerCase().includes('telefono'); }),
      note:       headers.findIndex(function(h) { return h.toLowerCase().includes('note'); }),
      timestamp:  headers.findIndex(function(h) { return h.toLowerCase().includes('timestamp'); })
    };

    var now = new Date();
    var stats = { archived: 0, generated: 0, errors: 0 };
    var toArchive = [];
    var toGenerateRecurring = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[cols.id] || !row[cols.data] || !row[cols.ora]) continue;
      var eventId = row[cols.id].toString();
      if (eventId.includes('-V')) continue;

      var isRecurring   = row[cols.ricorrente] === true || row[cols.ricorrente] === 'true';
      var eventType     = row[cols.evento];
      var duration      = getDurationMinutes(eventType);
      var eventDateTime = new Date(new Date(row[cols.data]).toDateString() + ' ' + row[cols.ora]);
      var expiryTime    = new Date(eventDateTime.getTime() + (duration + 30) * 60000);

      if (now >= expiryTime) {
        if (isRecurring) {
          toGenerateRecurring.push({ row: row.slice(), rowIndex: i, cols: cols });
        } else {
          toArchive.push({ row: row.slice(), rowIndex: i });
        }
      }
    }

    console.log('Trovati: ' + toGenerateRecurring.length + ' ricorrenti, ' + toArchive.length + ' da archiviare');

    for (var g = 0; g < toGenerateRecurring.length; g++) {
      var item = toGenerateRecurring[g];
      try {
        var success = createNextRecurrence(item.row, item.cols, sheet);
        if (success) {
          stats.generated++;
          toArchive.push({ row: item.row, rowIndex: item.rowIndex });
        } else {
          stats.errors++;
        }
      } catch (error) {
        stats.errors++;
        console.log('Errore ricorrenza: ' + error.message);
      }
    }

    var sortedToArchive = toArchive.sort(function(a, b) { return b.rowIndex - a.rowIndex; });
    for (var a = 0; a < sortedToArchive.length; a++) {
      try {
        if (moveToArchive(sortedToArchive[a].row, sortedToArchive[a].rowIndex, sheet)) {
          stats.archived++;
        } else {
          stats.errors++;
        }
      } catch (error) {
        stats.errors++;
      }
    }

    SpreadsheetApp.flush();
    clearAllCaches();

    var message;
    if (stats.generated === 0 && stats.archived === 0) {
      message = 'Nessuna operazione necessaria.';
    } else if (stats.generated > 0 && stats.archived === 0) {
      message = stats.generated + ' ricorrenze generate';
    } else if (stats.generated === 0 && stats.archived > 0) {
      message = stats.archived + ' eventi archiviati';
    } else {
      message = stats.generated + ' ricorrenze generate, ' + stats.archived + ' eventi archiviati';
    }

    return { status: 'ok', data: stats, message: message };

  } catch (error) {
    console.error('Errore processo: ' + error.message);
    return { status: 'error', message: error.message };
  } finally {
    lock.releaseLock();
  }
}

// =====================================================
// HELPER CONDIVISI
// Usati da createNextRecurrence E da updateBookingDetail
// =====================================================

/**
 * Cerca la prima data disponibile partendo da startDate + 7 giorni.
 * Tenta fino a 4 settimane in caso di conflitto.
 * Restituisce un oggetto Date oppure null se nessuna data è libera.
 */
function findNextAvailableDate(startDate, ora, campo) {
  var targetDate = new Date(startDate.getTime() + 7 * 24 * 60 * 60 * 1000);
  var attempts = 0;
  while (attempts < 4) {
    if (!checkTimeConflict(targetDate, ora, campo)) {
      return targetDate;
    }
    debugLog('⚠️ Conflitto il ' + targetDate.toLocaleDateString('it-IT') + ', provo settimana successiva');
    targetDate = new Date(targetDate.getTime() + 7 * 24 * 60 * 60 * 1000);
    attempts++;
  }
  debugLog('❌ Nessuna data libera nelle prossime 4 settimane');
  return null;
}

/**
 * Crea un evento virtuale (-V) nel foglio a partire dalla riga master.
 * masterRow     — array dei valori della riga master
 * masterHeaders — array degli header del foglio (riga 0)
 * targetDate    — oggetto Date della data dell'evento virtuale
 */
function createVirtualEvent(masterRow, masterHeaders, targetDate) {
  try {
    var sheet = getSheet();
    var headersLower = masterHeaders.map(function(h) {
      return h ? h.toString().toLowerCase().trim() : '';
    });

    var idCol    = headersLower.indexOf('id');
    var dataCol  = headersLower.indexOf('data');
    var recurCol = headersLower.indexOf('ricorrente');
    var tsCol    = headersLower.indexOf('timestamp');
    var noteCol  = headersLower.indexOf('note');

    var newVirtualRow = masterRow.slice();
    var masterId = newVirtualRow[idCol].toString().replace('-V', '');

    newVirtualRow[idCol]    = masterId + '-V';
    newVirtualRow[dataCol]  = targetDate;
    newVirtualRow[recurCol] = false;
    if (tsCol >= 0)   newVirtualRow[tsCol] = new Date();
    if (noteCol >= 0) {
      var note = newVirtualRow[noteCol] ? newVirtualRow[noteCol].toString() : '';
      note = note.replace(' [Ricorrenza virtuale]', '');
      newVirtualRow[noteCol] = note + ' [Ricorrenza virtuale]';
    }

    sheet.appendRow(newVirtualRow);
    SpreadsheetApp.flush();
    debugLog('✅ Evento virtuale ' + masterId + '-V creato per ' + targetDate.toLocaleDateString('it-IT'));
    return true;
  } catch (error) {
    debugLog('❌ Errore createVirtualEvent: ' + error.message);
    return false;
  }
}

// =====================================================
// LOGICA RICORRENZA AUTOMATICA
// =====================================================

function createNextRecurrence(originalRow, cols, sheet) {
  try {
    var originalId    = originalRow[cols.id].toString();
    var originalTime  = originalRow[cols.ora];
    var originalCampo = originalRow[cols.campo];

    console.log('🔄 Processo ricorrenza per master ' + originalId);

    // PASSO 1: Cerca evento virtuale (-V)
    var virtualEventId = originalId + '-V';
    var data = sheet.getDataRange().getValues();
    var virtualEventRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][cols.id] && data[i][cols.id].toString() === virtualEventId) {
        virtualEventRow = i;
        break;
      }
    }
    if (virtualEventRow === -1) {
      console.log('❌ Evento virtuale ' + virtualEventId + ' non trovato');
      return false;
    }

    // PASSO 2: Promuovi virtuale a master
    var virtualEvent = data[virtualEventRow];
    sheet.getRange(virtualEventRow + 1, cols.id + 1).setValue(originalId);
    sheet.getRange(virtualEventRow + 1, cols.ricorrente + 1).setValue(true);
    if (cols.note >= 0 && virtualEvent[cols.note]) {
      sheet.getRange(virtualEventRow + 1, cols.note + 1).setValue(
        virtualEvent[cols.note].toString().replace(' [Ricorrenza virtuale]', '')
      );
    }
    SpreadsheetApp.flush();
    console.log('✅ ' + virtualEventId + ' promosso a master ' + originalId);

    // PASSO 3: Crea nuovo evento virtuale usando gli helper condivisi
    var virtualDate   = new Date(virtualEvent[cols.data]);
    var nextAvailable = findNextAvailableDate(virtualDate, originalTime, originalCampo);
    if (nextAvailable) {
      createVirtualEvent(virtualEvent, data[0], nextAvailable);
    } else {
      console.log('❌ Nessuna data disponibile nelle prossime 4 settimane, evento virtuale non creato');
    }

    return true;

  } catch (error) {
    console.error('❌ Errore createNextRecurrence: ' + error.message);
    return false;
  }
}

// =====================================================
// ARCHIVIAZIONE
// =====================================================

function moveToArchive(row, rowIndex, mainSheet) {
  try {
    var ss           = SpreadsheetApp.openById(SHEET_ID);
    var archiveSheet = ss.getSheetByName(SHEET_NAMES.ARCHIVE);

    if (!archiveSheet) {
      var hdrs = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
      archiveSheet = ss.insertSheet(SHEET_NAMES.ARCHIVE);
      archiveSheet.getRange(1, 1, 1, hdrs.length).setValues([hdrs]);
      archiveSheet.setFrozenRows(1);
    }

    var archiveData = archiveSheet.getDataRange().getValues();
    for (var i = 1; i < archiveData.length; i++) {
      if (archiveData[i][0] === row[0]) {
        mainSheet.deleteRow(rowIndex + 1);
        SpreadsheetApp.flush();
        return true;
      }
    }

    var archiveRow = row.slice();
    var noteIdx = archiveRow.length > 10 ? archiveRow.length - 2 : archiveRow.length - 1;
    archiveRow[noteIdx] = (archiveRow[noteIdx] || '') + ' [Archiviato: ' + new Date().toLocaleString('it-IT') + ']';
    archiveSheet.appendRow(archiveRow);
    SpreadsheetApp.flush();
    mainSheet.deleteRow(rowIndex + 1);
    SpreadsheetApp.flush();
    return true;

  } catch (error) {
    console.error('Errore archiviazione: ' + error.message);
    return false;
  }
}

// =====================================================
// CONTROLLO CONFLITTI v3.3
// Fix: buffer 30 minuti post evento straordinario
// Questo impedisce che una ricorrenza venga schedulata
// immediatamente dopo la fine di un evento straordinario.
// L'admin potrà sempre forzare manualmente se necessario.
// =====================================================

function checkTimeConflict(date, time, campo) {
  try {
    var allBookings = getBookings();
    var dateStr  = date.toISOString().split('T')[0];
    var newStart = new Date(dateStr + 'T' + time + ':00');
    var newEnd   = new Date(newStart.getTime() + (getDurationMinutes(campo) || 60) * 60000);

    for (var i = 0; i < allBookings.length; i++) {
      var b     = allBookings[i];
      var bDate = b.data ? b.data.split('T')[0] : '';
      if (bDate !== dateStr) continue;

      var isExtraordinary = b.evento === 'compleanno' || b.evento === 'eventi';

      if (!isExtraordinary) {
        // Controlla sovrapposizione campi solo per eventi normali
        var occupied  = b.campo === 'entrambi' ? ['calcetto', 'pallavolo'] : [b.campo];
        var requested = campo   === 'entrambi' ? ['calcetto', 'pallavolo'] : [campo];
        var fieldHit  = false;
        for (var c = 0; c < requested.length; c++) {
          if (occupied.indexOf(requested[c]) !== -1) { fieldHit = true; break; }
        }
        if (!fieldHit) continue;
      }

      // Calcola finestra temporale dell'evento esistente
      var existingDuration = getDurationMinutes(b.evento) || 60;
      var bStart = new Date(bDate + 'T' + b.ora + ':00');
      var bEnd   = new Date(bStart.getTime() + existingDuration * 60000);

      // ✅ FIX v3.3: Per eventi straordinari aggiungi 30 minuti di buffer post-evento.
      // Un evento che inizia esattamente quando finisce un evento straordinario
      // viene bloccato — l'admin può sbloccare manualmente se necessario.
      if (isExtraordinary) {
        bEnd = new Date(bEnd.getTime() + 30 * 60000);
        debugLog('📅 Evento straordinario ' + b.id + ': finestra estesa fino alle ' +
          bEnd.toLocaleTimeString('it-IT') + ' (buffer 30 min)');
      }

      if (newStart < bEnd && newEnd > bStart) {
        debugLog('⚠️ Conflitto con ' + b.id + ' (' + b.evento + ') il ' + bDate + ' alle ' + b.ora);
        return true;
      }
    }

    return false;

  } catch (error) {
    console.error('Errore checkTimeConflict: ' + error.message);
    return false;
  }
}

// =====================================================
// UTILITY
// =====================================================

function getDurationMinutes(eventType) {
  var d = { 'calcetto': 90, 'pallavolo': 120, 'compleanno': 180, 'eventi': 180 };
  return d[eventType] || 60;
}

function generateId() {
  var now = new Date();
  var t = now.getHours() * 10000 + now.getMinutes() * 100 + now.getSeconds();
  var r = Math.floor(Math.random() * 100);
  return 'ANSPI-' + ((t + r) % 100000).toString().padStart(5, '0');
}

function archiveOldBookings()   { return checkAndGenerateRecurringEvents(); }
function manualRecurringCheck() { return checkAndGenerateRecurringEvents(); }

function cleanupDuplicateRecurringEvents() {
  try {
    var sheet   = getSheet();
    var data    = sheet.getDataRange().getValues();
    var headers = data[0];
    var nameCol  = headers.findIndex(function(h) { return h.toLowerCase().includes('nome'); });
    var dateCol  = headers.findIndex(function(h) { return h.toLowerCase().includes('data'); });
    var timeCol  = headers.findIndex(function(h) { return h.toLowerCase().includes('ora'); });
    var recurCol = headers.findIndex(function(h) { return h.toLowerCase().includes('ricorrente'); });
    var seen = {};
    var toDelete = [];
    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      if (row[recurCol] === true || row[recurCol] === 'true') {
        var key = row[nameCol] + '_' + row[dateCol] + '_' + row[timeCol];
        if (seen[key]) { toDelete.push(i); } else { seen[key] = true; }
      }
    }
    for (var d = 0; d < toDelete.length; d++) { sheet.deleteRow(toDelete[d] + 1); }
    SpreadsheetApp.flush();
    clearAllCaches();
    return { status: 'ok', data: { deletedCount: toDelete.length }, message: toDelete.length + ' duplicati eliminati' };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}
// =====================================================
// checkInvisibleRecurrences
// Da aggiungere in RecurringUtils.gs
// =====================================================

/**
 * Controlla se la data/ora/campo richiesta coincide con una ricorrenza
 * "invisibile" — cioè non presente nel foglio ma proiettabile dalla
 * data master con passo settimanale (N >= 2, quindi dalla 3ª settimana
 * in poi: master+14, master+21, master+28, ...).
 *
 * Restituisce:
 *   { conflict: false }
 *   oppure
 *   { conflict: true, masterDate: 'YYYY-MM-DD', masterNome: '...', nextVisible: 'DD/MM/YYYY' }
 */
function checkInvisibleRecurrences(data, ora, campo) {
  try {
    var allBookings = getBookings();
    var targetDate  = new Date(data + 'T00:00:00');
    var targetTime  = targetDate.getTime();

    // Solo calcetto e pallavolo
    var relevantFields = ['calcetto', 'pallavolo'];
    if (relevantFields.indexOf(campo) === -1) {
      return { conflict: false };
    }

    var ONE_WEEK_MS = 7 * 24 * 60 * 60 * 1000;

    for (var i = 0; i < allBookings.length; i++) {
      var b = allBookings[i];

      // Solo master ricorrenti (non virtuali)
      var isRecurring = b.ricorrente === true || b.ricorrente === 'true';
      if (!isRecurring) continue;
      if (b.id && b.id.toString().indexOf('-V') !== -1) continue;

      // Solo stesso tipo campo
      if (b.campo !== campo && b.campo !== 'entrambi' && campo !== 'entrambi') continue;

      // Solo stesso orario
      if (b.ora !== ora) continue;

      // Calcola distanza in settimane dalla data master
      var masterDate = new Date(b.data + 'T00:00:00');
      var masterTime = masterDate.getTime();
      var diffMs     = targetTime - masterTime;

      // Deve essere un multiplo esatto di 7 giorni
      if (diffMs <= 0) continue;
      if (diffMs % ONE_WEEK_MS !== 0) continue;

      var weeksAhead = diffMs / ONE_WEEK_MS;

      // N >= 2 significa dalla 3ª settimana in poi (master=0, virtuale=1, invisibili=2+)
      if (weeksAhead < 2) continue;

      // Trovata una ricorrenza invisibile
      // Calcola la prossima data visibile (la virtuale = master + 7)
      var nextVisible = new Date(masterTime + ONE_WEEK_MS);
      var nextVisibleStr = nextVisible.toLocaleDateString('it-IT', {
        day: '2-digit', month: '2-digit', year: 'numeric'
      });

      debugLog('⚠️ Ricorrenza invisibile rilevata: ' + b.id +
        ' master ' + b.data + ' target ' + data +
        ' (' + weeksAhead + ' settimane avanti)');

      return {
        conflict:    true,
        bookingId:   b.id,
        masterDate:  b.data,
        masterNome:  b.nome,
        evento:      b.evento,
        campo:       b.campo,
        ora:         b.ora,
        weeksAhead:  weeksAhead,
        nextVisible: nextVisibleStr
      };
    }

    return { conflict: false };

  } catch (error) {
    debugLog('❌ Errore checkInvisibleRecurrences: ' + error.message);
    return { conflict: false };
  }
}
