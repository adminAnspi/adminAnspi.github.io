// =====================================================
// BOOKING.GS - GESTIONE PRENOTAZIONI E CACHE
// Versione Aggiornata con Headers Dinamici
// + updateBookingDetail integrato con helper RecurringUtils
// =====================================================

function formatDateValue(cellValue) {
  if (cellValue instanceof Date) {
    var year  = cellValue.getFullYear();
    var month = String(cellValue.getMonth() + 1).padStart(2, '0');
    var day   = String(cellValue.getDate()).padStart(2, '0');
    return year + '-' + month + '-' + day;
  } else if (typeof cellValue === 'string') {
    if (cellValue.includes('-') && cellValue.length >= 10) {
      return cellValue.split('T')[0];
    } else {
      var parsedDate = new Date(cellValue);
      if (!isNaN(parsedDate.getTime())) {
        var y = parsedDate.getFullYear();
        var m = String(parsedDate.getMonth() + 1).padStart(2, '0');
        var d = String(parsedDate.getDate()).padStart(2, '0');
        return y + '-' + m + '-' + d;
      }
    }
  }
  return String(cellValue);
}

function getBookingsForCalendar() {
  debugLog('🚀 getBookingsForCalendar: Caricamento con cache');
  try {
    var cache    = CacheService.getScriptCache();
    var cacheKey = CACHE_KEYS.PUBLIC_CALENDAR;
    var cachedData = cache.get(cacheKey);
    if (cachedData) {
      var bookings = JSON.parse(cachedData);
      debugLog('⚡ CACHE HIT: ' + bookings.length + ' eventi');
      return bookings;
    }

    var sheet   = getSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      cache.put(cacheKey, JSON.stringify([]), CACHE_DURATION);
      return [];
    }

    var allData = sheet.getDataRange().getValues();
    var headers = allData[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    var dataRows = allData.slice(1);
    var idCol = headers.indexOf('id');

    var result = dataRows
      .filter(function(row) { return row[idCol] && row[idCol].toString().trim() !== ''; })
      .map(function(row) {
        var booking = {};
        headers.forEach(function(header, i) {
          if (i < row.length) {
            var v = row[i];
            if (header === 'data' && v) {
              booking[header] = formatDateValue(v);
            } else if (v instanceof Date) {
              booking[header] = v.toISOString();
            } else if (header === 'ricorrente' || header === 'archiviato' || header === 'remindersent') {
              booking[header] = v === true || v === 'TRUE' || v === 'sì';
            } else {
              booking[header] = v;
            }
          }
        });
        return booking;
      })
      .filter(function(b) { return b.data && b.data.toString().trim() !== ''; });

    cache.put(cacheKey, JSON.stringify(result), CACHE_DURATION);
    debugLog('⚡ Caricati e cachati: ' + result.length + ' eventi');
    return result;

  } catch (error) {
    debugLog('❌ Errore getBookingsForCalendar: ' + error.message);
    return [];
  }
}

function getSheet() {
  try {
    SpreadsheetApp.flush();
    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAMES.EVENTS);
    if (!sheet) throw new Error('Foglio "' + SHEET_NAMES.EVENTS + '" non trovato.');
    SpreadsheetApp.flush();
    return sheet;
  } catch (error) {
    debugLog('❌ Errore in getSheet: ' + error.message);
    throw error;
  }
}

function debugLog(message) {
  if (typeof console !== 'undefined' && console.log) console.log(message);
  Logger.log(message);
}

function deleteBooking(bookingId) {
  debugLog('🗑️ deleteBooking chiamata per ID: ' + bookingId);
  if (!bookingId) return { status: 'error', message: 'ID prenotazione mancante' };

  try {
    var sheet = getSheet();
    var data  = sheet.getDataRange().getValues();
    if (data.length < 2) return { status: 'error', message: 'Nessun dato nel foglio' };

    var headers    = data[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    var idColIndex = headers.indexOf('id');
    if (idColIndex === -1) return { status: 'error', message: 'Struttura foglio non valida' };

    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][idColIndex] && data[i][idColIndex].toString() === bookingId.toString()) {
        rowIndex = i;
        break;
      }
    }
    if (rowIndex === -1) return { status: 'error', message: 'Prenotazione non trovata' };

    sheet.deleteRow(rowIndex + 1);
    debugLog('✅ Prenotazione ' + bookingId + ' eliminata');
    return { status: 'ok', message: 'Prenotazione eliminata con successo' };

  } catch (error) {
    return { status: 'error', message: 'Errore eliminazione: ' + error.message };
  }
}

function getBookings() {
  return getOrSetCache(CACHE_KEYS.BOOKINGS, function() {
    var sheet  = getSheet();
    var values = sheet.getDataRange().getValues();
    if (values.length < 2) return [];

    var headers  = values[0].map(function(h) { return h.toLowerCase().trim(); });
    var dataRows = values.slice(1);

    return dataRows
      .map(function(row) {
        var booking = {};
        headers.forEach(function(header, i) {
          if (i < row.length) {
            var v = row[i];
            if (header === 'data' && v) {
              booking[header] = formatDateValue(v);
            } else if (v instanceof Date) {
              booking[header] = v.toISOString();
            } else if (header === 'ricorrente' || header === 'archiviato' || header === 'remindersent') {
              booking[header] = v === true || v === 'TRUE' || v === 'sì';
            } else {
              booking[header] = v;
            }
          }
        });
        return booking;
      })
      .filter(function(b) {
        return b.id && b.id.toString().trim() !== '' && b.data && b.data.toString().trim() !== '';
      });
  }, CACHE_DURATION);
}

function getArchivedBookings() {
  debugLog('📦 getArchivedBookings: Caricamento eventi archiviati');
  try {
    var cache    = CacheService.getScriptCache();
    var cacheKey = CACHE_KEYS.ARCHIVED;
    var cachedData = cache.get(cacheKey);

    if (cachedData) {
      try {
        var bookings = JSON.parse(cachedData);
        if (Array.isArray(bookings)) {
          debugLog('⚡ CACHE HIT archivio: ' + bookings.length + ' eventi');
          return { status: 'ok', data: { bookings: bookings } };
        } else if (bookings && typeof bookings === 'object') {
          var converted = Object.values(bookings);
          cache.remove(cacheKey);
          cache.put(cacheKey, JSON.stringify(converted), 600);
          return { status: 'ok', data: { bookings: converted } };
        } else {
          cache.remove(cacheKey);
        }
      } catch (parseError) {
        cache.remove(cacheKey);
      }
    }

    var ss           = SpreadsheetApp.openById(SHEET_ID);
    var archiveSheet = ss.getSheetByName(SHEET_NAMES.ARCHIVE);
    if (!archiveSheet) return { status: 'ok', data: { bookings: [] } };

    var allData = archiveSheet.getDataRange().getValues();
    if (allData.length < 2) {
      cache.put(cacheKey, JSON.stringify([]), 600);
      return { status: 'ok', data: { bookings: [] } };
    }

    var headers  = allData[0].map(function(h) { return h ? h.toString().toLowerCase().trim() : ''; });
    var dataRows = allData.slice(1);

    var result = dataRows
      .map(function(row) {
        var booking = {};
        headers.forEach(function(header, i) {
          if (header && i < row.length) {
            var v = row[i];
            if (header === 'data' && v) {
              booking[header] = formatDateValue(v);
            } else if (v instanceof Date) {
              booking[header] = v.toISOString();
            } else if (header === 'ricorrente' || header === 'archiviato' || header === 'remindersent') {
              booking[header] = v === true || v === 'TRUE' || v === 'sì';
            } else {
              booking[header] = v;
            }
          }
        });
        return booking;
      })
      .filter(function(b) {
        return b.id && b.id.toString().trim() !== '' && b.data && b.data.toString().trim() !== '';
      });

    cache.put(cacheKey, JSON.stringify(result), 600);
    debugLog('✅ Archivio caricato: ' + result.length + ' eventi');
    return { status: 'ok', data: { bookings: result } };

  } catch (error) {
    debugLog('❌ Errore fatale getArchivedBookings: ' + error.message);
    return { status: 'ok', data: { bookings: [] } };
  }
}

function diagnoseDuplicateBookings() {
  try {
    var sheet = getSheet();
    var all   = sheet.getDataRange().getValues();
    if (all.length < 2) return { status: 'ok', data: { totalRows: 0, byId: [], byComposite: [] } };

    var headers = all[0].map(function(h) { return h ? h.toString().toLowerCase().trim() : ''; });
    var idx = {
      id: headers.indexOf('id'), nome: headers.indexOf('nome'),
      data: headers.indexOf('data'), ora: headers.indexOf('ora'),
      campo: headers.indexOf('campo'), evento: headers.indexOf('evento'),
      ricorrente: headers.indexOf('ricorrente'), timestamp: headers.indexOf('timestamp'),
      stato: headers.indexOf('stato')
    };

    var byIdMap = {};
    var byCompositeMap = {};
    var rows = all.slice(1);

    rows.forEach(function(row, i) {
      var idVal        = idx.id >= 0 ? row[idx.id] : '';
      var compositeKey = [row[idx.nome], formatDateValue(row[idx.data]), row[idx.ora], row[idx.campo], row[idx.evento]].join('|');
      var rowInfo = {
        rowIndex: i + 2, id: idVal, nome: row[idx.nome],
        data: formatDateValue(row[idx.data]), ora: row[idx.ora],
        campo: row[idx.campo], evento: row[idx.evento],
        ricorrente: row[idx.ricorrente] === true || row[idx.ricorrente] === 'true',
        timestamp: row[idx.timestamp], stato: row[idx.stato]
      };
      if (idVal) {
        if (!byIdMap[idVal]) byIdMap[idVal] = [];
        byIdMap[idVal].push(rowInfo);
      }
      if (!byCompositeMap[compositeKey]) byCompositeMap[compositeKey] = [];
      byCompositeMap[compositeKey].push(rowInfo);
    });

    var idDuplicates = [];
    for (var id in byIdMap) { if (byIdMap[id].length > 1) idDuplicates.push({ key: id, count: byIdMap[id].length, rows: byIdMap[id] }); }
    var compositeDuplicates = [];
    for (var key in byCompositeMap) { if (byCompositeMap[key].length > 1) compositeDuplicates.push({ key: key, count: byCompositeMap[key].length, rows: byCompositeMap[key] }); }

    return { status: 'ok', data: {
      stats: { totalRows: rows.length, duplicateIds: idDuplicates.length, duplicateComposite: compositeDuplicates.length },
      byId: idDuplicates, byComposite: compositeDuplicates
    }};
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

// =====================================================
// updateBookingDetail
// Integrato con gli helper condivisi di RecurringUtils:
//   - findNextAvailableDate(startDate, ora, campo)
//   - createVirtualEvent(masterRow, masterHeaders, targetDate)
//   - checkTimeConflict(date, time, campo)
// =====================================================
function updateBookingDetail(bookingId, field, value) {
  try {
    if (!bookingId || !field) {
      return { status: 'error', message: 'Parametri mancanti' };
    }

    var sheet   = getSheet();
    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return h ? h.toString().toLowerCase().trim() : ''; });

    var idCol    = headers.indexOf('id');
    var fieldCol = headers.indexOf(field.toLowerCase().trim());

    if (idCol === -1)    return { status: 'error', message: 'Colonna ID non trovata' };
    if (fieldCol === -1) return { status: 'error', message: 'Campo "' + field + '" non trovato nel foglio' };

    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][idCol] && data[i][idCol].toString().trim() === bookingId.toString().trim()) {
        rowIndex = i;
        break;
      }
    }
    if (rowIndex === -1) return { status: 'error', message: 'Prenotazione non trovata: ' + bookingId };

    // ===== GESTIONE CAMPO RICORRENTE =====
    if (field.toLowerCase() === 'ricorrente') {
      var boolValue = value === true || value === 'true';
      sheet.getRange(rowIndex + 1, fieldCol + 1).setValue(boolValue);
      SpreadsheetApp.flush();
      invalidateAllCache();

      var deletedFutureEvents = 0;

      if (!boolValue) {
        // Disabilita ricorrenza: elimina eventuali eventi virtuali (-V)
        var virtualId = bookingId + '-V';
        var freshData = sheet.getDataRange().getValues();
        var freshHeaders = freshData[0].map(function(h) { return h ? h.toString().toLowerCase().trim() : ''; });
        var freshIdCol = freshHeaders.indexOf('id');
        for (var j = freshData.length - 1; j >= 1; j--) {
          if (freshData[j][freshIdCol] && freshData[j][freshIdCol].toString().trim() === virtualId) {
            sheet.deleteRow(j + 1);
            deletedFutureEvents++;
            SpreadsheetApp.flush();
          }
        }
        invalidateAllCache();
        debugLog('✅ Ricorrenza disabilitata, ' + deletedFutureEvents + ' eventi virtuali eliminati');

      } else {
        // Abilita ricorrenza: crea evento virtuale per la settimana successiva
        // Usa findNextAvailableDate e createVirtualEvent da RecurringUtils.gs
        try {
          var allBookingsList = getBookings();
          var masterBooking = null;
          for (var k = 0; k < allBookingsList.length; k++) {
            if (allBookingsList[k].id === bookingId) { masterBooking = allBookingsList[k]; break; }
          }

          if (masterBooking) {
            var bookingDate   = new Date(masterBooking.data);
            var nextAvailable = findNextAvailableDate(bookingDate, masterBooking.ora, masterBooking.campo);

            if (nextAvailable) {
              // Rileggi il foglio aggiornato per passare la riga corretta
              var updatedData    = sheet.getDataRange().getValues();
              var updatedHeaders = updatedData[0];
              var masterRow      = updatedData[rowIndex];
              createVirtualEvent(masterRow, updatedHeaders, nextAvailable);
              invalidateAllCache();
              debugLog('✅ Ricorrenza abilitata, evento virtuale creato per ' + nextAvailable.toLocaleDateString('it-IT'));
            } else {
              debugLog('⚠️ Ricorrenza abilitata ma nessuna data libera nelle prossime 4 settimane');
            }
          }
        } catch (virtualErr) {
          debugLog('⚠️ Errore creazione evento virtuale: ' + virtualErr.message);
        }
      }

      return {
        status: 'ok',
        message: 'Campo ricorrente aggiornato',
        data: { bookingId: bookingId, field: field, value: boolValue, deletedFutureEvents: deletedFutureEvents }
      };
    }

    // ===== AGGIORNAMENTO GENERICO =====
    sheet.getRange(rowIndex + 1, fieldCol + 1).setValue(value);
    SpreadsheetApp.flush();
    invalidateAllCache();

    debugLog('✅ Campo "' + field + '" aggiornato per ' + bookingId);
    return {
      status: 'ok',
      message: 'Campo "' + field + '" aggiornato con successo',
      data: { bookingId: bookingId, field: field, value: value }
    };

  } catch (error) {
    debugLog('❌ Errore updateBookingDetail: ' + error.message);
    return { status: 'error', message: error.message };
  }
}
