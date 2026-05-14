// =====================
// Trigger and Setup Utilities
// =====================

function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // Setup Foglio Prenotazioni
  let prenotazioniSheet = ss.getSheetByName(SHEET_NAME);
  if (!prenotazioniSheet) {
    prenotazioniSheet = ss.insertSheet(SHEET_NAME, 0);
  }
  const headers = ['ID', 'Timestamp', 'Nome', 'Email', 'Telefono', 'Evento', 'Campo', 'Data', 'Ora', 'Note', 'Ricorrente', 'Stato', 'ReminderSent', 'AdminActionBy', 'Archiviato'];
  prenotazioniSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  prenotazioniSheet.setFrozenRows(1);

  // Setup Foglio Admin
  let adminSheet = ss.getSheetByName(ADMIN_SHEET_NAME);
  if (!adminSheet) {
    adminSheet = ss.insertSheet(ADMIN_SHEET_NAME, 1);
    const adminHeaders = ['Email', 'PasswordHash'];
    adminSheet.getRange(1, 1, 1, adminHeaders.length).setValues([adminHeaders]).setFontWeight("bold");
    adminSheet.setFrozenRows(1);
  }

  // Setup Foglio Reset Password
  let resetSheet = ss.getSheetByName("PasswordResets");
  if (!resetSheet) {
    resetSheet = ss.insertSheet("PasswordResets", 2);
    const resetHeaders = ['Email', 'Token', 'Expiration'];
    resetSheet.getRange(1, 1, 1, resetHeaders.length).setValues([resetHeaders]).setFontWeight("bold");
    resetSheet.setFrozenRows(1);
  }

  // Setup Foglio Archivio
  let archiveSheet = ss.getSheetByName(ARCHIVE_SHEET_NAME);
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet(ARCHIVE_SHEET_NAME, 3);
    archiveSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    archiveSheet.setFrozenRows(1);
  }

  SpreadsheetApp.flush();
  return { status: 'ok', message: 'Verifica setup completata.' };
}

function setupEventNotifications() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;

    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'checkUpcomingEventsAndNotify' ||
          trigger.getHandlerFunction() === 'sendDailySummaryToAdmin') {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
      }
    }

    debugLog(`🗑️ Rimossi ${deletedCount} trigger esistenti`);

    ScriptApp.newTrigger('checkUpcomingEventsAndNotify').timeBased().everyMinutes(30).create();
    ScriptApp.newTrigger('sendDailySummaryToAdmin').timeBased().everyDays(1).atHour(7).create();

    debugLog('✅ Sistema notifiche eventi attivato!');

    return {
      status: 'ok',
      message: 'Sistema di notifiche attivato con successo!',
      data: {
        triggersCreated: 2,
        triggersDeleted: deletedCount,
        schedule: {
          dailySummary: 'Ogni giorno alle 7:00',
          eventReminders: 'Ogni 30 minuti'
        }
      }
    };

  } catch (error) {
    debugLog(`❌ Errore setup notifiche: ${error.message}`);
    return {
      status: 'error',
      message: `Errore attivazione notifiche: ${error.message}`
    };
  }
}

/**
 * Crea un trigger one‑shot per inviare il reminder esattamente a start − 30 minuti.
 * Salva anche la prenotazione programmata in ScriptProperties per consentire al handler
 * di sapere quali ID inviare quando scatta.
 */
function scheduleSingleReminder(bookingId, bookingDateStr, bookingTimeStr) {
  try {
    if (!bookingId || !bookingDateStr || !bookingTimeStr) {
      debugLog('❌ scheduleSingleReminder: parametri mancanti');
      return { status: 'error', message: 'Parametri mancanti' };
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const timezone = ss.getSpreadsheetTimeZone();

    // Parse data e ora
    const eventDate = parseDateItalyOrISO(bookingDateStr);
    if (!eventDate) {
      debugLog(`❌ scheduleSingleReminder: data non valida (${bookingDateStr})`);
      return { status: 'error', message: 'Data non valida' };
    }
    const parts = bookingTimeStr.toString().split(':');
    if (parts.length < 2) {
      debugLog(`❌ scheduleSingleReminder: ora non valida (${bookingTimeStr})`);
      return { status: 'error', message: 'Ora non valida' };
    }
    const h = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10);
    const eventDateTime = new Date(eventDate.getFullYear(), eventDate.getMonth(), eventDate.getDate(), h, m, 0);

    const notifyTime = new Date(eventDateTime.getTime() - 30 * 60000);
    const now = new Date(new Date().toLocaleString('en-US', { timeZone: timezone }));

    // Gestione quota: limita a 18 trigger attivi per sicurezza
    const triggers = ScriptApp.getProjectTriggers();
    const singleHandlers = triggers.filter(t => t.getHandlerFunction && t.getHandlerFunction() === 'runSingleReminder');
    if (singleHandlers.length >= 18) {
      debugLog(`⚠️ Limite trigger raggiunto (${singleHandlers.length}). Salvo solo in properties, niente nuovo trigger.`);
    } else {
      // Se notifyTime è passato o troppo vicino, non creare trigger "at" impossibile
      if (notifyTime > now) {
        ScriptApp.newTrigger('runSingleReminder').timeBased().at(notifyTime).create();
      } else {
        // se è entro 1 minuto nel futuro/già passato, affida al handler la spedizione immediata
        ScriptApp.newTrigger('runSingleReminder').timeBased().after(60 * 1000).create();
      }
    }

    // Salva metadata in Properties per il riconoscimento nel handler
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty('single_reminders') || '[]';
    let arr = [];
    try { arr = JSON.parse(raw); } catch(e) { arr = []; }
    arr.push({ id: bookingId.toString(), notifyISO: notifyTime.toISOString() });
    props.setProperty('single_reminders', JSON.stringify(arr));

    debugLog(`⏰ Programmato reminder singolo per ${bookingId} alle ${Utilities.formatDate(notifyTime, timezone, 'dd/MM/yyyy HH:mm')}`);
    return { status: 'ok', message: 'Reminder programmato', data: { id: bookingId, notifyTime: notifyTime } };
  } catch (error) {
    debugLog(`❌ scheduleSingleReminder: ${error.message}`);
    return { status: 'error', message: error.message };
  }
}

/** Handler invocato dai trigger one‑shot: invia i reminder i cui notifyISO
 * sono entro ±2 minuti dall'orario corrente locale e li rimuove da Properties.
 */
function runSingleReminder() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const timezone = ss.getSpreadsheetTimeZone();
    const nowLocal = new Date(new Date().toLocaleString('en-US', { timeZone: timezone }));

    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty('single_reminders') || '[]';
    let arr = [];
    try { arr = JSON.parse(raw); } catch(e) { arr = []; }

    const keep = [];
    let sentCount = 0;
    arr.forEach(item => {
      try {
        const scheduled = new Date(item.notifyISO);
        const diffMin = Math.abs((nowLocal.getTime() - scheduled.getTime()) / 60000);
        if (diffMin <= 2) {
          const ok = sendReminderForBookingId(item.id);
          if (ok) sentCount++;
        } else {
          keep.push(item);
        }
      } catch(e) {
        keep.push(item);
      }
    });

    props.setProperty('single_reminders', JSON.stringify(keep));
    debugLog(`📨 runSingleReminder: inviati ${sentCount}, residui ${keep.length}`);
  } catch (error) {
    debugLog(`❌ runSingleReminder: ${error.message}`);
  }
}

// Invia reminder per l'ID prenotazione cercando la riga nel foglio
function sendReminderForBookingId(bookingId) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAMES.EVENTS);
    const timezone = ss.getSpreadsheetTimeZone();
    if (!sheet) return false;

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = findHeaderIndex(headers, ['ID', 'Codice', 'Code', 'BookingID']);
    const dateCol = headers.indexOf('Data');
    const timeCol = headers.indexOf('Ora');
    const eventCol = headers.indexOf('Evento');
    const nameCol = headers.indexOf('Nome');
    const notifiedCol = headers.indexOf('ReminderSent');
    if (idCol < 0 || dateCol < 0 || timeCol < 0) return false;

    let targetRowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(bookingId)) {
        targetRowIndex = i; break;
      }
    }
    if (targetRowIndex < 0) return false;
    const row = data[targetRowIndex];

    if (row[notifiedCol] === true || row[notifiedCol] === 'TRUE') {
      // Già notificato
      return true;
    }

    const eventDate = parseDateItalyOrISO(row[dateCol]);
    const parts = String(row[timeCol]).split(':');
    const h = parseInt(parts[0], 10); const m = parseInt(parts[1], 10);
    const eventDateTime = new Date(eventDate.getFullYear(), eventDate.getMonth(), eventDate.getDate(), h, m, 0);

    const eventType = row[eventCol] ? row[eventCol].toString().toLowerCase() : '';
    let duration = 90;
    if (eventType && DURATION_MAP[eventType]) duration = DURATION_MAP[eventType];

    const eventTypeName = getShortEventName(eventType) || eventType;
    const message = `⚽ Promemoria ${eventTypeName}!\n\n` +
                    `Tra 30 minuti inizia l'evento:\n` +
                    `📅 ${Utilities.formatDate(eventDate, timezone, 'dd/MM/yyyy')}\n` +
                    `🕐 ${String(row[timeCol])}\n` +
                    `👤 ${row[nameCol] || 'N/D'}\n` +
                    `⏱️ Durata: ${duration} minuti`;

    const sent = sendTelegramNotification(message, 'reminder');
    if (sent) {
      sheet.getRange(targetRowIndex + 1, notifiedCol + 1).setValue(true);
      SpreadsheetApp.flush();
      return true;
    }
    return false;
  } catch (error) {
    debugLog(`❌ sendReminderForBookingId: ${error.message}`);
    return false;
  }
}

function parseDateItalyOrISO(dateStr) {
  try {
    if (!dateStr) return null;
    const s = dateStr.toString();
    if (s.includes('/')) {
      const p = s.split('/');
      if (p.length === 3) return new Date(parseInt(p[2],10), parseInt(p[1],10)-1, parseInt(p[0],10));
    } else if (s.includes('-')) {
      const d = new Date(s);
      if (!isNaN(d.getTime())) return d;
    } else if (dateStr instanceof Date) {
      return dateStr;
    }
  } catch(e) {}
  return null;
}

function findHeaderIndex(headers, candidates) {
  for (const c of candidates) {
    const idx = headers.indexOf(c);
    if (idx >= 0) return idx;
  }
  return -1;
}

function createDailyArchiveTrigger() {
  try {
    deleteArchiveTriggers();
    ScriptApp.newTrigger('checkAndGenerateRecurringEvents').timeBased().everyDays(1).atHour(0).create();
    debugLog('✅ Trigger giornaliero consolidato creato per mezzanotte');
    return { status: 'ok', message: 'Trigger consolidato creato con successo' };
  } catch (error) {
    debugLog(`❌ Errore creazione trigger: ${error.message}`);
    return { status: 'error', message: `Errore: ${error.message}` };
  }
}

function deleteArchiveTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'checkAndGenerateRecurringEvents') {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
      }
    }
    debugLog(`🗑️ Rimossi ${deletedCount} trigger di archiviazione`);
    return { status: 'ok', deletedCount: deletedCount };
  } catch (error) {
    debugLog(`❌ Errore rimozione trigger: ${error.message}`);
    return { status: 'error', message: `Errore: ${error.message}` };
  }
}

function disableAllNotificationTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    const deletedTriggers = [];

    for (const trigger of triggers) {
      const handlerFunction = trigger.getHandlerFunction();
      if (handlerFunction === 'checkUpcomingEventsAndNotify' ||
          handlerFunction === 'sendDailySummaryToAdmin' ||
          handlerFunction === 'checkAndGenerateRecurringEvents') {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
        deletedTriggers.push(handlerFunction);
      }
    }

    debugLog(`🗑️ Rimossi ${deletedCount} trigger di notifica: ${deletedTriggers.join(', ')}`);

    return {
      status: 'ok',
      deletedCount: deletedCount,
      deletedTriggers: deletedTriggers,
      message: `Disattivati ${deletedCount} trigger di notifica`
    };

  } catch (error) {
    debugLog(`❌ Errore disattivazione trigger: ${error.message}`);
    return {
      status: 'error',
      message: `Errore disattivazione trigger: ${error.message}`
    };
  }
}

function getSimpleTriggerStatus() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    if (!triggers || triggers.length === 0) {
      return { active: false, message: 'Nessun trigger installato', triggers: [] };
    }
    const triggerList = triggers.map(tr => {
      let nextRun = null;
      if (tr.getHandlerFunction && tr.getTriggerSource) {
        if (tr.getTriggerSource() === ScriptApp.TriggerSource.CLOCK) {
          nextRun = '(prossima esecuzione non disponibile)';
        }
      }
      return {
        functionName: tr.getHandlerFunction(),
        type: tr.getEventType(),
        triggerSource: tr.getTriggerSource(),
        uniqueId: tr.getUniqueId ? tr.getUniqueId() : undefined,
        nextRun: nextRun
      };
    });
    return {
      active: true,
      count: triggers.length,
      triggers: triggerList
    };
  } catch (error) {
    debugLog('❌ Errore getSimpleTriggerStatus: ' + error.message);
    return { active: false, error: error.message };
  }
}

function getTriggerInfo() {
  return getSimpleTriggerStatus();
}

function testSheetConnection() {
  try {
    debugLog("🔍 Test connessione semplice...");
    const ss = SpreadsheetApp.openById(SHEET_ID);
    debugLog(`✅ Spreadsheet aperto: ${ss.getName()}`);
    const sheet = ss.getSheetByName(SHEET_NAMES.EVENTS);
    if (!sheet) {
      throw new Error(`Foglio "${SHEET_NAMES.EVENTS}" non trovato`);
    }

    const rows = sheet.getLastRow();
    const columns = sheet.getLastColumn();

    debugLog(`✅ Foglio eventi: ${rows} righe, ${columns} colonne`);

    return {
      success: true,
      message: "Connessione OK",
      data: {
        rows: rows,
        columns: columns
      }
    };

  } catch (error) {
    debugLog(`❌ Errore test connessione: ${error.message}`);
    return {
      success: false,
      message: `Errore: ${error.message}`
    };
  }
}

function safeTestConnection() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    return {
      success: true,
      message: "Sistema funzionante",
      details: {
        spreadsheet: "OK",
        mainSheet: sheet ? "OK" : "ERROR"
      }
    };

  } catch (error) {
    return {
      success: false,
      message: "Errore di connessione"
    };
  }
}
function cleanupDisabledTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let deleted = 0;
  for (const trigger of triggers) {
    const fn = trigger.getHandlerFunction();
    if (fn === 'runSingleReminder') {
      ScriptApp.deleteTrigger(trigger);
      deleted++;
    }
  }
  Logger.log(`🗑️ Eliminati ${deleted} trigger runSingleReminder`);
}
function resetAllTriggers() {
  // Elimina TUTTI i trigger senza eccezioni
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log(`🗑️ Eliminati ${triggers.length} trigger`);
  
  // Ricrea solo quelli necessari su Head
  ScriptApp.newTrigger('checkUpcomingEventsAndNotify').timeBased().everyMinutes(30).create();
  ScriptApp.newTrigger('sendDailySummaryToAdmin').timeBased().everyDays(1).atHour(7).create();
  ScriptApp.newTrigger('checkAndGenerateRecurringEvents').timeBased().everyDays(1).atHour(0).create();
  
  Logger.log('✅ 3 trigger ricreati su Head');
}
function testRiepilogoSingolo() {
  // Pulisci la cache del riepilogo per forzare un nuovo invio
  const cache = CacheService.getScriptCache();
  const today = Utilities.formatDate(new Date(), 
    SpreadsheetApp.openById(SHEET_ID).getSpreadsheetTimeZone(), 
    'yyyy-MM-dd');
  cache.remove(`daily_summary_${today}`);
  Logger.log('🧹 Cache riepilogo pulita');
  
  // Esegui il riepilogo
  const result = sendDailySummaryToAdmin();
  Logger.log('Risultato: ' + JSON.stringify(result));
}
