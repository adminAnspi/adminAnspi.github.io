// =====================================================
// UTILS.GS - FUNZIONI DI UTILITÀ E ARCHIVIAZIONE
// Versione Aggiornata con Fix Timezone
// =====================================================

/**
 * ✅ FUNZIONE PRINCIPALE - Controllo eventi e archiviazione automatica
 * Fix timezone: usa Europe/Rome invece di UTC
 */
function checkUpcomingEventsAndNotify() {
const lock = LockService.getScriptLock();
if (!lock.tryLock(5000)) {
  Logger.log('⏳ Archiviazione già in corso da altro processo, skip');
  return { archivedCount: 0, notificationsSent: 0 };
}
try {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const mainSheet = ss.getSheetByName(SHEET_NAMES.EVENTS);
  
  if (!mainSheet) {
    Logger.log('❌ Foglio Prenotazioni1 non trovato');
    return { archivedCount: 0, notificationsSent: 0 };
  }

  // ✅ EXIT RAPIDO: controlla se ci sono eventi da gestire
  const now = new Date();
  const timezone = ss.getSpreadsheetTimeZone();
  const nowItaly = new Date(now.toLocaleString('en-US', { timeZone: timezone }));
  const twoHoursLater = new Date(nowItaly.getTime() + 2 * 60 * 60000);
  
  const quickData = mainSheet.getDataRange().getValues();
  const quickHeaders = quickData[0].map(h => h ? h.toString().toLowerCase().trim() : '');
  const qDate = quickHeaders.indexOf('data');
  const qTime = quickHeaders.indexOf('ora');
  const qArchive = quickHeaders.indexOf('archiviato');
  
  const hasWork = quickData.slice(1).some(row => {
    if (!row[qDate] || !row[qTime]) return false;
    if (row[qArchive] === true || row[qArchive] === 'TRUE') return false;
    try {
      let eventDate = row[qDate];
      if (typeof eventDate === 'string' && eventDate.includes('/')) {
        const p = eventDate.split('/');
        eventDate = new Date(p[2], p[1]-1, p[0]);
      } else if (!(eventDate instanceof Date)) {
        eventDate = new Date(eventDate);
      }
      const timeParts = row[qTime].toString().split(':');
      const eventDateTime = new Date(
        eventDate.getFullYear(), eventDate.getMonth(), eventDate.getDate(),
        parseInt(timeParts[0]), parseInt(timeParts[1]), 0
      );
      // Ha lavoro se: evento passato (da archiviare) O entro 2 ore (reminder)
      const expiryTime = new Date(eventDateTime.getTime() + 90 * 60000 + 30 * 60000);
      return nowItaly >= expiryTime || eventDateTime <= twoHoursLater;
    } catch(e) { return false; }
  });
  
  if (!hasWork) {
    Logger.log('⚡ Nessun evento imminente o da archiviare - exit rapido');
    return { archivedCount: 0, notificationsSent: 0 };
  }
  
  Logger.log('✅ Eventi da gestire trovati - procedo con controllo completo');
    const archiveSheet = ss.getSheetByName(SHEET_NAMES.ARCHIVE);
    
    if (!mainSheet) {
      Logger.log('❌ Foglio Prenotazioni1 non trovato');
      return;
    }
    
    // Assicurati che il foglio Archivio esista
    if (!archiveSheet) {
      ss.insertSheet('Archivio_Prenotazioni');
      Logger.log('📦 Foglio Archivio_Prenotazioni creato');
    }
    
    // ⭐ FIX TIMEZONE: Usa l'ora corretta per l'Italia
    //const now = new Date();
    //const timezone = ss.getSpreadsheetTimeZone(); // Prende timezone dallo spreadsheet (Europe/Rome)
    //const nowItaly = new Date(now.toLocaleString('en-US', { timeZone: timezone }));
    
    Logger.log('=== INIZIO CONTROLLO EVENTI ===');
    Logger.log('🌍 Timezone spreadsheet: ' + timezone);
    Logger.log('🕐 Ora UTC: ' + now.toISOString());
    Logger.log('🇮🇹 Ora Italia: ' + nowItaly.toLocaleString('it-IT'));
    
    const data = mainSheet.getDataRange().getValues();
    
    // Trova gli indici delle colonne
 const headers = data[0].map(h => h ? h.toString().toLowerCase().trim() : '');
const dateColIndex = headers.indexOf('data');
const timeColIndex = headers.indexOf('ora');
const eventTypeColIndex = headers.indexOf('evento');
const statusColIndex = headers.indexOf('stato');
const notifiedColIndex = headers.indexOf('remindersent');
const nameColIndex = headers.indexOf('nome');
const archiveColIndex = headers.indexOf('archiviato');
    
    Logger.log('📊 Indici colonne trovati:');
    Logger.log(`   Data: ${dateColIndex}, Ora: ${timeColIndex}, Evento: ${eventTypeColIndex}`);
    Logger.log(`   Nome: ${nameColIndex}, Stato: ${statusColIndex}`);
    Logger.log(`   ReminderSent: ${notifiedColIndex}, Archiviato: ${archiveColIndex}`);
    
    // Array per tenere traccia delle righe da eliminare (per archiviazione)
    const rowsToDelete = [];
    let notificationsSent = 0;
    
    // Elabora ogni riga (salta l'intestazione)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Salta righe vuote
      if (!row[dateColIndex] || !row[timeColIndex]) {
        continue;
      }
      
      // Salta eventi già archiviati
      if (row[archiveColIndex] === true || row[archiveColIndex] === 'TRUE') {
        continue;
      }
      
      // Costruisci data e ora dell'evento
      let eventDate = row[dateColIndex];
      if (typeof eventDate === 'string') {
        // Supporta sia formato italiano (28/10/2025) che ISO (2025-10-28)
        if (eventDate.includes('/')) {
          const parts = eventDate.split('/');
          if (parts.length === 3) {
            eventDate = new Date(parts[2], parts[1] - 1, parts[0]);
          } else {
            continue;
          }
        } else if (eventDate.includes('-')) {
          eventDate = new Date(eventDate);
        } else {
          continue;
        }
      } else if (!(eventDate instanceof Date)) {
        continue;
      }
      
      // Estrai ora e minuti
      let eventHour, eventMinute;
      const timeStr = row[timeColIndex];
      if (typeof timeStr === 'string') {
        const timeParts = timeStr.split(':');
        if (timeParts.length >= 2) {
          eventHour = parseInt(timeParts[0]);
          eventMinute = parseInt(timeParts[1]);
        } else {
          continue;
        }
      } else {
        continue;
      }
      
      // Crea timestamp evento
      const eventDateTime = new Date(
        eventDate.getFullYear(),
        eventDate.getMonth(),
        eventDate.getDate(),
        eventHour,
        eventMinute,
        0
      );
      
      // Calcola durata usando DURATION_MAP basato sul tipo di evento
      const eventType = row[eventTypeColIndex] ? row[eventTypeColIndex].toString().toLowerCase() : '';
      let duration = 90; // Default calcetto
      
      if (eventType && DURATION_MAP[eventType]) {
        duration = DURATION_MAP[eventType];
      } else {
        Logger.log(`⚠️ Tipo evento non trovato (${eventType}), uso default: ${duration} min`);
      }
      
      // Calcola ora di fine evento
      const eventEndTime = new Date(eventDateTime.getTime() + duration * 60000);
      
      // Calcola tempo di scadenza (fine + 30 minuti)
      const expiryTime = new Date(eventEndTime.getTime() + 30 * 60000);
      
      // ===== ARCHIVIAZIONE IMMEDIATA (USA ORA ITALIANA) =====
      if (nowItaly >= expiryTime) {
        Logger.log(`🗄️ EVENTO DA ARCHIVIARE: ${row[nameColIndex]} (${eventType.toUpperCase()})`);
        Logger.log(`   📅 Data evento: ${eventDateTime.toLocaleString('it-IT')}`);
        Logger.log(`   ⏰ Fine evento: ${eventEndTime.toLocaleString('it-IT')}`);
        Logger.log(`   ⏳ Scadenza: ${expiryTime.toLocaleString('it-IT')}`);
        Logger.log(`   🇮🇹 Ora attuale: ${nowItaly.toLocaleString('it-IT')}`);
        Logger.log(`   ✅ Differenza: ${Math.round((nowItaly - expiryTime) / 60000)} minuti`);
        
        // Marca per l'archiviazione
        rowsToDelete.push({
          rowIndex: i,
          rowData: row,
          eventType: eventType,
          eventName: row[nameColIndex]
        });
        
        continue;
      }
      
      // ===== LOGICA NOTIFICHE =====
      
      // Calcola tempo per la notifica (30 minuti prima dell'inizio)
      const notifyTime = new Date(eventDateTime.getTime() - 30 * 60000);
      
      // Controlla se è il momento di inviare la notifica (USA ORA ITALIANA)
      if (nowItaly >= notifyTime && nowItaly < eventDateTime) {
        // Verifica se la notifica è già stata inviata
        if (row[notifiedColIndex] !== true && row[notifiedColIndex] !== 'TRUE') {
          Logger.log(`📢 NOTIFICA DA INVIARE: ${row[nameColIndex]} (${eventType.toUpperCase()})`);
          Logger.log(`   🕐 Evento: ${eventDateTime.toLocaleString('it-IT')}`);
          Logger.log(`   🇮🇹 Ora attuale: ${nowItaly.toLocaleString('it-IT')}`);
          
          // Invia notifica Telegram
          const eventTypeName = getShortEventName(eventType) || eventType;
          const message = `⚽ Promemoria ${eventTypeName}!\n\n` +
                         `Tra 30 minuti inizia l'evento:\n` +
                         `📅 ${Utilities.formatDate(eventDate, timezone, 'dd/MM/yyyy')}\n` +
                         `🕐 ${timeStr}\n` +
                         `👤 ${row[nameColIndex] || 'N/D'}\n` +
                         `⏱️ Durata: ${duration} minuti`;
          
          const sentSuccess = sendTelegramNotification(message, 'reminder');
          
          if (sentSuccess) {
            // Marca come notificato
            const notifiedCell = mainSheet.getRange(i + 1, notifiedColIndex + 1);
            notifiedCell.setValue(true);
            SpreadsheetApp.flush(); // Forza il salvataggio
            
            notificationsSent++;
            Logger.log(`✅ Notifica inviata e salvata per: ${row[nameColIndex]}`);
          } else {
            Logger.log(`❌ Errore invio notifica per: ${row[nameColIndex]}`);
          }
        }
      }

      // 🔁 Fallback: se il trigger scatta esattamente all'ora di inizio o entro 5 minuti dopo
      // e la notifica non è stata ancora inviata, manda un alert "Sta iniziando ora"
      else if (nowItaly >= eventDateTime && nowItaly < new Date(eventDateTime.getTime() + 5 * 60000)) {
        if (row[notifiedColIndex] !== true && row[notifiedColIndex] !== 'TRUE') {
          Logger.log(`⏰ NOTIFICA ULTIMO MINUTO: ${row[nameColIndex]} (${eventType.toUpperCase()})`);
          const eventTypeName = getShortEventName(eventType) || eventType;
          const message = `⏰ ${eventTypeName} inizia ORA!\n\n` +
                         `📅 ${Utilities.formatDate(eventDate, timezone, 'dd/MM/yyyy')}\n` +
                         `🕐 ${timeStr}\n` +
                         `👤 ${row[nameColIndex] || 'N/D'}`;
          const sentSuccess = sendTelegramNotification(message, 'reminder');
          if (sentSuccess) {
            const notifiedCell = mainSheet.getRange(i + 1, notifiedColIndex + 1);
            notifiedCell.setValue(true);
            SpreadsheetApp.flush();
            notificationsSent++;
          }
        }
      }
    }
    
    // ===== ESEGUI ARCHIVIAZIONE =====
    if (rowsToDelete.length > 0) {
      Logger.log(`\n📦 === INIZIO ARCHIVIAZIONE ===`);
      Logger.log(`📊 Eventi da archiviare: ${rowsToDelete.length}`);
      
      // Ordina per indice decrescente per eliminare dalla fine
      rowsToDelete.sort((a, b) => b.rowIndex - a.rowIndex);
      
      let successCount = 0;
      let eventTypesArchived = {};
      let archivedNames = [];
      
      const archiveSheetFinal = ss.getSheetByName('Archivio_Prenotazioni');
      
      rowsToDelete.forEach(item => {
        try {
          // Aggiungi al foglio Archivio
          archiveSheetFinal.appendRow(item.rowData);
          
          // Elimina dal foglio Prenotazioni (rowIndex + 1 perché le righe sono 1-indexed)
          mainSheet.deleteRow(item.rowIndex + 1);
          
          // Forza il salvataggio
          SpreadsheetApp.flush();
          
          // Conta per tipo
          const type = item.eventType || 'altro';
          eventTypesArchived[type] = (eventTypesArchived[type] || 0) + 1;
          archivedNames.push(`${item.eventName} (${type})`);
          
          successCount++;
          Logger.log(`✅ [${successCount}/${rowsToDelete.length}] Archiviato: ${item.eventName} (${type})`);
        } catch (error) {
          Logger.log(`❌ Errore archiviazione riga ${item.rowIndex}: ${error}`);
        }
      });
      
      // Invia notifica riepilogativa dettagliata
      if (successCount > 0) {
        let archiveMessage = `🗄️ Archiviazione automatica completata\n\n`;
        archiveMessage += `✅ ${successCount} evento/i archiviato/i\n\n`;
        
        // Dettaglio per tipo
        archiveMessage += `📊 Dettaglio per tipo:\n`;
        for (let type in eventTypesArchived) {
          const icon = EVENT_TYPES[type] ? EVENT_TYPES[type].icon : '📌';
          const name = EVENT_TYPES[type] ? EVENT_TYPES[type].name : type;
          archiveMessage += `${icon} ${name}: ${eventTypesArchived[type]}\n`;
        }
        
        archiveMessage += `\n🕐 Ora: ${nowItaly.toLocaleString('it-IT')}`;
        
        sendTelegramNotification(archiveMessage, 'default');
        
        Logger.log(`\n✅ === ARCHIVIAZIONE COMPLETATA ===`);
        Logger.log(`📈 Successo: ${successCount}/${rowsToDelete.length} eventi`);
        Logger.log(`📝 Eventi archiviati: ${archivedNames.join(', ')}`);
      }
    } else {
      Logger.log('\n📭 Nessun evento da archiviare');
    }
    
    Logger.log('\n=== FINE CONTROLLO EVENTI ===\n');
    
    return {
      archivedCount: rowsToDelete.length,
      notificationsSent: notificationsSent
    };
    
  } catch (error) {
    Logger.log('❌ ERRORE in checkUpcomingEventsAndNotify: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    
    try {
      sendTelegramNotification(`❌ ERRORE Sistema Archiviazione\n\n${error.toString()}`, 'urgent');
    } catch (e) {
      Logger.log('Impossibile inviare notifica errore: ' + e);
    }
    
    return {
      archivedCount: 0,
      notificationsSent: 0,
      error: error.toString()
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Funzione di test per checkUpcomingEventsAndNotify
 */
function testCheckAndArchive() {
  Logger.log('=== TEST ARCHIVIAZIONE CON DURATION_MAP ===');
  Logger.log('DURATION_MAP configurato:');
  for (let type in DURATION_MAP) {
    Logger.log(`  ${type}: ${DURATION_MAP[type]} minuti`);
  }
  Logger.log('');
  
  const result = checkUpcomingEventsAndNotify();
  
  Logger.log('');
  Logger.log('=== TEST COMPLETATO ===');
  Logger.log('Controlla il Log per vedere i risultati');
  Logger.log(`Archiviati: ${result.archivedCount}`);
  Logger.log(`Notifiche: ${result.notificationsSent}`);
  
  return {
    status: 'ok',
    data: result
  };
}
