// =====================
// FUNZIONE doGet - PUNTO DI INGRESSO WEB APP
// =====================
function doGet(e) {
  var queryString = e && e.parameter ? e.parameter : {};
  var template = HtmlService.createTemplateFromFile('index');
  template.queryString = queryString;
  return template.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ✅ Funzione include
function include(filename) {
  try {
    console.log('📁 Inclusione file: ' + filename);
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    console.error('❌ Errore inclusione file ' + filename + ': ' + error);
    return '<!-- Errore caricamento ' + filename + ': ' + error.message + ' -->';
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Gestione Campo')
    .addItem('Test Notifica Telegram', 'testTelegramNotification')
    .addItem('Invia Riepilogo Giornaliero', 'sendDailySummaryToAdmin')
    .addItem('Pulisci Cache Riepilogo', 'clearDailySummaryCache')
    .addItem('Setup Notifiche/Trigger', 'setupEventNotifications')
    .addItem('Test Connessione Fogli', 'testSheetConnection')
    .addItem('🔍 Debug Eventi Ricorrenti', 'debugRecurringDisplay')
    .addItem('🔧 Correggi Visualizzazione Ricorrenti', 'fixRecurringDisplay')
    .addItem('🚫 Disabilita Archiviazione Auto (2h)', 'disableAutoArchive')
    .addItem('✅ Riabilita Archiviazione Auto', 'enableAutoArchive')
    .addItem('Test Sistema Ricorrenze', 'checkAndGenerateRecurringEvents')
    .addItem('Pulisci Cache', 'clearCache')
    .addItem('🧹 Test Pulizia Cache Auto', 'cleanupCacheEvery12Hours')
    .addSeparator()
    .addItem('🔑 Reset Password Tutti Admin', 'resetAllAdminPasswords')
    .addItem('🔑 Imposta Password Temp Tutti', 'setTempPasswordForAll')
    .addItem('📬 Controlla Bounce Due Ore', 'checkTwoHourEmailBounces')
    .addToUi();
}

// =====================
// FUNZIONE PER INVIARE INFORMAZIONI PRENOTAZIONI DUE ORE
// =====================
function sendTwoHourBookingInfoToUser(email, name, masterEmail, extraInfo) {
  try {
    if (!isMasterAdmin(masterEmail)) {
      return { status: 'error', message: 'Solo l\'admin master può inviare queste informazioni' };
    }
    var trackingId = Utilities.getUuid();
    logTwoHourEmailEvent({
      id: trackingId,
      email: email,
      name: name,
      status: 'sent',
      subject: 'ℹ️ Informazione Importante - Prenotazioni di Due Ore',
      extraInfo: extraInfo || ''
    });
    sendTwoHourBookingInfo(email, name, extraInfo, trackingId, masterEmail);
    return { status: 'ok' };
  } catch (error) {
    console.error('Errore invio informazioni prenotazioni due ore:', error);
    return { status: 'error', message: error.toString() };
  }
}

// =====================
// LOG EMAIL DUE ORE
// =====================
function getEmailLogSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Email_Log');
  if (!sheet) {
    sheet = ss.insertSheet('Email_Log');
    sheet.getRange(1, 1, 1, 8).setValues([[
      'Timestamp', 'Tipo', 'Recipient', 'Name', 'Status', 'TrackingID', 'Subject', 'ExtraInfo'
    ]]);
  }
  return sheet;
}

function logTwoHourEmailEvent(event) {
  try {
    var sheet = getEmailLogSheet();
    var row = [
      new Date().toISOString(),
      'two_hour_info',
      (event.email || ''),
      (event.name || ''),
      (event.status || ''),
      (event.id || ''),
      (event.subject || ''),
      (event.extraInfo || '')
    ];
    sheet.appendRow(row);
  } catch (e) {
    console.error('Errore log email due ore:', e);
  }
}

// =====================
// CONTROLLO BOUNCE EMAIL DUE ORE
// =====================
function checkTwoHourEmailBounces() {
  try {
    var query = 'newer_than:14d (from:mailer-daemon OR from:postmaster) subject:"Informazione Importante - Prenotazioni di Due Ore"';
    var threads = GmailApp.search(query);
    var bounceCount = 0;
    threads.forEach(function(thread) {
      var messages = thread.getMessages();
      messages.forEach(function(msg) {
        var subject = msg.getSubject() || '';
        var body = msg.getPlainBody() || '';
        var toMatch = body.match(/Original-Recipient\s*:\s*([^\s\n]+)/i) || body.match(/Final-Recipient\s*:\s*RFC822;\s*([^\s\n]+)/i);
        var bouncedTo = toMatch ? toMatch[1] : '';
        var trackingId = '';
        var idMatchSubject = subject.match(/\[TRACK:([^\]]+)\]/);
        var idMatchBody = body.match(/\[TRACK:([^\]]+)\]/);
        if (idMatchSubject) {
          trackingId = idMatchSubject[1];
        } else if (idMatchBody) {
          trackingId = idMatchBody[1];
        }
        logTwoHourEmailEvent({
          id: trackingId,
          email: bouncedTo,
          name: '',
          status: 'bounced',
          subject: subject,
          extraInfo: 'Delivery failure'
        });
        bounceCount++;
      });
    });
    return { status: 'ok', message: 'Bounce rilevati: ' + bounceCount };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

// =====================
// FUNZIONI CACHE UNIFICATE
// =====================
function getEventsForDateUnified(dateString) {
  var UNIFIED_CACHE_KEY = dateString ? ('events_unified_cache_' + dateString) : CACHE_KEYS.UNIFIED_ALL;
  var cache = CacheService.getScriptCache();
  try {
    var cachedData = cache.get(UNIFIED_CACHE_KEY);
    if (cachedData) {
      var parsed = JSON.parse(cachedData);
      console.log('📦 Cache HIT unificata per ' + (dateString || 'ALL') + ': ' + parsed.length + ' eventi');
      return parsed;
    }
    console.log('🔍 Cache MISS unificata per ' + (dateString || 'ALL') + ', caricamento dati...');
    var allEvents = getBookings();
    if (!dateString) {
      cache.put(CACHE_KEYS.UNIFIED_ALL, JSON.stringify(allEvents), CACHE_DURATION);
      return allEvents;
    }
    var filteredEvents = allEvents.filter(function(event) {
      var eventDate = new Date(event.data).toLocaleDateString('it-IT');
      var targetDate = new Date(dateString).toLocaleDateString('it-IT');
      return eventDate === targetDate;
    });
    cache.put(UNIFIED_CACHE_KEY, JSON.stringify(filteredEvents), CACHE_DURATION);
    console.log('✅ Cache SALVATA unificata per ' + dateString + ': ' + filteredEvents.length + ' eventi');
    return filteredEvents;
  } catch (error) {
    console.error('❌ Errore cache unificata ' + dateString + ': ' + error.message);
    return getBookings();
  }
}

function getEventsForDatePublic(dateString) {
  console.log('👥 getEventsForDatePublic chiamata per: ' + dateString);
  return getEventsForDateUnified(dateString);
}

function getEventsForDate(dateString) {
  console.log('👨‍💼 getEventsForDate (admin) chiamata per: ' + dateString);
  return getEventsForDateUnified(dateString);
}

function invalidateAllCache() {
  try {
    console.log('🧹 === INVALIDAZIONE CACHE COMPLETA ===');
    var cache = CacheService.getScriptCache();
    var cacheKeys = [
      'all_bookings',
      'calendar_bookings_fast',
      CACHE_KEYS.ARCHIVED,
      CACHE_KEYS.EVENTS,
      CACHE_KEYS.USERS,
      CACHE_KEYS.CONFIG,
      CACHE_KEYS.BOOKINGS,
      'public_calendar_cache',
      'public_events_cache',
      CACHE_KEYS.UNIFIED_ALL
    ];
    var removedCount = 0;
    cacheKeys.forEach(function(key) {
      try {
        cache.remove(key);
        removedCount++;
        console.log('🗑️ Cache rimossa: ' + key);
      } catch (e) {
        console.log('⚠️ Errore rimozione cache ' + key + ': ' + e.message);
      }
    });
    var currentDate = new Date();
    for (var i = -7; i <= 30; i++) {
      var testDate = new Date(currentDate.getTime() + i * 24 * 60 * 60 * 1000);
      var dateString = testDate.toISOString().split('T')[0];
      var unifiedKey = 'events_unified_cache_' + dateString;
      try {
        cache.remove(unifiedKey);
        removedCount++;
      } catch (e) {}
    }
    console.log('✅ Cache invalidazione completata: ' + removedCount + ' chiavi rimosse');
    return removedCount;
  } catch (error) {
    console.error('❌ Errore invalidazione cache completa:', error.message);
    return 0;
  }
}

function getCacheStatus() {
  try {
    var cache = CacheService.getScriptCache();
    var testKeys = [
      'all_bookings',
      'calendar_bookings_fast',
      CACHE_KEYS.EVENTS,
      'public_calendar_cache',
      CACHE_KEYS.UNIFIED_ALL
    ];
    var cacheStatus = {};
    testKeys.forEach(function(key) {
      var cachedData = cache.get(key);
      cacheStatus[key] = cachedData ? 'HIT' : 'MISS';
    });
    return {
      status: 'ok',
      data: cacheStatus,
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

// =====================
// FUNZIONE serverRequest
// =====================
function serverRequest(action, data) {
  try {
    switch (action) {
      case 'getBookings':
      case 'getBookingsForCalendar':
      case 'refreshData':
        invalidateAllCache();
        return { status: 'ok', data: { bookings: getBookings() } };

      case 'createBooking':
        var createResult = createBooking(data);
        if (createResult.status === 'ok') invalidateAllCache();
        return createResult;

      case 'cancelBooking':
        var cancelResult = cancelBooking(data.bookingId);
        if (cancelResult.status === 'ok') invalidateAllCache();
        return cancelResult;

      case 'deleteBooking':
        var deleteResult = deleteBooking(data.bookingId);
        if (deleteResult.status === 'ok') invalidateAllCache();
        return deleteResult;

      case 'checkAndGenerateRecurringEvents':
        var recurringResult = checkAndGenerateRecurringEvents();
        invalidateAllCache();
        return recurringResult;

      case 'clearCache':
      case 'clearBookingsCache':
        var cleared = invalidateAllCache();
        return { status: 'ok', message: cleared + ' cache pulite', clearedCount: cleared };

      case 'invalidateAllCache':
        var invalidated = invalidateAllCache();
        return { status: 'ok', message: invalidated + ' cache invalidate', invalidatedCount: invalidated };

      case 'setupSmartRecurringTrigger': return setupSmartRecurringTrigger();
      case 'manualArchive': return archiveOldBookings();
      case 'testTelegramNotification': return testTelegramNotification();
      case 'diagnoseTelegram': return diagnoseTelegram();
      case 'sendDailySummaryToAdmin': return sendDailySummaryToAdmin();
      case 'adminLogin': return adminLogin(data);
      case 'createAdminPassword': return createAdminPassword(data);
      case 'requestPasswordReset': return requestPasswordReset(data);
      case 'performPasswordReset': return performPasswordReset(data);
      case 'updateBookingDetail': return updateBookingDetail(data.bookingId, data.field, data.value);
      case 'updateEventStatus': return updateEventStatus(data);
      case 'getAppVersion': return { status: 'ok', data: { version: APP_VERSION } };
      case 'getEventStats': return getEventStats();
      case 'getArchivedBookings': return getArchivedBookings();
      case 'getArchiveTriggerInfo': return getArchiveTriggerInfo();
      case 'testMarkDailySummary': return testMarkDailySummary();
      case 'markDailySummary': return markDailySummarySent(data.date);
      case 'clearDailySummaryCache': return clearDailySummaryCache();
      case 'fixDuplicateDailySummary': return fixDuplicateDailySummary();
      case 'debugRecurringDisplay': return debugRecurringDisplay();
      case 'fixRecurringDisplay': return fixRecurringDisplay();
      case 'disableAutoArchive': return disableAutoArchive();
      case 'enableAutoArchive': return enableAutoArchive();
      case 'getTriggerInfo': return getTriggerInfo();
      case 'setupEventNotifications': return setupEventNotifications();
      case 'disableAllNotificationTriggers': return disableAllNotificationTriggers();
      case 'getAllAdmins': return getAllAdmins();
      case 'addNewAdmin': return addNewAdmin(data);
      case 'removeAdmin': return removeAdmin(data);
      case 'resetAdminPassword': return resetAdminPassword(data);
      case 'initializeMissingAdmins': return initializeMissingAdmins();
      case 'testAdminAuth': return testAdminAuth();
      case 'ping': return { status: 'ok', message: 'pong', timestamp: new Date().toISOString() };
      case 'cleanupDuplicateEvents': return cleanupDuplicateRecurringEvents();
      case 'getCacheStatus': return getCacheStatus();
      case 'getConfiguration': return getConfiguration();
      case 'saveConfiguration': return saveConfiguration(data);

      case 'getDurationMap':
        try {
          if (typeof DURATION_MAP !== 'undefined') {
            return { status: 'ok', data: { durationMap: DURATION_MAP } };
          }
          return { status: 'error', message: 'DURATION_MAP non definito' };
        } catch (e) {
          return { status: 'error', message: e.message };
        }

      case 'getLoyaltyStats':
        try {
          var bookingsAll = getBookings();
          var archivedAll = getArchivedBookings();
          var useArchivedOnly = data && data.archiveOnly === true;
          var fromDateStr = data && data.fromDate ? String(data.fromDate) : '';
          var fromDate = fromDateStr ? new Date(fromDateStr) : null;
          var source = useArchivedOnly ? (archivedAll || []) : (bookingsAll || []).concat(archivedAll || []);
          var map = {};
          source.forEach(function(b) {
            if (!b || !b.email) return;
            if (useArchivedOnly && fromDate) {
              try {
                var d = new Date(String(b.data).split('T')[0]);
                if (isNaN(d.getTime()) || d < fromDate) return;
              } catch(e) { return; }
            }
            var em = String(b.email).toLowerCase().trim();
            if (!em) return;
            if (b.stato && String(b.stato).toLowerCase() === 'rifiutato') return;
            var rec = map[em] || { email: em, nome: b.nome || '', count: 0, last: b.data || '' };
            rec.count += 1;
            rec.nome = rec.nome || (b.nome || '');
            rec.last = b.data || rec.last;
            map[em] = rec;
          });
          var list = Object.keys(map).map(function(k) {
            var rec = map[k];
            return Object.assign({}, rec, { nextFreeEligible: (rec.count % 10 === 0 && rec.count > 0) });
          });
          return { status: 'ok', data: { users: list } };
        } catch(e) {
          return { status: 'error', message: e.message };
        }

      case 'getLoyaltyStatsUnified':
        try {
          debugLog('📊 Inizio getLoyaltyStatsUnified');
          var bookingsAll = getBookings();
          var archivedResult = getArchivedBookings();
          var archivedAll = [];
          if (archivedResult && archivedResult.status === 'ok' && archivedResult.data && archivedResult.data.bookings) {
            archivedAll = archivedResult.data.bookings;
          }
          var fromDateStr = data && data.fromDate ? String(data.fromDate) : '';
          var fromDate = fromDateStr ? new Date(fromDateStr) : null;
          var map = {};

          var isLoyaltyEligible = function(booking) {
            if (!booking || !booking.email || !booking.nome) return false;
            if (booking.stato && String(booking.stato).toLowerCase() === 'rifiutato') return false;
            var eventoType = booking.evento ? String(booking.evento).toLowerCase() : '';
            return eventoType === 'calcetto' || eventoType === 'pallavolo';
          };

          var normalizeName = function(name) {
            return String(name).trim().replace(/\s+/g, ' ').toLowerCase();
          };

          var capitalizeName = function(name) {
            return String(name).trim().replace(/\s+/g, ' ').split(' ').map(function(word) {
              if (word.length === 0) return word;
              return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
            }).join(' ');
          };

          var createUserKey = function(email, nome) {
            return String(email).toLowerCase().trim() + '|||' + normalizeName(nome);
          };

          var processBooking = function(b, applyDateFilter) {
            if (!isLoyaltyEligible(b)) return;
            if (applyDateFilter && fromDate) {
              try {
                var eventDate = new Date(String(b.data).split('T')[0]);
                if (isNaN(eventDate.getTime()) || eventDate < fromDate) return;
              } catch(e) { return; }
            }
            var userKey = createUserKey(b.email, b.nome);
            if (!map[userKey]) {
              map[userKey] = { email: String(b.email).toLowerCase().trim(), nome: capitalizeName(b.nome), count: 0, last: '' };
            }
            map[userKey].count += 1;
            if (!map[userKey].last || b.data > map[userKey].last) map[userKey].last = b.data || '';
          };

          if (bookingsAll && Array.isArray(bookingsAll)) bookingsAll.forEach(function(b) { processBooking(b, false); });
          if (archivedAll && Array.isArray(archivedAll)) archivedAll.forEach(function(b) { processBooking(b, true); });

          var list = [];
          for (var userKey in map) {
            if (map.hasOwnProperty(userKey)) {
              var user = map[userKey];
              user.nextFreeEligible = (user.count > 0 && user.count % 10 === 0);
              list.push(user);
            }
          }
          list.sort(function(a, b) { return b.count - a.count; });
          debugLog('✅ getLoyaltyStatsUnified completato: ' + list.length + ' utenti');
          return { status: 'ok', data: { users: list } };
        } catch(e) {
          return { status: 'error', message: e.message };
        }

      case 'testSheetConnection':
        try {
          return testSheetConnection();
        } catch (error) {
          return { success: false, message: 'Test fallito: ' + error.message };
        }

      case 'safeTestConnection': return safeTestConnection();

      case 'testSmartRecurring':
        try {
          var result = checkAndGenerateRecurringEvents();
          invalidateAllCache();
          return result;
        } catch (error) {
          return { status: 'error', message: error.message };
        }

      case 'setupSmartTrigger': return createDailyArchiveTrigger();
      case 'resetSmartSystem': return deleteArchiveTriggers();

      case 'manualRecurringCheck':
        var manualResult = checkAndGenerateRecurringEvents();
        invalidateAllCache();
        return manualResult;

      case 'cleanupDuplicateRecurringEvents': return cleanupDuplicateRecurringEvents();
      case 'debugNonArchivedEvents': return debugNonArchivedEvents();

      case 'getRecurringEvents':
        var recurringEvents = getBookings().filter(function(event) {
          return event.ricorrente === 'sì' || event.ricorrente === true;
        });
        return { status: 'ok', data: { events: recurringEvents } };

      case 'attivaTriggerCompleto': return attivaTriggerCompleto();
      case 'disattivaTriggerCompleto': return disattivaTriggerCompleto();
      case 'getTriggerStatusDetailed': return getTriggerStatusDetailed();
      case 'updateAllTriggersToLatestVersion': return updateAllTriggersToLatestVersion();

      case 'getDebugInfo':
        try {
          var sheet = getSheet();
          var sheetData = sheet.getDataRange().getValues();
          var headers = sheetData[0] || [];
          var recurringEventsDebug = sheetData.filter(function(row, index) {
            return index > 0 && row[headers.indexOf('ricorrente')] === true;
          });
          return {
            status: 'ok',
            data: {
              spreadsheetId: SHEET_ID,
              sheetNames: SHEET_NAMES,
              timestamp: new Date().toISOString(),
              totalRows: sheetData.length,
              recurringEventsCount: recurringEventsDebug.length,
              cacheStatus: getCacheStatus(),
              triggerStatus: getSimpleTriggerStatus()
            }
          };
        } catch (error) {
          return { status: 'error', message: error.message };
        }
        case 'checkInvisibleRecurrences':
        return checkInvisibleRecurrences(data.data, data.ora, data.campo);

      default:
        throw new Error('Azione non riconosciuta: ' + action);
    }
  } catch (error) {
    debugLog(error);
    return { status: 'error', message: error.message };
  }
}

// =====================
// FUNZIONI ADMIN
// =====================
function testAdminAuth() {
  try {
    var adminSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.USERS);
    if (!adminSheet) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Foglio AdminUsers non trovato' })).setMimeType(ContentService.MimeType.JSON);
    }
    var dataRange = adminSheet.getDataRange().getValues();
    if (dataRange.length < 2) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Nessun admin nel foglio', rows: dataRange.length })).setMimeType(ContentService.MimeType.JSON);
    }
    var headers = dataRange[0].map(function(h) { return h.toString().toLowerCase(); });
    var emailCol = headers.indexOf('email');
    if (emailCol === -1) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Colonna email non trovata', headers: headers })).setMimeType(ContentService.MimeType.JSON);
    }
    var testEmails = ADMIN_EMAILS_FOR_NOTIFICATIONS;
    var foundAdmins = [];
    for (var i = 1; i < dataRange.length; i++) {
      var rowEmail = dataRange[i][emailCol];
      if (rowEmail) {
        var cleanEmail = rowEmail.toString().toLowerCase().trim();
        if (testEmails.includes(cleanEmail)) {
          foundAdmins.push({ row: i, email: cleanEmail, fullRow: dataRange[i] });
        }
      }
    }
    var testResult = findAdminByEmail('pasqualem27@gmail.com');
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'Test completato',
      data: {
        totalRows: dataRange.length,
        headers: headers,
        emailColumn: emailCol,
        foundAdmins: foundAdmins,
        findAdminByEmailResult: testResult,
        allEmails: dataRange.slice(1).map(function(row) { return row[emailCol]; }).filter(function(email) { return email; })
      }
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString(), stack: error.stack })).setMimeType(ContentService.MimeType.JSON);
  }
}

function resetAllAdminPasswords() {
  try {
    var adminSheet = getAdminSheet();
    var data = adminSheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return h.toLowerCase(); });
    var emailCol = headers.indexOf('email');
    var passwordHashCol = headers.indexOf('passwordhash');
    var resetCount = 0;
    var resetEmails = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][emailCol]) {
        adminSheet.getRange(i + 1, passwordHashCol + 1).setValue('');
        resetCount++;
        resetEmails.push(data[i][emailCol].toString().toLowerCase().trim());
      }
    }
    return { status: 'ok', message: 'Reset completato! ' + resetCount + ' amministratori resettati.', resetCount: resetCount, emails: resetEmails };
  } catch (error) {
    return { status: 'error', message: 'Errore: ' + error.message };
  }
}

function setTempPasswordForAll() {
  try {
    var adminSheet = getAdminSheet();
    var data = adminSheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return h.toLowerCase(); });
    var emailCol = headers.indexOf('email');
    var passwordHashCol = headers.indexOf('passwordhash');
    var setCount = 0;
    var setEmails = [];
    var tempPassword = TEMP_ADMIN_PASSWORD;
    for (var i = 1; i < data.length; i++) {
      if (data[i][emailCol]) {
        var tempPasswordHash = hashPassword(tempPassword);
        adminSheet.getRange(i + 1, passwordHashCol + 1).setValue(tempPasswordHash);
        setCount++;
        setEmails.push(data[i][emailCol].toString().toLowerCase().trim());
      }
    }
    return { status: 'ok', message: 'Password temporanea impostata! ' + setCount + ' amministratori aggiornati. Password: ' + tempPassword, setCount: setCount, emails: setEmails, tempPassword: tempPassword };
  } catch (error) {
    return { status: 'error', message: 'Errore: ' + error.message };
  }
}

// =====================
// DEBUG EVENTI RICORRENTI
// =====================
function debugRecurringDisplay() {
  try {
    debugLog('🔍 DEBUG: Analisi sistema eventi ricorrenti');
    var now = new Date();
    var today = now.toISOString().split('T')[0];
    var allBookings = getBookings();
    debugLog('📊 Eventi totali nel foglio: ' + allBookings.length);
    var recurringEvents = allBookings.filter(function(b) {
      return b.ricorrente === true || b.ricorrente === 'TRUE' || b.ricorrente === 'true' || b.ricorrente === 'sì';
    });
    debugLog('🔄 Eventi ricorrenti trovati: ' + recurringEvents.length);
    var archivedRecurring = recurringEvents.filter(function(b) { return b.archiviato; });
    debugLog('📦 Eventi ricorrenti archiviati: ' + archivedRecurring.length);
    if (archivedRecurring.length > 0) {
      debugLog('⚠️ PROBLEMA: Eventi ricorrenti archiviati erroneamente!');
      archivedRecurring.forEach(function(event, index) {
        debugLog('   ' + (index + 1) + '. ' + event.nome + ' - ' + event.evento + ' - ' + event.data + ' - ID: ' + event.id);
      });
    }
    var todayEvents = allBookings.filter(function(b) {
      if (!b.data) return false;
      return b.data.split('T')[0] === today;
    });
    debugLog('📅 Eventi per oggi (' + today + '): ' + todayEvents.length);
    var calendarBookings = getBookingsForCalendar();
    var recurringInCalendar = calendarBookings.filter(function(b) {
      return b.ricorrente === true || b.ricorrente === 'TRUE' || b.ricorrente === 'true' || b.ricorrente === 'sì';
    });
    debugLog('📅 Eventi ricorrenti nel calendario: ' + recurringInCalendar.length);
    return {
      status: 'ok',
      data: { totalEvents: allBookings.length, recurringEvents: recurringEvents.length, archivedRecurring: archivedRecurring.length, todayEvents: todayEvents.length, recurringInCalendar: recurringInCalendar.length, hasProblem: archivedRecurring.length > 0 },
      message: 'Analisi completata: ' + recurringEvents.length + ' ricorrenti, ' + archivedRecurring.length + ' archiviati erroneamente'
    };
  } catch (error) {
    return { status: 'error', message: 'Errore: ' + error.message };
  }
}

function fixRecurringDisplay() {
  try {
    debugLog('🔧 CORREZIONE: Risoluzione problema visualizzazione eventi ricorrenti');
    invalidateAllCache();
    var allBookings = getBookings();
    var recurringEvents = allBookings.filter(function(b) {
      return b.ricorrente === true || b.ricorrente === 'TRUE' || b.ricorrente === 'true' || b.ricorrente === 'sì';
    });
    var archivedRecurring = recurringEvents.filter(function(b) { return b.archiviato; });
    var calendarResult = getBookingsForCalendar();
    var today = new Date().toISOString().split('T')[0];
    var todayEvents = allBookings.filter(function(b) {
      if (!b.data) return false;
      return b.data.split('T')[0] === today;
    });
    return {
      status: 'ok',
      data: { totalEvents: allBookings.length, recurringEvents: recurringEvents.length, archivedRecurring: archivedRecurring.length, todayEvents: todayEvents.length, calendarEvents: calendarResult.length },
      message: 'Correzione completata: ' + recurringEvents.length + ' ricorrenti, ' + archivedRecurring.length + ' archiviati, ' + todayEvents.length + ' per oggi'
    };
  } catch (error) {
    return { status: 'error', message: 'Errore: ' + error.message };
  }
}

// =====================
// GESTIONE TRIGGER
// =====================
function attivaTriggerCompleto() {
  try {
    var existingTriggers = ScriptApp.getProjectTriggers();
    existingTriggers.forEach(function(trigger) { ScriptApp.deleteTrigger(trigger); });
    ScriptApp.newTrigger('checkUpcomingEventsAndNotify').timeBased().everyMinutes(30).create();
    ScriptApp.newTrigger('sendDailySummaryToAdmin').timeBased().atHour(7).everyDays(1).create();
    try {
      sendTelegramNotification('✅ Sistema Trigger Attivato\n\n📊 2 trigger creati\n🔄 checkUpcomingEventsAndNotify (ogni 30 min)\n🔄 sendDailySummaryToAdmin (ore 7:00)\n\n⏰ ' + new Date().toLocaleString('it-IT'), 'default', { source: 'attivaTriggerCompleto' });
    } catch(e) {}
    return { status: 'ok', message: 'Trigger attivati con successo', data: { triggersCreated: 2, timestamp: new Date().toISOString() } };
  } catch (error) {
    try { sendTelegramNotification('❌ ERRORE Attivazione Trigger\n\n' + error.message, 'urgent', { source: 'attivaTriggerCompleto:error' }); } catch(e) {}
    return { status: 'error', message: 'Errore attivazione trigger: ' + error.message };
  }
}

function disattivaTriggerCompleto() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    var deletedCount = 0;
    var deletedTriggers = [];
    triggers.forEach(function(trigger) {
      deletedTriggers.push(trigger.getHandlerFunction());
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    });
    try {
      sendTelegramNotification('🚫 Sistema Trigger Disattivato\n\n📊 ' + deletedCount + ' trigger eliminati\n\n⏰ ' + new Date().toLocaleString('it-IT'), 'default', { source: 'disattivaTriggerCompleto' });
    } catch(e) {}
    return { status: 'ok', message: 'Tutti i trigger sono stati disattivati', data: { triggersDeleted: deletedCount, deletedTriggers: deletedTriggers, timestamp: new Date().toISOString() } };
  } catch (error) {
    try { sendTelegramNotification('❌ ERRORE Disattivazione Trigger\n\n' + error.message, 'urgent', { source: 'disattivaTriggerCompleto:error' }); } catch(e) {}
    return { status: 'error', message: 'Errore disattivazione trigger: ' + error.message };
  }
}

function getTriggerStatusDetailed() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    var triggerDetails = triggers.map(function(trigger) {
      var info = { funzione: trigger.getHandlerFunction(), tipo: trigger.getTriggerSource().toString(), evento: trigger.getEventType().toString() };
      if (info.funzione === 'checkUpcomingEventsAndNotify') {
        info.descrizione = 'Archiviazione automatica (ogni 30 min)'; info.icona = '🔄';
      } else if (info.funzione === 'sendDailySummaryToAdmin') {
        info.descrizione = 'Riepilogo giornaliero (ore 7:00)'; info.icona = '📧';
      } else {
        info.descrizione = info.funzione; info.icona = '⚙️';
      }
      return info;
    });
    return { status: 'ok', data: { active: triggers.length > 0, count: triggers.length, triggers: triggerDetails, timestamp: new Date().toISOString() } };
  } catch (error) {
    return { status: 'error', message: error.message, data: { active: false, count: 0, triggers: [] } };
  }
}

function updateAllTriggersToLatestVersion() {
  try {
    var existingTriggers = ScriptApp.getProjectTriggers();
    existingTriggers.forEach(function(trigger) { ScriptApp.deleteTrigger(trigger); });
    ScriptApp.newTrigger('checkUpcomingEventsAndNotify').timeBased().everyMinutes(30).create();
    ScriptApp.newTrigger('sendDailySummaryToAdmin').timeBased().atHour(7).everyDays(1).create();
    try { sendTelegramNotification('🔄 Trigger Aggiornati\n\n✅ 2 trigger ricreati con ultima versione\n\n⏰ ' + new Date().toLocaleString('it-IT'), 'default'); } catch(e) {}
    return { status: 'ok', message: 'Trigger aggiornati con successo', triggersCreated: 2 };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

// =====================
// CONTROLLO CONFLITTI
// =====================
function getAllEventsForDate(dateString, allBookings) {
  debugLog('🔍 getAllEventsForDate cercando data: "' + dateString + '"');
  var eventsForDay = allBookings.filter(function(b) {
    var bookingDatePart = b.data ? b.data.split('T')[0] : '';
    return bookingDatePart === dateString;
  });
  debugLog('📊 Eventi fisici trovati per ' + dateString + ': ' + eventsForDay.length);
  return eventsForDay.filter(function(event) { return event.stato !== 'in_attesa' && event.stato !== 'rifiutato'; });
}

function checkBookingConflicts(newBookingData, allEventsOnDate) {
  var newBookingStart = new Date(newBookingData.data.split('T')[0] + 'T' + newBookingData.ora + ':00');
  var newBookingDuration = DURATION_MAP[newBookingData.evento] || 60;
  var newBookingEnd = new Date(newBookingStart.getTime() + newBookingDuration * 60000);

  var isExclusiveEvent = function(evento) { return evento === 'compleanno' || evento === 'eventi'; };
  var getPhysicalFields = function(booking) {
    if (isExclusiveEvent(booking.evento)) return ['calcetto', 'pallavolo', 'exclusive'];
    if (booking.campo === 'entrambi') return ['calcetto', 'pallavolo'];
    if (booking.campo === 'calcetto') return ['calcetto'];
    if (booking.campo === 'pallavolo') return ['pallavolo'];
    return [];
  };

  var newBookingFields = getPhysicalFields(newBookingData);

  for (var i = 0; i < allEventsOnDate.length; i++) {
    var existingEvent = allEventsOnDate[i];
    if (newBookingData.id && existingEvent.id === newBookingData.id) continue;
    if (existingEvent.stato === 'rifiutato' || existingEvent.stato === 'in_attesa') continue;

    var existingStart = new Date(existingEvent.data.split('T')[0] + 'T' + existingEvent.ora + ':00');
    var existingDuration = DURATION_MAP[existingEvent.evento] || 60;
    var existingEnd = new Date(existingStart.getTime() + existingDuration * 60000);

    if (newBookingStart >= existingEnd || newBookingEnd <= existingStart) continue;

    if (isExclusiveEvent(newBookingData.evento) || isExclusiveEvent(existingEvent.evento)) {
      return { type: 'exclusive_conflict', conflictEvent: existingEvent, message: 'Il ' + newBookingData.data.split('T')[0] + ' è già prenotato per un evento speciale.', suggestions: { before: null, after: null } };
    }

    var existingBookingFields = getPhysicalFields(existingEvent);
    if (newBookingFields.some(function(field) { return existingBookingFields.includes(field); })) {
      return { type: 'field_conflict', conflictEvent: existingEvent };
    }
  }
  return null;
}

function generateTimeSuggestions(newBooking, allEventsOnDate, conflictEvent) {
  var newBookingDuration = DURATION_MAP[newBooking.evento] || 60;
  var conflictStart = new Date(conflictEvent.data.split('T')[0] + 'T' + conflictEvent.ora + ':00');
  var conflictDuration = DURATION_MAP[conflictEvent.evento] || 60;
  var conflictEnd = new Date(conflictStart.getTime() + conflictDuration * 60000);
  var suggestionBefore = new Date(conflictStart.getTime() - newBookingDuration * 60000);
  var suggestionAfter = new Date(conflictEnd.getTime());
  var spreadsheetTimeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var formatTime = function(date) { return Utilities.formatDate(date, spreadsheetTimeZone, "HH:mm"); };

  var isExclusiveEvent = function(evento) { return evento === 'compleanno' || evento === 'eventi'; };
  var getPhysicalFields = function(booking) {
    if (isExclusiveEvent(booking.evento)) return ['calcetto', 'pallavolo', 'exclusive'];
    if (booking.campo === 'entrambi') return ['calcetto', 'pallavolo'];
    if (booking.campo === 'calcetto') return ['calcetto'];
    if (booking.campo === 'pallavolo') return ['pallavolo'];
    return [];
  };

  var isSlotFree = function(slotStart) {
    var slotEnd = new Date(slotStart.getTime() + newBookingDuration * 60000);
    for (var j = 0; j < allEventsOnDate.length; j++) {
      var checkEvent = allEventsOnDate[j];
      if (checkEvent.id === conflictEvent.id) continue;
      if (newBooking.id && checkEvent.id === newBooking.id) continue;
      if (checkEvent.stato === 'rifiutato' || checkEvent.stato === 'in_attesa') continue;
      var checkStart = new Date(checkEvent.data.split('T')[0] + 'T' + checkEvent.ora + ':00');
      var checkDuration = DURATION_MAP[checkEvent.evento] || 60;
      var checkEnd = new Date(checkStart.getTime() + checkDuration * 60000);
      if ((slotStart < checkEnd) && (slotEnd > checkStart)) {
        var newBookingFields = getPhysicalFields(newBooking);
        var checkEventFields = getPhysicalFields(checkEvent);
        if (newBookingFields.some(function(field) { return checkEventFields.includes(field); })) return false;
      }
    }
    return true;
  };

  var isValidTime = function(date) {
    var hour = date.getHours();
    if (hour >= OPENING_HOURS.lunchStart && hour < OPENING_HOURS.lunchEnd) return false;
    if (hour < OPENING_HOURS.start || hour >= OPENING_HOURS.end) return false;
    if (date.getDay() === 0) return false;
    return true;
  };

  var beforeValid = isValidTime(suggestionBefore) && isSlotFree(suggestionBefore);
  var afterValid = isValidTime(suggestionAfter) && isSlotFree(suggestionAfter);

  return {
    before: beforeValid ? formatTime(suggestionBefore) : null,
    after: afterValid ? formatTime(suggestionAfter) : null
  };
}

// =====================
// CREAZIONE PRENOTAZIONI
// =====================
function createBooking(bookingData) {
  console.log("🔍 createBooking chiamata con:", JSON.stringify(bookingData, null, 2));

  if (!bookingData) {
    return { status: 'error', message: "Dati di prenotazione mancanti. Riprova o contatta l'amministratore." };
  }
  if (typeof bookingData !== 'object') {
    return { status: 'error', message: "Formato dati non valido. Riprova o contatta l'amministratore." };
  }
  if (!bookingData.data) {
    return { status: 'error', message: "Data dell'evento mancante. Seleziona una data valida." };
  }

  var lock = LockService.getScriptLock();
  var autoStato = '';

  try {
    var lockAcquired = lock.tryLock(10000);
    if (!lockAcquired) {
      return { status: 'error', message: "Sistema temporaneamente occupato. Riprova tra qualche secondo." };
    }

    // ===== VALIDAZIONE DATA E ORARIO =====
    var checkDateString, checkOraString, checkDateTime;
    try {
      checkDateString = bookingData.data.split('T')[0];
      checkOraString = bookingData.ora;
      if (!checkOraString) throw new Error("Ora mancante");
      checkDateTime = new Date(checkDateString + 'T' + checkOraString + ':00');
      if (isNaN(checkDateTime.getTime())) throw new Error("Data/ora non valida");
    } catch (dateError) {
      return { status: 'error', message: "Formato data/ora non valido: " + dateError.message };
    }

    var now = new Date();
    if (checkDateTime < now) {
      return { status: 'error', message: "❌ Non è possibile prenotare per una data/ora già passata." };
    }

    // ===== CARICAMENTO EVENTI E CONTROLLO CONFLITTI =====
    var allRelevantBookings;
    try {
      allRelevantBookings = getBookings();
    } catch (bookingsError) {
      return { status: 'error', message: "Errore nel caricamento delle prenotazioni esistenti. Riprova." };
    }

    var eventsOnDate = getAllEventsForDate(checkDateString, allRelevantBookings);
    var newBookingForCheck = {
      data: checkDateString,
      ora: checkOraString,
      campo: bookingData.campo,
      evento: bookingData.evento,
      id: bookingData.id || ''
    };

    var conflict = checkBookingConflicts(newBookingForCheck, eventsOnDate);
    if (conflict) {
      var conflictEvent = conflict.conflictEvent;
      var baseResponse = {
        status: 'conflict_suggestions',
        conflictDetails: {
          nome: conflictEvent && conflictEvent.nome,
          evento: conflictEvent && conflictEvent.evento,
          campo: conflictEvent && conflictEvent.campo,
          ora: conflictEvent && conflictEvent.ora
        },
        suggestions: { before: null, after: null }
      };
      if (conflict.type === 'exclusive_conflict') {
        return Object.assign({}, baseResponse, { message: conflict.message || 'Data occupata da evento speciale', suggestions: conflict.suggestions || baseResponse.suggestions });
      } else {
        var timeSuggestions = generateTimeSuggestions(newBookingForCheck, eventsOnDate, conflictEvent);
        return Object.assign({}, baseResponse, { message: 'Campo selezionato occupato per questo orario', suggestions: timeSuggestions || baseResponse.suggestions });
      }
    }

    // ===== INSERIMENTO NEL FOGLIO =====
    var sheet;
    try {
      sheet = getSheet();
      if (!sheet) throw new Error("Impossibile accedere al foglio prenotazioni");
    } catch (sheetError) {
      return { status: 'error', message: "Errore di accesso ai dati. Contatta l'amministratore." };
    }

    try {
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var isSpecialEvent = bookingData.evento === 'compleanno' || bookingData.evento === 'eventi';
      var isAdminBooking = bookingData.adminActionBy && bookingData.adminActionBy.trim() !== '';
      autoStato = isSpecialEvent ? (isAdminBooking ? 'approvato' : 'in_attesa') : '';

      var newRow = headers.map(function(header) {
        var key = header ? header.toLowerCase().trim().replace(/\s/g, '') : '';
        if (key === 'timestamp') return new Date();
        if (key === 'stato') return autoStato;
        if (key === 'adminactionby') return bookingData.adminActionBy || '';
        if (key === 'archiviato') return '';
        return bookingData[key] !== undefined ? bookingData[key] : '';
      });

      sheet.appendRow(newRow);
      SpreadsheetApp.flush();

      // Formattazione promo free
      try {
        if (bookingData.promofree === 'SI') {
          var lastRow = sheet.getLastRow();
          var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return h ? h.toString().toLowerCase().trim() : ''; });
          var promoColIndex = headerRow.indexOf('promofree');
          if (promoColIndex >= 0) {
            var cell = sheet.getRange(lastRow, promoColIndex + 1);
            cell.setValue('SI');
            cell.setBackground('#fffbe6');
            cell.setFontColor('#e67e22');
            cell.setNote('🎁 Prenotazione gratuita (11ª)');
          }
        }
      } catch (formatErr) {}

    } catch (insertError) {
      return { status: 'error', message: "Errore nel salvataggio dei dati. Riprova." };
    }

    // ===== PREPARA DATI PER NOTIFICHE =====
    var bookingDataForNotification = Object.assign({}, bookingData, { stato: autoStato || 'confermato' });

    // Verifica promo fedeltà
    try {
      var emailForCount = bookingData.email || '';
      var totalCount = 0;
      try {
        var allBookingsForPromo = getBookings();
        var archivedForPromo = getArchivedBookings();
        var arrForPromo = (allBookingsForPromo || []).concat(archivedForPromo || []);
        totalCount = arrForPromo.filter(function(b) {
          if (!b || !b.email) return false;
          if (String(b.email).toLowerCase().trim() !== String(emailForCount).toLowerCase().trim()) return false;
          if (b.stato && String(b.stato).toLowerCase() === 'rifiutato') return false;
          return true;
        }).length;
      } catch(e) {}
      if (((totalCount + 1) % 11 === 0) && totalCount >= 10) bookingDataForNotification.promofree = 'SI';
    } catch(e) {}

    // ===== NOTIFICA TELEGRAM =====
    try {
      var isSpecialForTelegram = bookingData.evento === 'compleanno' || bookingData.evento === 'eventi';
      var notificationType = isSpecialForTelegram ? 'special_event' : 'new_booking';
      var policyOk = shouldSendTelegramNotification(bookingDataForNotification, notificationType);
      console.log('🔔 Policy Telegram: ' + (policyOk ? 'CONSENTE' : 'BLOCCA'));
      sendBookingCreatedNotification(bookingDataForNotification, autoStato || 'confermato');
      try {
        var sched = scheduleSingleReminder(bookingDataForNotification.id, bookingDataForNotification.data, bookingDataForNotification.ora);
        console.log('⏰ scheduleSingleReminder: ' + sched.status + ' ' + (sched.message || ''));
      } catch(schedErr) {
        console.error('⚠️ Errore scheduleSingleReminder: ' + schedErr.message);
      }
    } catch(telegramError) {
      console.error('⚠️ Errore Telegram: ' + telegramError.message);
    }

    // ===== SUCCESSO =====
    console.log('✅ Prenotazione ' + bookingData.id + ' creata con successo');
    return { status: 'ok', data: { newBooking: bookingData, stato: autoStato } };

  } catch(error) {
    console.error('❌ ERRORE CRITICO in createBooking: ' + error.toString());
    console.error('📝 Stack: ' + error.stack);
    return { status: 'error', message: "Errore interno del sistema. L'amministratore è stato notificato.", error: error.message };

  } finally {
    try {
      lock.releaseLock();
      console.log("🔓 Lock rilasciato");
    } catch (lockError) {
      console.error("⚠️ Errore rilascio lock:", lockError);
    }
  }
}

// =====================
// CONFIGURAZIONE
// =====================
function getConfiguration() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Configurazione');
    if (!sheet) return { status: 'error', message: 'Foglio Configurazione non trovato' };
    var data = sheet.getDataRange().getValues();
    var config = {};
    for (var i = 1; i < data.length; i++) {
      var key = data[i][0] ? data[i][0].toString().trim() : '';
      var value = data[i][1] !== undefined ? data[i][1].toString().trim() : '';
      var desc = data[i][2] ? data[i][2].toString().trim() : '';
      if (key) config[key] = { value: value, description: desc };
    }
    return { status: 'ok', data: { config: config } };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

function saveConfiguration(data) {
  try {
    var key = data.key;
    var value = data.value;
    var adminEmail = data.adminEmail;
    if (!adminEmail || adminEmail.toLowerCase() !== MASTER_ADMIN_EMAIL.toLowerCase()) {
      return { status: 'error', message: 'Non autorizzato' };
    }
    if (!key || value === undefined) {
      return { status: 'error', message: 'Parametri mancanti' };
    }
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Configurazione');
    if (!sheet) return { status: 'error', message: 'Foglio non trovato' };
    var sheetData = sheet.getDataRange().getValues();
    for (var i = 1; i < sheetData.length; i++) {
      if (sheetData[i][0] && sheetData[i][0].toString().trim() === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        SpreadsheetApp.flush();
        invalidateAllCache();
        return { status: 'ok', message: 'Configurazione "' + key + '" aggiornata' };
      }
    }
    return { status: 'error', message: 'Chiave "' + key + '" non trovata' };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}
