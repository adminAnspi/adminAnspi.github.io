// =====================
// 7. FUNZIONI NOTIFICHE E TELEGRAM - VERSIONE CORRETTA v3.0
// =====================
function sendTelegramNotification(message, type, options) {
  options = options || {};

  if (!message || message.trim() === '') {
    var src = options.source || 'sorgente_sconosciuta';
    var t = type || 'default';
    debugLog('⚠️ Messaggio Telegram vuoto, skip (type=' + t + ', source=' + src + ')');
    return false;
  }

  if (!TELEGRAM_BOT_TOKEN) {
    debugLog('⚠️ Token Telegram non configurato');
    return false;
  }

  var emojiMap = {
    'new_booking': '📋',
    'cancellation': '❌',
    'special_event': '🎉',
    'reminder': '⏰',
    'urgent': '🚨',
    'default': '🔔',
    'same_day_alert': '⚡',
    'daily_summary': '📊'
  };

  var prefix = emojiMap[type] || emojiMap['default'];

  var htmlMessage = message
    .replace(/\*([^\*]+)\*/g, '<b>$1</b>')
    .replace(/_([^_]+)_/g, '<i>$1</i>')
    .replace(/`([^`]+)`/g, '<code>$1</code>');

  var text = prefix + ' <b>Campo Don Michele FIORE - ANSPI</b>\n\n' + htmlMessage + '\n\n⏰ ' + new Date().toLocaleString('it-IT') + '\n<i>by ver.' + APP_VERSION + '</i>';

  var url = 'https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/sendMessage';

  var chatIds = [];
  try {
    chatIds = getAdminChatIds();
  } catch(e) {
    chatIds = [];
  }
  if (!chatIds || !Array.isArray(chatIds) || chatIds.length === 0) {
    chatIds = (typeof TELEGRAM_CHAT_IDS !== 'undefined' && Array.isArray(TELEGRAM_CHAT_IDS) && TELEGRAM_CHAT_IDS.length > 0)
      ? TELEGRAM_CHAT_IDS
      : [TELEGRAM_MASTER_CHAT_ID];
  }

  debugLog('📱 Invio Telegram a ' + chatIds.length + ' destinatari...');

  var successCount = 0;
  var errors = [];

  for (var i = 0; i < chatIds.length; i++) {
    var chatId = chatIds[i];
    try {
      var response = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          chat_id: chatId,
          text: text,
          parse_mode: 'HTML',
          disable_notification: type === 'silent',
          disable_web_page_preview: true
        }),
        muteHttpExceptions: true
      });

      var result = JSON.parse(response.getContentText());

      if (result.ok) {
        successCount++;
        debugLog('✅ Telegram inviato a ' + chatId);
      } else {
        errors.push('Chat ' + chatId + ': ' + result.description);
        debugLog('❌ Errore Telegram ' + chatId + ': ' + result.description);
      }
    } catch (e) {
      errors.push('Chat ' + chatId + ': ' + e.message);
      debugLog('❌ Errore invio Telegram a ' + chatId + ': ' + e.message);
    }
  }

  var success = successCount > 0;
  debugLog('📊 Telegram: ' + successCount + '/' + chatIds.length + ' inviati con successo');
  if (errors.length > 0) {
    debugLog('⚠️ Errori: ' + errors.join(', '));
  }

  return success;
}

function testTelegramConnection() {
  debugLog('🧪 Test connessione Telegram...');
  var testMessage = '🧪 <b>TEST SISTEMA TELEGRAM</b>\n\n' +
    '✅ Connessione funzionante\n' +
    '📅 Data: ' + new Date().toLocaleDateString('it-IT') + '\n' +
    '🕐 Ora: ' + new Date().toLocaleTimeString('it-IT');
  var result = sendTelegramNotification(testMessage, 'default', { source: 'testTelegramConnection' });
  return {
    status: result ? 'ok' : 'error',
    message: result ? 'Telegram funzionante' : 'Errore invio Telegram',
    sent: result
  };
}

function diagnoseTelegram() {
  try {
    var diagnostics = {
      timestamp: new Date().toISOString(),
      tokenPresent: !!TELEGRAM_BOT_TOKEN,
      getMe: null,
      chatTests: []
    };

    if (!TELEGRAM_BOT_TOKEN) {
      debugLog('❌ Telegram: token mancante');
      return { status: 'error', message: 'Token Telegram non configurato', data: diagnostics };
    }

    try {
      var meResp = UrlFetchApp.fetch('https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/getMe', {
        method: 'get',
        muteHttpExceptions: true
      });
      var meJson = JSON.parse(meResp.getContentText());
      diagnostics.getMe = meJson;
      if (!meJson.ok) {
        debugLog('❌ getMe fallito: ' + (meJson.description || 'unknown'));
      } else {
        debugLog('✅ getMe OK: @' + (meJson.result && meJson.result.username));
      }
    } catch (e) {
      diagnostics.getMe = { ok: false, error: e.message };
      debugLog('❌ Errore getMe: ' + e.message);
    }

    var chatIds = [];
    try {
      chatIds = getAdminChatIds();
    } catch(e) {
      chatIds = [];
    }
    if (!chatIds || !Array.isArray(chatIds) || chatIds.length === 0) {
      chatIds = (typeof TELEGRAM_CHAT_IDS !== 'undefined' && Array.isArray(TELEGRAM_CHAT_IDS) && TELEGRAM_CHAT_IDS.length > 0)
        ? TELEGRAM_CHAT_IDS
        : [TELEGRAM_MASTER_CHAT_ID];
    }

    for (var i = 0; i < chatIds.length; i++) {
      var chatId = chatIds[i];
      var chatDiag = { chatId: chatId, getChat: null, sendTest: null };

      try {
        var chatResp = UrlFetchApp.fetch('https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/getChat', {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify({ chat_id: chatId }),
          muteHttpExceptions: true
        });
        var chatJson = JSON.parse(chatResp.getContentText());
        chatDiag.getChat = chatJson;
        if (chatJson.ok) {
          debugLog('✅ getChat OK per ' + chatId);
        } else {
          debugLog('❌ getChat fallito per ' + chatId + ': ' + (chatJson.description || 'unknown'));
        }
      } catch (e) {
        chatDiag.getChat = { ok: false, error: e.message };
        debugLog('❌ Errore getChat per ' + chatId + ': ' + e.message);
      }

      try {
        var pingResp = UrlFetchApp.fetch('https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/sendMessage', {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify({
            chat_id: chatId,
            text: '🔎 Diagnosi Telegram: ping di prova (' + new Date().toLocaleString('it-IT') + ')',
            parse_mode: 'HTML',
            disable_notification: true,
            disable_web_page_preview: true
          }),
          muteHttpExceptions: true
        });
        var pingJson = JSON.parse(pingResp.getContentText());
        chatDiag.sendTest = pingJson;
        if (pingJson.ok) {
          debugLog('✅ Ping inviato a ' + chatId);
        } else {
          debugLog('❌ Ping fallito per ' + chatId + ': ' + (pingJson.description || 'unknown'));
        }
      } catch (e) {
        chatDiag.sendTest = { ok: false, error: e.message };
        debugLog('❌ Errore ping per ' + chatId + ': ' + e.message);
      }

      diagnostics.chatTests.push(chatDiag);
    }

    var anySuccess = false;
    for (var j = 0; j < diagnostics.chatTests.length; j++) {
      if (diagnostics.chatTests[j].sendTest && diagnostics.chatTests[j].sendTest.ok) {
        anySuccess = true;
        break;
      }
    }

    return {
      status: anySuccess ? 'ok' : 'error',
      message: anySuccess ? 'Diagnosi OK: almeno un invio riuscito' : 'Diagnosi fallita: nessun invio riuscito',
      data: diagnostics
    };
  } catch (error) {
    debugLog('❌ Errore diagnosi Telegram: ' + error.message);
    return { status: 'error', message: error.message };
  }
}

function formatDateForEmail(dateInput) {
  try {
    if (!dateInput) return 'N/D';
    if (dateInput instanceof Date) {
      return Utilities.formatDate(dateInput, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }
    var str = dateInput.toString();
    var isoPart = str.split('T')[0];
    var m = isoPart.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) {
      return m[3] + '/' + m[2] + '/' + m[1];
    }
    var d = new Date(str);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }
    return str;
  } catch (e) {
    debugLog('❌ Errore formatDateForEmail: ' + e.message);
    try { return dateInput.toString(); } catch (e2) { return 'N/D'; }
  }
}

function sendSameDayBookingAlert(bookingData) {
  try {
    var timezone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    var todayLocal = Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd');

    var bookingDateStr = bookingData && bookingData.data ? bookingData.data.toString() : '';
    var bookingDateIso = bookingDateStr.split('T')[0];

    if (bookingDateIso.indexOf('/') !== -1) {
      var parts = bookingDateIso.split('/');
      if (parts.length === 3) {
        bookingDateIso = parts[2] + '-' + (parts[1].length < 2 ? '0' + parts[1] : parts[1]) + '-' + (parts[0].length < 2 ? '0' + parts[0] : parts[0]);
      }
    }

    if (bookingDateIso !== todayLocal) {
      debugLog('📅 Prenotazione non per oggi (' + bookingDateIso + ' vs ' + todayLocal + ') - nessun alert');
      return false;
    }

    debugLog('⚡ INVIO ALERT STESSO GIORNO: Nuova prenotazione dopo avviso giornaliero');

    var eventType = bookingData.evento || '';
    var campoType = bookingData.campo || '';
    var eventIcon = (EVENT_TYPES[eventType] && EVENT_TYPES[eventType].icon) || '📅';
    var eventName = (EVENT_TYPES[eventType] ? EVENT_TYPES[eventType].icon + ' ' + EVENT_TYPES[eventType].name : eventType);
    var campoMap = { calcetto: 'Calcetto', pallavolo: 'Pallavolo', entrambi: 'Entrambi', 'admin-managed': 'Da definire' };
    var campoName = campoMap[campoType] || campoType;

    var bookingDateForMsg = bookingDateIso || (bookingData.data ? bookingData.data.toString().split('T')[0] : todayLocal);

    var message = '⚡ <b>NUOVA PRENOTAZIONE STESSO GIORNO!</b>\n\n';
    message += '📅 <b>Data:</b> ' + formatDateForEmail(bookingDateForMsg) + '\n';
    message += '🕐 <b>Ora:</b> ' + bookingData.ora + '\n';
    message += '👤 <b>Cliente:</b> ' + bookingData.nome + '\n';
    message += '📧 <b>Email:</b> ' + (bookingData.email || 'N/D') + '\n';
    message += '📞 <b>Telefono:</b> ' + (bookingData.telefono || 'N/D') + '\n';
    message += eventIcon + ' <b>Evento:</b> ' + eventName + '\n';
    message += '🏟️ <b>Campo:</b> ' + campoName + '\n';
    message += '🔐 <b>Codice:</b> ' + bookingData.id + '\n';

    if (bookingData.note && bookingData.note.trim()) {
      message += '📝 <b>Note:</b>\n' + bookingData.note + '\n';
    }
    if (bookingData.evento === 'compleanno' || bookingData.evento === 'eventi') {
      message += '\n🎉 <b>EVENTO SPECIALE</b> - Richiede attenzione immediata!\n';
      if (bookingData.stato === 'in_attesa') {
        message += '⏳ <b>Stato:</b> In attesa di approvazione\n';
      }
    }
    if (bookingData.adminActionBy) {
      message += '\n👨‍💼 <b>Creata da:</b> ' + bookingData.adminActionBy + '\n';
    }
    message += '\n📱 <b>Alert automatico</b> - Avviso giornaliero già inviato alle 7:00';

    sendTelegramNotification(message, 'same_day_alert', { source: 'sendSameDayBookingAlert' });
    debugLog('✅ Alert stesso giorno inviato per ' + bookingData.id);
    return true;

  } catch (error) {
    debugLog('❌ Errore invio alert stesso giorno: ' + error.message);
    return false;
  }
}

function sendUrgentNotification(message, title, type, options) {
  var fullMessage = title ? '<b>' + title + '</b>\n\n' + message : message;
  sendTelegramNotification(fullMessage, type || 'urgent', Object.assign({}, options || {}, { source: 'sendUrgentNotification' }));
  debugLog('📱 Notifica urgente inviata: ' + (title || '(senza titolo)'));
}

// =====================
// NOTIFICA CREAZIONE PRENOTAZIONE — solo Telegram, email gestita da Code.gs
// =====================
function sendBookingCreatedNotification(bookingData, stato) {
  try {
    if (!bookingData || typeof bookingData !== 'object') {
      debugLog('❌ sendBookingCreatedNotification: bookingData non valido');
      return false;
    }

    var registrationTimestamp = new Date();
    var isSpecialEvent = bookingData.evento === 'compleanno' || bookingData.evento === 'eventi';
    var type = isSpecialEvent ? 'special_event' : 'new_booking';

    debugLog('📋 === NOTIFICA NUOVA REGISTRAZIONE ===');
    debugLog('⏰ Timestamp registrazione: ' + registrationTimestamp.toLocaleString('it-IT'));
    debugLog('📅 Data evento: ' + bookingData.data);
    debugLog('🎯 Tipo: ' + type);

    if (!shouldSendTelegramNotification(bookingData, type)) {
      debugLog('🔕 Notifica prenotazione saltata per policy');
      return false;
    }

    var eventType = bookingData.evento || '';
    var campoType = bookingData.campo || '';
    var eventIcon = (EVENT_TYPES[eventType] && EVENT_TYPES[eventType].icon) || '📅';
    var eventName = (EVENT_TYPES[eventType] ? EVENT_TYPES[eventType].icon + ' ' + EVENT_TYPES[eventType].name : eventType);
    var campoMap = { calcetto: 'Calcetto', pallavolo: 'Pallavolo', entrambi: 'Entrambi', 'admin-managed': 'Da definire' };
    var campoName = campoMap[campoType] || campoType;
    var bookingDate = (bookingData.data || '').toString().split('T')[0];

    var message = isSpecialEvent
      ? '🎉 <b>Nuova Richiesta Evento Speciale</b>\n\n'
      : eventIcon + ' <b>Nuova Prenotazione Registrata</b>\n\n';

    message += '📅 <b>Data Evento:</b> ' + formatDateForEmail(bookingDate) + '\n';
    message += '🕐 <b>Ora:</b> ' + bookingData.ora + '\n';
    message += '👤 <b>Cliente:</b> ' + bookingData.nome + '\n';
    message += '📧 <b>Email:</b> ' + (bookingData.email || 'N/D') + '\n';
    message += '📞 <b>Telefono:</b> ' + (bookingData.telefono || 'N/D') + '\n';
    message += eventIcon + ' <b>Evento:</b> ' + eventName + '\n';
    message += '🏟️ <b>Campo:</b> ' + campoName + '\n';
    message += '🔐 <b>Codice:</b> ' + bookingData.id + '\n';

    if (bookingData.note && bookingData.note.trim()) {
      message += '📝 <b>Note:</b>\n' + bookingData.note + '\n';
    }
    if (isSpecialEvent) {
      message += '\n🎯 <b>Stato:</b> ' + (stato || bookingData.stato || 'in_attesa') + '\n';
    }
    if (bookingData.adminActionBy) {
      message += '\n👨‍💼 <b>Creata da:</b> ' + bookingData.adminActionBy + '\n';
    }
    message += '\n⏰ <b>Registrata:</b> ' + registrationTimestamp.toLocaleString('it-IT');

    // ✅ SOLO TELEGRAM — email gestita da Code.gs
    var sent = sendTelegramNotification(message, type, { source: 'sendBookingCreatedNotification' });

    if (sent) {
      debugLog('✅ Notifica registrazione inviata (type=' + type + ') per ' + bookingData.id);
      try {
        var alertSent = sendSameDayBookingAlert(bookingData);
        if (alertSent) {
          debugLog('⚡ Alert urgenza stesso giorno inviato per ' + bookingData.id);
        }
      } catch (e) {
        debugLog('⚠️ Alert stesso giorno non inviato: ' + e.message);
      }
    } else {
      debugLog('⚠️ Notifica registrazione non inviata');
    }

    return sent;

  } catch (error) {
    debugLog('❌ Errore sendBookingCreatedNotification: ' + error.message);
    return false;
  }
}

// =====================
// EMAIL
// =====================
function sendEmailNotification(to, subject, htmlBody, bcc) {
  try {
    GmailApp.sendEmail(to, subject,
      'Hai ricevuto una conferma di prenotazione dal Campo Don Michele FIORE - ANSPI. Apri questa email in un client che supporta HTML per visualizzarla correttamente.',
      {
        htmlBody: htmlBody,
        bcc: bcc || '',
        name: 'Campo Don Michele FIORE - ANSPI',
        replyTo: 'prenotazionecampoanspi@gmail.com'
      }
    );
    debugLog('📧 Email inviata a ' + to);
  } catch (e) {
    debugLog('❌ Errore invio email: ' + e.message);
  }
}

function sendTwoHourBookingInfo(to, nome, extraInfo, trackingId, adminEmail) {
  try {
    var msg = 'Informazione Prenotazioni\nCliente: ' + (nome || '') +
      (trackingId ? ('\nTracking: ' + trackingId) : '') +
      (extraInfo && extraInfo.trim() ? ('\nMessaggio admin:\n' + extraInfo) : '');
    sendTelegramNotification(msg, 'urgent');
    return true;
  } catch (e) {
    return false;
  }
}

function testTelegramNotification() {
  try {
    var testMessage = '🧪 TEST SISTEMA NOTIFICHE\n\nData/ora: ' + new Date().toLocaleString('it-IT') +
      '\nSistema Telegram: FUNZIONANTE\n\n📋 Tipi di notifiche attive:\n' +
      '• 📅 Riepilogo giornaliero (ore 7:00)\n' +
      '• ⏰ Promemoria eventi (ogni 30 min)\n' +
      '• 📋 Nuove prenotazioni\n' +
      '• ⚡ Alert stesso giorno';
    debugLog('🧪 Invio messaggio di test Telegram...');
    sendTelegramNotification(testMessage, 'urgent', { source: 'testTelegramNotification' });
    debugLog('✅ Test Telegram completato');
    return { status: 'ok', message: 'Messaggio di test inviato su Telegram' };
  } catch (error) {
    debugLog('❌ Errore test Telegram: ' + error.message);
    return { status: 'error', message: 'Errore test Telegram: ' + error.message };
  }
}

function checkIfDailySummarySent(dateString) {
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'daily_summary_' + dateString;
    var result = cache.get(cacheKey) !== null;
    debugLog('🗃️ Cache riepilogo per ' + dateString + ': ' + (result ? 'TROVATO' : 'NON TROVATO'));
    return result;
  } catch (error) {
    debugLog('❌ Errore controllo cache riepilogo: ' + error.message);
    return false;
  }
}

function markDailySummarySent(dateString) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'daily_summary_' + dateString;
  cache.put(cacheKey, 'sent', 86400);
}

// =====================
// EMAIL TEMPLATE
// =====================
function createEmailTemplate(bookingData, stato) {
  stato = stato || '';
  var isExclusiveEvent = bookingData.evento === 'compleanno' || bookingData.evento === 'eventi';
  var eventIconMap = { 'calcetto': '&#9917;', 'pallavolo': '&#127952;', 'compleanno': '&#127881;', 'eventi': '&#127914;' };
  var eventIcon = eventIconMap[bookingData.evento] || '&#128197;';
  var eventName = getEventName(bookingData.evento);
  var campoName = getCampoName(bookingData.campo);

  var statusMessage = (isExclusiveEvent && stato === 'in_attesa')
    ? '<div class="status-badge status-pending">&#9203; In Valutazione</div><p>La tua richiesta per evento speciale è stata ricevuta e sarà valutata dall\'amministratore.</p>'
    : '<div class="status-badge status-confirmed">&#10003; Confermata (salvo diversa indicazione dell\'admin)</div><p>La tua prenotazione è stata <strong>confermata</strong> con successo!</p>';

  var eventBannerBg = isExclusiveEvent
    ? 'linear-gradient(135deg, #fd79a8 0%, #fdcb6e 100%)'
    : 'linear-gradient(135deg, #74b9ff 0%, #0984e3 100%)';

  var noteRow = bookingData.note
    ? '<div class="detail-row"><div class="detail-label">&#128221; Note</div><div class="detail-value" style="white-space: pre-wrap; text-align: left;">' + bookingData.note + '</div></div>'
    : '';

  var infoBox = (isExclusiveEvent && stato === 'in_attesa')
    ? '<div class="info-box"><strong>&#8505;&#65039; Informazioni Importanti</strong><br>Gli eventi speciali richiedono una valutazione amministrativa. Riceverai conferma definitiva a breve.</div>'
    : (bookingData.evento === 'eventi' ? '' : '<div class="info-box"><strong>&#8505;&#65039; Cosa Portare</strong><br>&#8226; Scarpe adatte al campo<br>&#8226; Abbigliamento sportivo<br>&#8226; Acqua per idratarsi</div>');

  var closingText = isExclusiveEvent
    ? 'Per qualsiasi domanda o modifica, contatta l\'amministratore.'
    : 'Per cancellare la prenotazione, usa il codice sopra nell\'apposita sezione del sito.';

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">'
    + '<meta name="viewport" content="width=device-width, initial-scale=1.0">'
    + '<title>Conferma Prenotazione Campo ANSPI</title>'
    + '<style>'
    + '* { margin: 0; padding: 0; box-sizing: border-box; }'
    + 'body { font-family: Segoe UI, Tahoma, Geneva, Verdana, sans-serif; background: linear-gradient(135deg, #a8e6cf 0%, #7fcdcd 30%, #74b9ff 60%, #81ecec 100%); padding: 20px; line-height: 1.6; }'
    + '.email-container { max-width: 600px; margin: 0 auto; background: white; border-radius: 20px; overflow: hidden; box-shadow: 0 20px 60px rgba(116,185,255,0.15); }'
    + '.header { background: linear-gradient(135deg, #6a89cc 0%, #b8e994 100%); color: white; padding: 40px 30px; text-align: center; }'
    + '.header h1 { font-size: 28px; margin-bottom: 10px; font-weight: 700; }'
    + '.header .subtitle { font-size: 16px; opacity: 0.9; }'
    + '.event-banner { background: ' + eventBannerBg + '; color: white; padding: 25px; text-align: center; font-size: 24px; font-weight: 700; }'
    + '.content { padding: 40px 30px; }'
    + '.booking-details { background: #f8f9fa; border-radius: 15px; padding: 25px; margin: 20px 0; }'
    + '.detail-row { display: flex; justify-content: space-between; align-items: center; padding: 12px 0; border-bottom: 1px solid #e9ecef; }'
    + '.detail-row:last-child { border-bottom: none; }'
    + '.detail-label { font-weight: 600; color: #495057; display: flex; align-items: center; gap: 8px; }'
    + '.detail-value { font-weight: 700; color: #2d3436; text-align: right; }'
    + '.status-badge { display: inline-block; padding: 8px 16px; border-radius: 25px; font-size: 14px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; margin: 15px 0; }'
    + '.status-confirmed { background: linear-gradient(135deg, #00cec9 0%, #55efc4 100%); color: white; }'
    + '.status-pending { background: linear-gradient(135deg, #fdcb6e 0%, #e17055 100%); color: white; }'
    + '.info-box { background: rgba(116,185,255,0.1); border-left: 4px solid #74b9ff; padding: 20px; margin: 25px 0; border-radius: 0 10px 10px 0; }'
    + '.footer { background: #2d3436; color: white; padding: 30px; text-align: center; }'
    + '.footer h3 { margin-bottom: 15px; color: #74b9ff; }'
    + '.contact-info { display: flex; justify-content: center; gap: 30px; flex-wrap: wrap; margin-top: 20px; }'
    + '.contact-item { display: flex; align-items: center; gap: 8px; font-size: 14px; }'
    + '@media (max-width: 600px) { .email-container { margin: 10px; } .content { padding: 20px; } .header { padding: 30px 20px; } .contact-info { flex-direction: column; gap: 15px; } .detail-row { flex-direction: column; align-items: flex-start; gap: 5px; } .detail-value { text-align: left; } }'
    + '</style></head><body>'
    + '<div class="email-container">'
    + '<div class="header"><h1>&#9917; Campo Don Michele FIORE &#9917;</h1><div class="subtitle">&#127942; Parrocchia Sant\'Agostino - ANSPI &#127942;</div></div>'
    + '<div class="event-banner">' + eventIcon + ' ' + eventName + (isExclusiveEvent ? '<br><small style="font-size:16px;opacity:0.9;">Evento Esclusivo</small>' : '') + '</div>'
    + '<div class="content">'
    + '<h2 style="color:#2d3436;margin-bottom:10px;">Ciao ' + bookingData.nome + '! &#128075;</h2>'
    + statusMessage
    + '<div class="booking-details">'
    + '<div class="detail-row"><div class="detail-label">&#128272; Codice Prenotazione</div><div class="detail-value" style="color:#e17055;font-family:monospace;">' + bookingData.id + '</div></div>'
    + '<div class="detail-row"><div class="detail-label">' + eventIcon + ' Evento</div><div class="detail-value">' + eventName + '</div></div>'
    + '<div class="detail-row"><div class="detail-label">&#127967; Campo</div><div class="detail-value">' + campoName + '</div></div>'
    + '<div class="detail-row"><div class="detail-label">&#128197; Data</div><div class="detail-value">' + formatDateForEmail(bookingData.data) + '</div></div>'
    + '<div class="detail-row"><div class="detail-label">&#128336; Orario</div><div class="detail-value">' + bookingData.ora + '</div></div>'
    + noteRow
    + '</div>'
    + infoBox
    + '<p style="margin-top:25px;color:#636e72;">' + closingText + '</p>'
    + '</div>'
    + '<div class="footer"><h3>&#128222; Contatti</h3>'
    + '<div class="contact-info">'
    + '<div class="contact-item"><span>&#128231;</span><span>prenotazionecampoanspi@gmail.com</span></div>'
    + '<div class="contact-item"><span>&#128205;</span><span>Parrocchia Sant\'Agostino - Giovinazzo (BA)</span></div>'
    + '<div class="contact-item"><span>&#128336;</span><span>Orari: 8:00 - 22:00</span></div>'
    + '</div>'
    + '<p style="margin-top:20px;font-size:14px;opacity:0.8;">Grazie per aver scelto il Campo Don Michele FIORE! &#9917;</p>'
    + '</div>'
    + '</div></body></html>';

  return html;
}

function createStatusUpdateEmailTemplate(bookingData, newStatus, reason) {
  reason = reason || '';
  var isApproved = newStatus === 'approvato';
  var eventIconMap = { 'calcetto': '⚽', 'pallavolo': '🏐', 'compleanno': '🎉', 'eventi': '🎪' };
  var eventIcon = eventIconMap[bookingData.evento] || '📅';
  var eventName = getEventName(bookingData.evento);
  var campoName = getCampoName(bookingData.campo);

  var statusConfig = {
    approvato: {
      color: '#28a745', icon: '✅', title: 'EVENTO APPROVATO!',
      message: 'Fantastico! Il tuo evento speciale è stato <strong>approvato</strong> e confermato.',
      info: 'Il tuo evento è ora ufficialmente confermato. Preparati per una giornata speciale!'
    },
    rifiutato: {
      color: '#dc3545', icon: '❌', title: 'EVENTO NON APPROVATO',
      message: 'Ci dispiace, ma la tua richiesta di evento speciale <strong>non può essere approvata</strong> per la data richiesta.',
      info: 'Contatta l\'amministratore per discutere alternative o modifiche alla tua richiesta.'
    }
  };

  var config = statusConfig[newStatus] || statusConfig['rifiutato'];
  var reasonHtml = reason
    ? '<div style="background:#fff3f3;border:1px solid #e17055;border-radius:10px;padding:15px;margin:20px 0;color:#c0392b;"><strong>Motivazione del rifiuto:</strong><p style="margin-top:5px;white-space:pre-wrap;"><em>' + reason + '</em></p></div>'
    : '';

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Aggiornamento Stato Evento - Campo ANSPI</title></head>'
    + '<body><div class="email-container">'
    + '<div class="header"><h1>⚽ Campo Don Michele FIORE ⚽</h1><div class="subtitle">🏆 Parrocchia Sant\'Agostino - ANSPI 🏆</div></div>'
    + '<div class="status-banner"><span class="status-icon">' + config.icon + '</span>' + config.title + '</div>'
    + '<div class="content">'
    + '<h2>Ciao ' + bookingData.nome + '! 👋</h2>'
    + '<div class="status-message"><strong>' + config.icon + ' Aggiornamento Importante</strong><br>' + config.message + '</div>'
    + reasonHtml
    + '<div class="booking-details">'
    + '<div class="detail-row"><div class="detail-label">🔐 Codice</div><div class="detail-value" style="color:#e17055;font-family:monospace;">' + bookingData.id + '</div></div>'
    + '<div class="detail-row"><div class="detail-label">' + eventIcon + ' Evento</div><div class="detail-value">' + eventName + '</div></div>'
    + '<div class="detail-row"><div class="detail-label">🏟️ Campo</div><div class="detail-value">' + campoName + '</div></div>'
    + '<div class="detail-row"><div class="detail-label">📅 Data</div><div class="detail-value">' + formatDateForEmail(bookingData.data) + '</div></div>'
    + '<div class="detail-row"><div class="detail-label">🕐 Orario</div><div class="detail-value">' + bookingData.ora + '</div></div>'
    + '<div class="detail-row"><div class="detail-label">📊 Stato</div><div class="detail-value" style="color:' + config.color + ';font-weight:bold;">' + config.icon + ' ' + newStatus.toUpperCase() + '</div></div>'
    + (bookingData.note ? '<div class="detail-row"><div class="detail-label">📝 Note</div><div class="detail-value" style="white-space:pre-wrap;text-align:left;">' + bookingData.note + '</div></div>' : '')
    + '</div>'
    + '<p style="margin-top:25px;color:#636e72;text-align:center;">' + (isApproved ? 'Ci vediamo al campo per il tuo evento speciale! 🎉' : 'Grazie per la comprensione. Siamo sempre disponibili per aiutarti.') + '</p>'
    + '</div>'
    + '<div class="footer"><h3>📞 Contatti</h3>'
    + '<div class="contact-info">'
    + '<div class="contact-item"><span>📧</span><span>prenotazionecampoanspi@gmail.com</span></div>'
    + '<div class="contact-item"><span>📍</span><span>Parrocchia Sant\'Agostino - Giovinazzo (BA)</span></div>'
    + '<div class="contact-item"><span>🕐</span><span>Orari: 8:00 - 22:00</span></div>'
    + '</div>'
    + '<p style="margin-top:20px;font-size:14px;opacity:0.8;">Grazie per aver scelto il Campo Don Michele FIORE! ⚽</p>'
    + '</div></div></body></html>';
}

// =====================
// RIEPILOGO GIORNALIERO
// =====================
function sendDailySummaryToAdmin() {
  try {
    Logger.log('📊 === RIEPILOGO GIORNALIERO ===');
    invalidateAllCache();
    Utilities.sleep(500);
    SpreadsheetApp.flush();

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAMES.EVENTS);

    if (!sheet) {
      sendTelegramNotification('❌ ERRORE: Foglio eventi non trovato', 'urgent');
      return { status: 'error', message: 'Foglio non trovato' };
    }

    var timezone = ss.getSpreadsheetTimeZone();
    var oggi = new Date();
    var oggiStr = Utilities.formatDate(oggi, timezone, 'yyyy-MM-dd');

    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return h ? h.toString().toLowerCase().trim() : ''; });

    var cols = {
      id: headers.indexOf('id'),
      nome: headers.indexOf('nome'),
      evento: headers.indexOf('evento'),
      data: headers.indexOf('data'),
      ora: headers.indexOf('ora'),
      campo: headers.indexOf('campo'),
      stato: headers.indexOf('stato')
    };

    var eventiOggi = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[cols.id]) continue;
      var dataEvento = row[cols.data];
      if (!dataEvento) continue;

      var dataStr = '';
      if (dataEvento instanceof Date) {
        dataStr = Utilities.formatDate(dataEvento, timezone, 'yyyy-MM-dd');
      } else if (typeof dataEvento === 'string') {
        dataStr = dataEvento.split('T')[0];
      } else {
        continue;
      }

      if (dataStr === oggiStr) {
        var evento = {
          id: row[cols.id],
          nome: row[cols.nome] || 'N/D',
          evento: row[cols.evento] || 'N/D',
          ora: row[cols.ora] || 'N/D',
          campo: row[cols.campo] || 'N/D',
          stato: row[cols.stato] || 'confermato'
        };
        if (evento.stato !== 'rifiutato') {
          eventiOggi.push(evento);
        }
      }
    }

    var message = '📊 *RIEPILOGO GIORNALIERO*\n';
    message += '📅 ' + oggi.toLocaleDateString('it-IT', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }) + '\n\n';

    if (eventiOggi.length === 0) {
      message += '🌟 *Nessun evento programmato per oggi*\n';
      message += '\n💡 Giornata libera per manutenzione o prenotazioni dell\'ultimo minuto';
    } else {
      message += '🎯 *' + eventiOggi.length + ' ' + (eventiOggi.length === 1 ? 'EVENTO' : 'EVENTI') + ' PROGRAMMATO/I:*\n\n';

      eventiOggi.sort(function(a, b) {
        return parseInt(a.ora.toString().split(':')[0]) - parseInt(b.ora.toString().split(':')[0]);
      });

      for (var j = 0; j < eventiOggi.length; j++) {
        var evt = eventiOggi[j];
        var icon = EVENT_TYPES[evt.evento] ? EVENT_TYPES[evt.evento].icon : '📌';
        var nomeEvento = EVENT_TYPES[evt.evento] ? EVENT_TYPES[evt.evento].name : evt.evento;
        var statoIcon = evt.stato === 'approvato' ? '✅' : (evt.stato === 'in_attesa' ? '⏳' : '✔️');

        message += icon + ' *' + nomeEvento + '*\n';
        message += '   ' + statoIcon + ' Ore ' + evt.ora + '\n';
        message += '   👤 ' + evt.nome + '\n';
        message += '   🏟️ ' + getCampoName(evt.campo) + '\n';
        if (evt.stato === 'in_attesa') {
          message += '   ⚠️ _In attesa di approvazione_\n';
        }
        message += '\n';
      }
    }

    message += '\n🕐 Generato: ' + new Date().toLocaleTimeString('it-IT');

    var cache = CacheService.getScriptCache();
    var cacheKey = 'daily_summary_' + oggiStr;
    var alreadySent = cache.get(cacheKey);

    if (alreadySent) {
      Logger.log('⏭️ Riepilogo già inviato oggi');
      return { status: 'skipped', message: 'Già inviato oggi', data: { eventiOggi: eventiOggi.length } };
    }

    var sent = sendTelegramNotification(message, 'daily_summary', { source: 'sendDailySummaryToAdmin' });

    if (sent) {
      cache.put(cacheKey, 'sent', 86400);
      Logger.log('✅ Telegram inviato e marcato');
    }

    return {
      status: 'ok',
      message: 'Riepilogo inviato: ' + eventiOggi.length + ' eventi',
      data: { eventi: eventiOggi.length, dettagli: eventiOggi }
    };

  } catch (error) {
    Logger.log('❌ ERRORE: ' + error.message);
    sendTelegramNotification('❌ ERRORE Riepilogo Giornaliero\n\n' + error.message, 'urgent');
    return { status: 'error', message: error.message };
  }
}

function shouldSendTelegramNotification(bookingData, notificationType) {
  try {
    if (!bookingData || typeof bookingData !== 'object') {
      debugLog('❌ shouldSendTelegramNotification: bookingData non valido');
      return false;
    }

    debugLog('🔔 === CONTROLLO NOTIFICA TELEGRAM ===');
    debugLog('📋 Tipo notifica: ' + (notificationType || 'prenotazione'));
    debugLog('👤 Cliente: ' + bookingData.nome);
    debugLog('📅 Data: ' + bookingData.data);
    debugLog('🕐 Ora: ' + bookingData.ora);
    debugLog('📝 Evento: ' + bookingData.evento);
    debugLog('📊 Stato: ' + (bookingData.stato || 'N/D'));
    debugLog('👨‍💼 Admin: ' + (bookingData.adminActionBy || '(vuoto)'));

    if (notificationType === 'daily_summary') {
      debugLog('📊 TELEGRAM: Riepilogo giornaliero - SEMPRE INVIA');
      return true;
    }

    if (bookingData.evento === 'compleanno' || bookingData.evento === 'eventi') {
      debugLog('🎉 TELEGRAM: Evento speciale - SEMPRE NOTIFICA');
      return true;
    }

    var isAdminAction = bookingData.adminActionBy &&
      bookingData.adminActionBy.trim() !== '' &&
      bookingData.adminActionBy !== 'undefined' &&
      bookingData.adminActionBy !== 'null';

    if (isAdminAction) {
      if (notificationType === 'new_booking' || notificationType === 'special_event') {
        debugLog('👨‍💼 TELEGRAM: Prenotazione creata da admin - INVIA');
        return true;
      }
      debugLog('🚫 TELEGRAM: SALTATO - Azione admin non di creazione');
      return false;
    }

    debugLog('✅ TELEGRAM: ABILITATO - Prenotazione utente');
    return true;

  } catch (error) {
    console.error('❌ Errore shouldSendTelegramNotification: ' + error.message);
    return true;
  }
}

function testMarkDailySummary() {
  try {
    var today = new Date().toISOString().split('T')[0];
    markDailySummarySent(today);
    var isMarked = checkIfDailySummarySent(today);
    return {
      status: 'ok',
      data: { date: today, marked: isMarked },
      message: 'Riepilogo test per ' + today + ': ' + (isMarked ? 'MARCATO' : 'NON MARCATO')
    };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

function clearDailySummaryCache() {
  try {
    var cache = CacheService.getScriptCache();
    var today = new Date().toISOString().split('T')[0];
    cache.remove('daily_summary_' + today);
    debugLog('✅ Cache riepilogo pulita per: ' + today);
    return { status: 'ok', data: { date: today, cacheCleared: true }, message: 'Cache riepilogo pulita per ' + today };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

function fixDuplicateDailySummary() {
  try {
    var today = new Date().toISOString().split('T')[0];
    clearDailySummaryCache();
    var disableResult = disableAllNotificationTriggers();
    var setupResult = setupEventNotifications();
    var testResult = testMarkDailySummary();
    return {
      status: 'ok',
      data: {
        date: today,
        cacheCleared: true,
        triggersRemoved: disableResult.deletedCount,
        triggersCreated: setupResult.data.triggersCreated,
        testResult: testResult
      },
      message: 'Problema riepiloghi duplicati risolto per ' + today
    };
  } catch (error) {
    return { status: 'error', message: 'Errore: ' + error.message };
  }
}

function testSameDayAlert() {
  try {
    var today = new Date().toISOString().split('T')[0];
    markDailySummarySent(today);
    var testBooking = {
      id: 'ANSPI-TEST-001',
      nome: 'Cliente Test',
      email: 'pasqualem27@gmail.com',
      telefono: '123456789',
      evento: 'calcetto',
      campo: 'calcetto',
      data: today,
      ora: '15:00',
      note: 'Test alert stesso giorno',
      adminActionBy: ''
    };
    var result = sendSameDayBookingAlert(testBooking);
    return {
      status: 'ok',
      data: { date: today, alertSent: result },
      message: 'Test alert stesso giorno per ' + today + ': ' + (result ? 'INVIATO' : 'NON INVIATO')
    };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

function getTelegramChatId(e) {
  try {
    var data = e && e.postData ? JSON.parse(e.postData.contents) : null;
    if (!data || !data.message || !data.message.chat || !data.message.chat.id) {
      return ContentService.createTextOutput('Dati non validi');
    }
    var chatId = data.message.chat.id;
    var url = 'https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/sendMessage';
    UrlFetchApp.fetch(url, {
      method: 'post',
      payload: { chat_id: chatId, text: '✅ Il tuo chat ID Telegram è: ' + chatId, parse_mode: 'HTML' }
    });
    return ContentService.createTextOutput('Chat ID inviato su Telegram: ' + chatId);
  } catch (e) {
    return ContentService.createTextOutput('Errore: ' + e.message);
  }
}

function identificaChatIds() {
  for (var i = 0; i < TELEGRAM_CHAT_IDS.length; i++) {
    var chatId = TELEGRAM_CHAT_IDS[i];
    var url = 'https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/sendMessage';
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        chat_id: chatId,
        text: '👤 Questo messaggio è destinato al chat ID: ' + chatId
      })
    });
  }
}
