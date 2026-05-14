// =====================
// 11. FUNZIONI ADMIN E AUTENTICAZIONE
// =====================
function updateEventStatus(data) {
  const { bookingId, newStatus, reason, adminEmail } = data;
  const validStates = Object.keys(STATUS_BADGES);
  if (!validStates.includes(newStatus)) throw new Error(`Stato non valido: ${newStatus}. Valori ammessi: ${validStates.join(', ')}`);
  const sheet = getSheet();
  const dataRange = sheet.getDataRange().getValues();
  const headers = dataRange[0].map(h => h ? h.toLowerCase().trim() : '');
  const idColumnIndex = headers.indexOf('id');
  const eventoColumnIndex = headers.indexOf('evento');
  const statoColumnIndex = headers.indexOf('stato');
  const adminActionColIndex = headers.indexOf('adminactionby');
  const nomeColumnIndex = headers.indexOf('nome');
  const emailColumnIndex = headers.indexOf('email');
  const telefonoColumnIndex = headers.indexOf('telefono');
  const campoColumnIndex = headers.indexOf('campo');
  const dataColumnIndex = headers.indexOf('data');
  const oraColumnIndex = headers.indexOf('ora');
  const noteColumnIndex = headers.indexOf('note');
  if (idColumnIndex === -1 || statoColumnIndex === -1) throw new Error('Colonne ID o Stato non trovate');
  for (let i = 1; i < dataRange.length; i++) {
    if (String(dataRange[i][idColumnIndex]).trim() === String(bookingId).trim()) {
      const evento = dataRange[i][eventoColumnIndex];
      if (evento !== 'compleanno' && evento !== 'eventi') throw new Error('Solo gli eventi speciali possono avere uno stato');
      const bookingData = {
        id: dataRange[i][idColumnIndex],
        nome: dataRange[i][nomeColumnIndex],
        email: dataRange[i][emailColumnIndex],
        telefono: dataRange[i][telefonoColumnIndex],
        evento: dataRange[i][eventoColumnIndex],
        campo: dataRange[i][campoColumnIndex],
        data: dataRange[i][dataColumnIndex],
        ora: dataRange[i][oraColumnIndex],
        note: dataRange[i][noteColumnIndex] || '',
        stato: newStatus
      };
      sheet.getRange(i + 1, statoColumnIndex + 1).setValue(newStatus);
      if (adminActionColIndex !== -1 && adminEmail) {
        sheet.getRange(i + 1, adminActionColIndex + 1).setValue(adminEmail);
      }
      SpreadsheetApp.flush();
      // Invalida tutta la cache dopo aggiornamento stato
      invalidateAllCache();

      try {
        var msg = `Stato evento aggiornato\nCliente: <b>${bookingData.nome}</b>\nEvento: <b>${bookingData.evento}</b>\nCampo: <b>${bookingData.campo}</b>\nData: <b>${bookingData.data}</b>\nOra: <b>${bookingData.ora}</b>\nNuovo stato: <b>${newStatus}</b>`;
        sendUrgentNotification(msg, 'Aggiornamento Stato Evento', 'special_event');
        if (bookingData.email && bookingData.email.includes('@')) {
          var subject = newStatus === 'approvato' ? `✅ Evento Approvato - ${bookingData.id}` : `❌ Evento Non Approvato - ${bookingData.id}`;
          var htmlBody = createStatusUpdateEmailTemplate(bookingData, newStatus, reason || '');
          MailApp.sendEmail({ to: bookingData.email, subject: subject, htmlBody: htmlBody });
        }
      } catch(emailError) {}
      return {
        status: 'ok',
        message: `Evento ${newStatus === 'approvato' ? 'approvato' : newStatus === 'rifiutato' ? 'rifiutato' : 'aggiornato'} con successo`,
        data: {
          bookingId: bookingId,
          newStatus: newStatus,
          emailSent: true,
          date: bookingData.data,
          ora: bookingData.ora
        }
      };
    }
  }
  throw new Error('Prenotazione non trovata');
}

function getEventStats() {
  try {
    return getOrSetCache(CACHE_KEYS.EVENT_STATS, () => {
      let eventBookings = [];
      try {
        eventBookings = getBookings();
      } catch (e) {
        debugLog('Errore recupero bookings in getEventStats: ' + e.message);
      }
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const weekStart = new Date(today);
      weekStart.setDate(today.getDate() - today.getDay() + 1); // Lunedì
      const stats = {
        total: eventBookings.length,
        today: eventBookings.filter(b => b.data && new Date(b.data).toDateString() === today.toDateString()).length,
        week: eventBookings.filter(b => b.data && new Date(b.data) >= weekStart).length,
        active: eventBookings.filter(b => b.data && new Date(b.data) >= today).length
      };
      const specialEvents = eventBookings.filter(b => b.evento === 'compleanno' || b.evento === 'eventi');
      const statusCounts = {
        tutti: specialEvents.length,
        in_attesa: specialEvents.filter(b => (b.stato || 'in_attesa') === 'in_attesa').length,
        approvato: specialEvents.filter(b => b.stato === 'approvato').length,
        rifiutato: specialEvents.filter(b => b.stato === 'rifiutato').length
      };
      return {
        status: 'ok',
        data: {
          ...stats,
          statusCounts: statusCounts
        }
      };
    }, 60); // 1 minuto
  } catch (e) {
    debugLog('Errore in getEventStats: ' + e.message);
    return {
      status: 'error',
      message: e.message,
      data: { statusCounts: {} }
    };
  }
}

function getAdminSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  if (!sheet) throw new Error(`Foglio amministratori "${SHEET_NAMES.USERS}" non trovato.`);
  return sheet;
}

function getPasswordSalt() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let salt = scriptProperties.getProperty(PASSWORD_SALT_KEY);
  if (!salt) {
    const hashBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Math.random().toString());
    salt = Utilities.base64Encode(hashBytes);
    scriptProperties.setProperty(PASSWORD_SALT_KEY, salt);
    debugLog('Nuovo sale per le password creato e salvato.');
  }
  return salt;
}

function findAdminByEmail(email) {
  if (!email) return null;
  const adminSheet = getAdminSheet();
  const data = adminSheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toLowerCase());
  const emailCol = headers.indexOf('email');
  const passwordHashCol = headers.indexOf('passwordhash');
  for (let i = 1; i < data.length; i++) {
    if (data[i][emailCol]) {
      if (data[i][emailCol].toString().toLowerCase().trim() === email.toLowerCase().trim()) {
        return {
          row: i + 1,
          email: data[i][emailCol],
          passwordHash: data[i][passwordHashCol]
        };
      }
    }
  }
  return null;
}

function hashPassword(password) {
  const salt = getPasswordSalt();
  const saltedPassword = password + salt;
  const hashBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, saltedPassword);
  return Utilities.base64Encode(hashBytes);
}

function adminLogin(data) {
  const { email, password } = data;
  const admin = findAdminByEmail(email);
  if (!admin) throw new Error("Email non autorizzata.");
  if (!admin.passwordHash) {
    if (password) throw new Error("Primo accesso: lascia il campo password vuoto per crearne una.");
    return { status: 'ok', data: { action: 'CREATE_PASSWORD' } };
  }
  if (!password) throw new Error("Password richiesta.");
  const providedPasswordHash = hashPassword(password);
  if (providedPasswordHash !== admin.passwordHash) throw new Error("Password errata.");
  return { status: 'ok', data: { action: 'LOGIN_SUCCESS' } };
}

function createAdminPassword(data) {
  const { email, password } = data;
  const admin = findAdminByEmail(email);
  if (!admin) throw new Error("Email non autorizzata.");
  if (admin.passwordHash) throw new Error("La password per questo account è già stata impostata.");
  if (!password || password.length < 6) throw new Error("La password deve essere di almeno 6 caratteri.");
  const newPasswordHash = hashPassword(password);
  const adminSheet = getAdminSheet();
  const adminHeaders = adminSheet.getRange(1, 1, 1, adminSheet.getLastColumn()).getValues()[0];
const passwordHashCol = adminHeaders.map(h => h.toLowerCase()).indexOf('passwordhash');
if (passwordHashCol === -1) throw new Error('Colonna passwordhash non trovata');
adminSheet.getRange(admin.row, passwordHashCol + 1).setValue(newPasswordHash);
  return { status: 'ok', message: 'Password creata con successo.' };
}

// ✅ FUNZIONI RIMOSSE: Sistema centralizzato non più necessario
// Torniamo al sistema semplice e funzionante con modal

function requestPasswordReset(data) {
  const email = data.email;
  const admin = findAdminByEmail(email);
  if (admin) {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const resetSheet = ss.getSheetByName("PasswordResets");
    if (!resetSheet) throw new Error("Foglio 'PasswordResets' non trovato. Eseguire il setup.");
    const token = Utilities.getUuid();
    const expiration = new Date(new Date().getTime() + 3600000); // Scade tra 1 ora
    resetSheet.appendRow([email, token, expiration]);
    const webAppUrl = ScriptApp.getService().getUrl();
    const resetLink = `${webAppUrl}?resetToken=${token}`;
    const subject = "🔑 Richiesta di Reset Password - Campo ANSPI";
    const body = `<p>Ciao,</p><p>Abbiamo ricevuto una richiesta di reset della password per il tuo account.</p><p>Clicca sul link qui sotto per impostare una nuova password. Il link scadrà tra un'ora:</p><p><a href="${resetLink}" style="padding: 10px 15px; background-color: #007bff; color: white; text-decoration: none; border-radius: 5px;">Imposta Nuova Password</a></p><p>Se non hai richiesto tu il reset, puoi ignorare questa email.</p><p>Grazie,<br>Sistema di Prenotazione ANSPI</p>`;
    MailApp.sendEmail(email, subject, "", { htmlBody: body });
  }
  return { status: 'ok', message: 'Se l\'email è registrata, riceverai un link per il reset.' };
}

function performPasswordReset(data) {
  const { token, newPassword } = data;
  if (!token || !newPassword || newPassword.length < 6) throw new Error("Token non valido o password troppo corta.");
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const resetSheet = ss.getSheetByName("PasswordResets");
  if (!resetSheet) throw new Error("Foglio 'PasswordResets' non trovato.");
  const dataRange = resetSheet.getDataRange().getValues();
  const now = new Date();
  let tokenFound = false;
  for (let i = dataRange.length - 1; i > 0; i--) {
    const [email, storedToken, expiration] = dataRange[i];
    if (storedToken === token) {
      tokenFound = true;
      if (new Date(expiration) < now) {
        resetSheet.deleteRow(i + 1);
        throw new Error("Token scaduto. Richiedi un nuovo reset.");
      }
      const admin = findAdminByEmail(email);
      if (!admin) throw new Error("Utente non trovato.");
      const newPasswordHash = hashPassword(newPassword);
      const adminSheet = getAdminSheet();
      adminSheet.getRange(admin.row, 2).setValue(newPasswordHash);
      resetSheet.deleteRow(i + 1);
      return { status: 'ok', message: 'Password aggiornata con successo!' };
    }
  }
  if (!tokenFound) throw new Error("Token non valido o già utilizzato.");
}

// === NUOVE FUNZIONI PER GESTIONE ADMIN MASTER ===

/**
 * Verifica se l'email è il master admin
 */
function isMasterAdmin(email) {
return email && email.toLowerCase().trim() === MASTER_ADMIN_EMAIL.toLowerCase();
}

function getAdminChatIds() {
  var sheet = getAdminSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return h ? h.toLowerCase().trim() : ''; });
  var chatCol = headers.indexOf('telegramchatid');
  var ids = [];
  if (chatCol !== -1) {
    for (var i = 1; i < data.length; i++) {
      var v = String(data[i][chatCol] || '').trim();
      if (v) ids.push(v);
    }
  }
  if (ids.length > 0) return ids;
  if (typeof TELEGRAM_CHAT_IDS !== 'undefined' && Array.isArray(TELEGRAM_CHAT_IDS) && TELEGRAM_CHAT_IDS.length > 0) return TELEGRAM_CHAT_IDS;
  return [TELEGRAM_MASTER_CHAT_ID];
}

/**
 * Ottiene la lista di tutti gli admin (solo per master admin)
 */
function getAllAdmins() {
  try {
    const adminSheet = getAdminSheet();
    const data = adminSheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toLowerCase());
    const emailCol = headers.indexOf('email');
    const passwordHashCol = headers.indexOf('passwordhash');
    
    const admins = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][emailCol]) {
        admins.push({
          email: data[i][emailCol],
          hasPassword: !!data[i][passwordHashCol],
          isMaster: isMasterAdmin(data[i][emailCol])
        });
      }
    }
    
    return { status: 'ok', data: { admins } };
  } catch (error) {
    debugLog(`❌ Errore nel recupero admin: ${error.message}`);
    return { status: 'error', message: error.message };
  }
}

/**
 * Aggiunge un nuovo admin (solo per master admin)
 */
function addNewAdmin(data) {
  const { masterEmail, newAdminEmail } = data;
  
  try {
    // Verifica che sia il master admin
    if (!isMasterAdmin(masterEmail)) {
      return { status: 'error', message: 'Solo il master admin può aggiungere nuovi admin.' };
    }
    
    // Verifica email valida
    if (!newAdminEmail || !newAdminEmail.includes('@')) {
      return { status: 'error', message: 'Email non valida.' };
    }
    
    const adminSheet = getAdminSheet();
    const dataRange = adminSheet.getDataRange().getValues();
    const headers = dataRange[0].map(h => h.toLowerCase());
    const emailCol = headers.indexOf('email');
    
    // Verifica se l'admin esiste già
    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][emailCol] && dataRange[i][emailCol].toLowerCase().trim() === newAdminEmail.toLowerCase().trim()) {
        return { status: 'error', message: 'Questo admin esiste già.' };
      }
    }
    
    // Aggiungi il nuovo admin (senza password, dovrà crearla al primo accesso)
    adminSheet.appendRow([newAdminEmail, '']);
    
    debugLog(`✅ Nuovo admin aggiunto: ${newAdminEmail}`);
    
    // Invia email di benvenuto
    sendWelcomeEmailToNewAdmin(newAdminEmail);
    
    return { 
      status: 'ok', 
      message: `Admin ${newAdminEmail} aggiunto con successo. Riceverà un'email di benvenuto.` 
    };
    
  } catch (error) {
    debugLog(`❌ Errore nell'aggiunta admin: ${error.message}`);
    return { status: 'error', message: error.message };
  }
}

/**
 * Rimuove un admin (solo per master admin)
 */
function removeAdmin(data) {
  const { masterEmail, adminEmailToRemove } = data;
  
  try {
    // Verifica che sia il master admin
    if (!isMasterAdmin(masterEmail)) {
      return { status: 'error', message: 'Solo il master admin può rimuovere admin.' };
    }
    
    // Non permettere di rimuovere se stesso
    if (isMasterAdmin(adminEmailToRemove)) {
      return { status: 'error', message: 'Non puoi rimuovere il master admin.' };
    }
    
    const adminSheet = getAdminSheet();
    const dataRange = adminSheet.getDataRange().getValues();
    const headers = dataRange[0].map(h => h.toLowerCase());
    const emailCol = headers.indexOf('email');
    
    // Trova l'admin da rimuovere
    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][emailCol] && dataRange[i][emailCol].toLowerCase().trim() === adminEmailToRemove.toLowerCase().trim()) {
        adminSheet.deleteRow(i + 1);
        debugLog(`✅ Admin rimosso: ${adminEmailToRemove}`);
        return { 
          status: 'ok', 
          message: `Admin ${adminEmailToRemove} rimosso con successo.` 
        };
      }
    }
    
    return { status: 'error', message: 'Admin non trovato.' };
    
  } catch (error) {
    debugLog(`❌ Errore nella rimozione admin: ${error.message}`);
    return { status: 'error', message: error.message };
  }
}

/**
 * Resetta la password di un admin (solo per master admin)
 */
function resetAdminPassword(data) {
  const { masterEmail, adminEmailToReset } = data;
  
  try {
    // Verifica che sia il master admin
    if (!isMasterAdmin(masterEmail)) {
      return { status: 'error', message: 'Solo il master admin può resettare le password.' };
    }
    
    const adminSheet = getAdminSheet();
    const dataRange = adminSheet.getDataRange().getValues();
    const headers = dataRange[0].map(h => h.toLowerCase());
    const emailCol = headers.indexOf('email');
    const passwordHashCol = headers.indexOf('passwordhash');
    
    // Trova l'admin e resetta la password
    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][emailCol] && dataRange[i][emailCol].toLowerCase().trim() === adminEmailToReset.toLowerCase().trim()) {
        // Rimuovi la password hash (l'admin dovrà creare una nuova password)
        adminSheet.getRange(i + 1, passwordHashCol + 1).setValue('');
        
        debugLog(`✅ Password resettata per: ${adminEmailToReset}`);
        
        // Invia email di notifica
        sendPasswordResetNotificationEmail(adminEmailToReset);
        
        return { 
          status: 'ok', 
          message: `Password resettata per ${adminEmailToReset}. Riceverà un'email di notifica.` 
        };
      }
    }
    
    return { status: 'error', message: 'Admin non trovato.' };
    
  } catch (error) {
    debugLog(`❌ Errore nel reset password: ${error.message}`);
    return { status: 'error', message: error.message };
  }
}

/**
 * Invia email di benvenuto al nuovo admin
 */
function sendWelcomeEmailToNewAdmin(email) {
  try {
    const subject = "👨‍💼 Benvenuto come Amministratore - Campo ANSPI";
    const webAppUrl = ScriptApp.getService().getUrl();
    const body = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #2c3e50;">👨‍💼 Benvenuto come Amministratore!</h2>
        <p>Ciao,</p>
        <p>Sei stato aggiunto come amministratore del sistema di prenotazioni del Campo Don Michele FIORE - ANSPI.</p>
        <p><strong>Per accedere al sistema:</strong></p>
        <ol>
          <li>Vai al <a href="${webAppUrl}" style="color: #3498db;">sistema di prenotazioni</a></li>
          <li>Clicca su "👨‍💼 Area Amministratore"</li>
          <li>Inserisci la tua email: <strong>${email}</strong></li>
          <li>Lascia vuoto il campo password e clicca "🚀 Accedi"</li>
          <li>Il sistema ti chiederà di creare una nuova password</li>
        </ol>
        <p><strong>Funzioni disponibili:</strong></p>
        <ul>
          <li>📊 Dashboard con statistiche</li>
          <li>📅 Gestione calendario eventi</li>
          <li>📋 Gestione prenotazioni</li>
          <li>🎉 Approvazione eventi speciali</li>
          <li>➕ Creazione prenotazioni</li>
        </ul>
        <p style="margin-top: 30px; padding: 15px; background-color: #f8f9fa; border-left: 4px solid #3498db;">
          <strong>🔐 Sicurezza:</strong> La tua password sarà crittografata e sicura. Non condividerla con nessuno.
        </p>
        <p>Grazie,<br>Sistema di Prenotazione ANSPI</p>
      </div>
    `;
    
    MailApp.sendEmail(email, subject, "", { htmlBody: body });
    debugLog(`📧 Email di benvenuto inviata a: ${email}`);
    
  } catch (error) {
    debugLog(`❌ Errore invio email benvenuto: ${error.message}`);
  }
}

/**
 * Invia email di notifica reset password
 */
function sendPasswordResetNotificationEmail(email) {
  try {
    const subject = "🔑 Password Resettata - Campo ANSPI";
    const webAppUrl = ScriptApp.getService().getUrl();
    const body = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: #e74c3c;">🔑 Password Resettata</h2>
        <p>Ciao,</p>
        <p>La tua password di amministratore è stata resettata dal master admin.</p>
        <p><strong>Per reimpostare la password:</strong></p>
        <ol>
          <li>Vai al <a href="${webAppUrl}" style="color: #3498db;">sistema di prenotazioni</a></li>
          <li>Clicca su "👨‍💼 Area Amministratore"</li>
          <li>Inserisci la tua email: <strong>${email}</strong></li>
          <li>Lascia vuoto il campo password e clicca "🚀 Accedi"</li>
          <li>Il sistema ti chiederà di creare una nuova password</li>
        </ol>
        <p style="margin-top: 30px; padding: 15px; background-color: #fff3cd; border-left: 4px solid #ffc107;">
          <strong>⚠️ Attenzione:</strong> Se non hai richiesto tu questo reset, contatta immediatamente il master admin.
        </p>
        <p>Grazie,<br>Sistema di Prenotazione ANSPI</p>
      </div>
    `;
    
    MailApp.sendEmail(email, subject, "", { htmlBody: body });
    debugLog(`📧 Email notifica reset inviata a: ${email}`);
    
  } catch (error) {
    debugLog(`❌ Errore invio email notifica reset: ${error.message}`);
  }
}

// ✅ NUOVA FUNZIONE: Inizializza admin mancanti
function initializeMissingAdmins() {
  try {
    const adminSheet = getAdminSheet();
    const dataRange = adminSheet.getDataRange().getValues();
    const headers = dataRange[0].map(h => h.toLowerCase());
    const emailCol = headers.indexOf('email');
    
    // Lista degli admin che dovrebbero esistere
    const requiredAdmins = ADMIN_EMAILS_FOR_NOTIFICATIONS;
    
    let addedCount = 0;
    
    for (const adminEmail of requiredAdmins) {
      let adminExists = false;
      
      // Verifica se l'admin esiste già
      for (let i = 1; i < dataRange.length; i++) {
        if (dataRange[i][emailCol] && 
            dataRange[i][emailCol].toString().toLowerCase().trim() === adminEmail.toLowerCase().trim()) {
          adminExists = true;
          break;
        }
      }
      
      // Se non esiste, aggiungilo
      if (!adminExists) {
        adminSheet.appendRow([adminEmail, '']);
        addedCount++;
        debugLog(`✅ Admin aggiunto: ${adminEmail}`);
      }
    }
    
    if (addedCount > 0) {
      debugLog(`✅ Inizializzazione completata: ${addedCount} admin aggiunti`);
      return { 
        status: 'ok', 
        message: `Inizializzazione completata: ${addedCount} admin aggiunti automaticamente.` 
      };
    } else {
      return { 
        status: 'ok', 
        message: 'Tutti gli admin richiesti sono già presenti.' 
      };
    }
    
  } catch (error) {
    debugLog(`❌ Errore inizializzazione admin: ${error.message}`);
    return { status: 'error', message: error.message };
  }
}
