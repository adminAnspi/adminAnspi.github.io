// =====================
// COSTANTI GLOBALI PER GOOGLE APPS SCRIPT
// =====================

var _props = PropertiesService.getScriptProperties();
var INIT_ATTEMPTS_LIMIT = 40; // Limite massimo tentativi inizializzazione
var LOCK_TIMEOUT = 10000; // 10 secondi per operazioni critiche
var MAX_EXECUTION_TIME = 25000; // 25 secondi timeout
var TEMP_ADMIN_PASSWORD = _props.getProperty('TEMP_ADMIN_PASSWORD');
var SHEET_ID = _props.getProperty('SHEET_ID');

// ✅ FIX CRITICO: EVENTS deve essere "Prenotazioni1" non "eventi"
var SHEET_NAMES = {
  EVENTS: "Prenotazioni1",
  USERS: "AdminUsers",
  CONFIG: "Configurazione",
  ARCHIVE: "Archivio_Prenotazioni"
};

var ADMIN_EMAILS_FOR_NOTIFICATIONS = [
  "prenotazionecampoanspi@gmail.com",
  "pasqualem27@gmail.com",
  "tommystrega@libero.it",
  "lor.tur.71@gmail.com",
];

var MASTER_ADMIN_EMAIL = _props.getProperty('MASTER_ADMIN_EMAIL');
var PASSWORD_SALT_KEY = 'PASSWORD_SALT';

var CACHE_KEYS = {
  BOOKINGS: 'bookings_cache',
  EVENTS: 'events_cache',
  USERS: 'users_cache',
  CONFIG: 'config_cache',
  ARCHIVED: 'archived_bookings',
  PUBLIC_CALENDAR: 'public_calendar_cache',
  PUBLIC_EVENTS: 'public_events_cache',
  UNIFIED_ALL: 'events_unified_cache_all',
  EVENT_STATS: 'event_stats' 
};

var CACHE_DURATION = 3600; // 1 ora

var DURATION_MAP = {
  calcetto: 90,
  pallavolo: 120,
  compleanno: 180,
  eventi: 180
};

var OPENING_HOURS = {
  start: 8,
  end: 22,
  lunchStart: 13,
  lunchEnd: 16
};

var ITALIAN_HOLIDAYS = [
  '01-01', '01-06', '04-25', '05-01', '06-02', '08-15', '11-01', '12-08', '12-25', '12-26'
];

var EVENT_TYPES = {
  calcetto: { name: 'Calcetto', icon: '⚽', color: 'var(--primary-soft)' },
  pallavolo: { name: 'Volley', icon: '🏐', color: 'var(--success-soft)' },
  compleanno: { name: 'Festa', icon: '🎉', color: 'var(--warning-soft)' },
  eventi: { name: 'Evento Straordinario', icon: '🎪', color: 'var(--danger-soft)' }
};

var STATUS_BADGES = {
  in_attesa: '<span style="background: #ffc107; color: #000; padding: 4px 8px; border-radius: 12px; font-size: 12px; font-weight: bold;">⏳ In Attesa</span>',
  approvato: '<span style="background: #28a745; color: #fff; padding: 4px 8px; border-radius: 12px; font-size: 12px; font-weight: bold;">✅ Approvato</span>',
  rifiutato: '<span style="background: #dc3545; color: #fff; padding: 4px 8px; border-radius: 12px; font-size: 12px; font-weight: bold;">❌ Rifiutato</span>'
};

var AVAILABILITY_ICONS = {
  available: '🟢',
  partial: '🟡',
  occupied: '🔴',
  special: '🔵',
  past: '⚫',
  closed: '🚪'
};

var ICONS = {
  calcetto: '⚽',
  pallavolo: '🏐',
  compleanno: '🎉',
  eventi: '🎪',
  cliente: '👤',
  email: '📧',
  telefono: '📞',
  evento: '🎯',
  campo: '🏟️',
  data: '📅',
  ora: '🕐',
  note: '📝',
  codice: '🔐',
  timestamp: '🕒'
};

var TELEGRAM_BOT_TOKEN = _props.getProperty('TELEGRAM_BOT_TOKEN');
var TELEGRAM_CHAT_IDS = ['686525181','1619944598'];
var TELEGRAM_MASTER_CHAT_ID = '686525181';

// =====================
// FUNZIONI HELPER CENTRALIZZATE
// =====================
function getEventName(type) {
  return EVENT_TYPES[type] ? EVENT_TYPES[type].icon + ' ' + EVENT_TYPES[type].name : type;
}

function getShortEventName(type) {
  return EVENT_TYPES[type] ? EVENT_TYPES[type].name : type;
}

function getEventColor(type) {
  return EVENT_TYPES[type] ? EVENT_TYPES[type].color : 'var(--secondary-soft)';
}

function getStatusBadge(stato) {
  return STATUS_BADGES[stato] || STATUS_BADGES['in_attesa'];
}

function getAvailabilityIcon(status) {
  return AVAILABILITY_ICONS[status] || '⚪';
}

function getCampoName(type) {
  var campos = {
    calcetto: 'Calcetto',
    pallavolo: 'Pallavolo',
    entrambi: 'Entrambi',
    'admin-managed': 'Da definire'
  };
  return campos[type] || type;
}
var APP_VERSION = '3.0';
