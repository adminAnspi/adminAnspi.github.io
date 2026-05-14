// =====================
// 4. FUNZIONI DI LETTURA DATI - AGGIORNATE
// =====================

/**
 * Funzione centralizzata per gestire la cache di qualsiasi dato
 * @param {string} key - Chiave univoca della cache
 * @param {function} fallbackFn - Funzione che calcola il dato se non presente in cache
 * @param {number} ttl - Tempo di vita della cache in secondi (default 600)
 * @returns {*} Il dato richiesto, dalla cache o calcolato
 */

function getAllCacheKeys() {
  return [
    'all_bookings',
    'calendar_bookings_fast',
    'bookings_data',
    'archived_bookings',
    'archived_bookings_safe',
    'event_stats',
    'public_calendar_cache',
    'public_events_cache',
    CACHE_KEYS.BOOKINGS,
    CACHE_KEYS.EVENTS,
    CACHE_KEYS.ARCHIVED,
    CACHE_KEYS.PUBLIC_CALENDAR,
    CACHE_KEYS.PUBLIC_EVENTS,
    CACHE_KEYS.UNIFIED_ALL,
    CACHE_KEYS.EVENT_STATS 
  ];
}

function getOrSetCache(key, fallbackFn, ttl = CACHE_DURATION) {
  const cache = CacheService.getScriptCache();
  let cached = cache.get(key);
  if (cached !== null) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      debugLog('Errore parsing cache per ' + key + ', ricalcolo.');
    }
  }
  // Calcola il dato e salva in cache
  const value = fallbackFn();
  try {
    cache.put(key, JSON.stringify(value), ttl);
  } catch (e) {
    debugLog('Errore salvataggio cache per ' + key + ': ' + e.message);
  }
  return value;
}

function clearBookingsCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove(CACHE_KEYS.BOOKINGS);
    debugLog('🗑️ Cache bookings pulita');
    return { status: 'ok', message: 'Cache pulita con successo' };
  } catch (error) {
    debugLog(`❌ Errore pulizia cache: ${error.message}`);
    return { status: 'error', message: error.message };
  }
}

function setCachedData(key, data) {
  try {
    const cache = CacheService.getScriptCache();
    cache.put(key, JSON.stringify(data), CACHE_DURATION);
    debugLog(`✅ Dati salvati in cache: ${key}`);
    return true;
  } catch (error) {
    debugLog(`❌ Errore salvataggio cache ${key}: ${error.toString()}`);
    return false;
  }
}

function getCacheStatus() {
  try {
    const cache = CacheService.getScriptCache();
    const status = {};

    Object.entries(CACHE_KEYS).forEach(([name, key]) => {
      const cached = cache.get(key);
      status[name] = {
        exists: !!cached,
        size: cached ? cached.length : 0
      };
    });

    // Aggiungi anche la cache legacy se serve
    const existingCache = cache.get('all_bookings');
    status.BOOKINGS_OLD = {
      exists: !!existingCache,
      size: existingCache ? existingCache.length : 0
    };

    return {
      success: true,
      status: status,
      cacheDuration: CACHE_DURATION
    };

  } catch (error) {
    return {
      success: false,
      message: `Errore stato cache: ${error.toString()}`
    };
  }
}

function clearAllCaches() {
  try {
    const cache = CacheService.getScriptCache();
    getAllCacheKeys().forEach(key => {
      try {
        cache.remove(key);
      } catch (e) {}
    });
    debugLog('🗑️ Tutte le cache pulite');
    return { status: 'ok' };
  } catch (e) {
    debugLog('❌ Errore pulizia: ' + e.message);
    return { status: 'error', message: e.message };
  }
}

function invalidateAllCache() {
  try {
    const cache = CacheService.getScriptCache();
    let removedCount = 0;
    getAllCacheKeys().forEach(key => { 
      try {
        cache.remove(key);
        removedCount++;
      } catch (e) {}
    });
    debugLog(`✅ Cache invalidate: ${removedCount} chiavi`);
    return removedCount;
  } catch (error) {
    debugLog('❌ Errore invalidazione: ' + error.message);
    return 0;
  }
}
