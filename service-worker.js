// ═══════════════════════════════════════════════
//  SERVICE WORKER — Dictamen Forestal Zapopan
//  Permite uso offline completo
// ═══════════════════════════════════════════════

const CACHE_NAME = 'dictamen-forestal-v32';
const BASE = '/dictamen-forestal';

// Archivos a cachear para uso offline
const ARCHIVOS = [
  BASE + '/',
  BASE + '/index.html',
  BASE + '/manifest.json',
  // Librerías externas
  'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.css',
  'https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;600;700&display=swap'
];

// ── INSTALACIÓN: cachear todos los archivos ──
self.addEventListener('install', function(event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      console.log('[SW] Cacheando archivos para uso offline...');
      return cache.addAll(ARCHIVOS).catch(function(err) {
        console.warn('[SW] Algunos archivos no se pudieron cachear:', err);
      });
    }).then(function() {
      return self.skipWaiting(); // Activar inmediatamente
    })
  );
});

// ── ACTIVACIÓN: limpiar cachés viejas ──
self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(key) {
          return key !== CACHE_NAME;
        }).map(function(key) {
          console.log('[SW] Eliminando caché vieja:', key);
          return caches.delete(key);
        })
      );
    }).then(function() {
      return self.clients.claim(); // Tomar control inmediato
    })
  );
});

// ── FETCH: estrategia Cache First con fallback a red ──
self.addEventListener('fetch', function(event) {
  // Ignorar peticiones al Apps Script de Drive (necesitan red)
  if (event.request.url.includes('script.google.com')) {
    return; // Dejar pasar sin interceptar
  }

  // Ignorar peticiones POST
  if (event.request.method !== 'GET') {
    return;
  }

  event.respondWith(
    caches.match(event.request).then(function(cachedResponse) {
      if (cachedResponse) {
        // Servir desde caché (funciona offline)
        return cachedResponse;
      }

      // No está en caché: intentar red
      return fetch(event.request).then(function(networkResponse) {
        // Si la respuesta es válida, guardarla en caché para la próxima vez
        if (networkResponse && networkResponse.status === 200 && networkResponse.type !== 'opaque') {
          var responseClone = networkResponse.clone();
          caches.open(CACHE_NAME).then(function(cache) {
            cache.put(event.request, responseClone);
          });
        }
        return networkResponse;
      }).catch(function() {
        // Sin red y sin caché: mostrar página offline si es navegación
        if (event.request.mode === 'navigate') {
          return caches.match(BASE + '/index.html');
        }
      });
    })
  );
});

// ── SYNC EN SEGUNDO PLANO (cuando recupera internet) ──
self.addEventListener('sync', function(event) {
  if (event.tag === 'sync-dictamen') {
    console.log('[SW] Sincronizando dictamen pendiente con Drive...');
    // La sincronización real la maneja la app cuando detecta conexión
  }
});

// ── MENSAJE desde la app ──
self.addEventListener('message', function(event) {
  if (event.data && event.data.action === 'skipWaiting') {
    self.skipWaiting();
  }
});
