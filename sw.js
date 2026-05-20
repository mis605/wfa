// ============================================================
// SW.JS - Service Worker (PWA)
// ============================================================

const CACHE_NAME = 'absen-wfa-v2.3';
const STATIC_ASSETS = [
  './',
  './index.html',
  './css/style.css',
  './js/config.js',
  './js/auth.js',
  './js/graph.js',
  './js/utils.js',
  './js/app.js',
  './js/msal-browser.min.js',
  './manifest.json',
  './icons/icon.svg',
  './icons/icon-192.png',
  './icons/icon-512.png',
  // Google Fonts
  'https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap',
];

// Install: cache static assets
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        return cache.addAll(STATIC_ASSETS.map(url => {
          return new Request(url, { cache: 'reload' });
        })).catch(err => {
          console.warn('SW: Some assets failed to cache:', err);
        });
      })
      .then(() => self.skipWaiting())
  );
});

// Activate: bersihkan cache lama
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then(cacheNames => {
        return Promise.all(
          cacheNames
            .filter(name => name !== CACHE_NAME)
            .map(name => caches.delete(name))
        );
      })
      .then(() => self.clients.claim())
  );
});

// Fetch: Network-first untuk API, Cache-first untuk static
self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  
  // Jangan cache Microsoft Auth / Graph API requests
  if (url.hostname.includes('microsoft') || 
      url.hostname.includes('microsoftonline') ||
      url.hostname.includes('graph.microsoft') ||
      url.hostname.includes('login.microsoft')) {
    return; // Biarkan network langsung
  }
  
  // Untuk asset lokal: Cache-first dengan network fallback
  if (url.origin === location.origin || 
      url.hostname === 'fonts.googleapis.com' ||
      url.hostname === 'fonts.gstatic.com') {
    event.respondWith(
      caches.match(event.request)
        .then(cached => {
          if (cached) return cached;
          return fetch(event.request)
            .then(response => {
              if (response && response.status === 200) {
                const cloned = response.clone();
                caches.open(CACHE_NAME)
                  .then(cache => cache.put(event.request, cloned));
              }
              return response;
            })
            .catch(() => {
              // Offline fallback untuk navigasi
              if (event.request.mode === 'navigate') {
                return caches.match('/index.html');
              }
            });
        })
    );
    return;
  }
  
  // MSAL CDN: Network-first
  if (url.hostname === 'alcdn.msauth.net') {
    event.respondWith(
      fetch(event.request)
        .then(response => {
          const cloned = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, cloned));
          return response;
        })
        .catch(() => caches.match(event.request))
    );
    return;
  }
});

// Background Sync untuk absen offline (opsional, future feature)
self.addEventListener('sync', (event) => {
  if (event.tag === 'sync-absensi') {
    event.waitUntil(syncPendingAbsensi());
  }
});

async function syncPendingAbsensi() {
  // TODO: implementasi sync offline queue
  console.log('SW: Syncing pending absensi...');
}
