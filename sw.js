// LHB HR System — Service Worker v1.0
// Handles caching for offline PWA support

const CACHE_NAME = 'lhb-app-v1';
const CACHE_URLS = [
  '/LHB-APP/hr-system.html',
  '/LHB-APP/qr-scan.html',
  '/LHB-APP/food-scan.html',
  '/LHB-APP/manifest.json',
];

// Install — cache core files
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      console.log('[SW] Caching core files');
      return cache.addAll(CACHE_URLS).catch(err => {
        console.warn('[SW] Cache failed:', err);
      });
    })
  );
  self.skipWaiting();
});

// Activate — clean old caches
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
      )
    )
  );
  self.clients.claim();
});

// Fetch — network first, cache fallback
self.addEventListener('fetch', event => {
  const url = event.request.url;

  // Skip Apps Script, Google APIs — always network
  if (url.includes('script.google.com') ||
      url.includes('googleapis.com') ||
      url.includes('api.telegram.org') ||
      url.includes('fonts.googleapis.com') ||
      url.includes('cdnjs.cloudflare.com') ||
      url.includes('accounts.google.com')) {
    return; // Let browser handle
  }

  event.respondWith(
    fetch(event.request)
      .then(response => {
        // Cache successful GET responses
        if (event.request.method === 'GET' && response.ok) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      })
      .catch(() => {
        // Network failed — try cache
        return caches.match(event.request).then(cached => {
          if (cached) return cached;
          // Return offline page for HTML requests
          if (event.request.destination === 'document') {
            return caches.match('/LHB-APP/hr-system.html');
          }
        });
      })
  );
});

// Push notification support (future)
self.addEventListener('push', event => {
  if (!event.data) return;
  const data = event.data.json();
  self.registration.showNotification(data.title || 'LHB HR', {
    body: data.body || '',
    icon: '/LHB-APP/icon-192.png',
    badge: '/LHB-APP/icon-192.png',
    tag: 'lhb-hr-notification',
  });
});
