// ============================================================
// SSKYMT Analytics — Service Worker v5.0.0 (PWA)
// Offline-first + Cache + Real-time sync when online
// ============================================================

const CACHE_NAME = 'sskymt-analytics-v5.0.0';
const OFFLINE_URL = './'; // fallback to main app

// App shell — critical resources
const PRECACHE = [
  './',
  './?page=index',
  // Add more if separate CSS/JS, but everything is inlined in index.html
];

// Install
self.addEventListener('install', (event) => {
  console.log('[SW] Installing SSKYMT PWA...');
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => {
        console.log('[SW] Precaching app shell');
        return cache.addAll(PRECACHE).catch(err => {
          console.warn('[SW] Precache partial fail (normal on first install):', err);
        });
      })
      .then(() => self.skipWaiting())
  );
});

// Activate — clean old caches
self.addEventListener('activate', (event) => {
  console.log('[SW] Activating...');
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.filter(k => k !== CACHE_NAME).map(k => {
          console.log('[SW] Deleting old cache:', k);
          return caches.delete(k);
        })
      );
    }).then(() => self.clients.claim())
  );
});

// Fetch strategy: Network first for API, Cache first for shell
self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // Skip non-GET and chrome-extension etc.
  if (event.request.method !== 'GET') return;

  // Skip Google Apps Script internal / google.script.run (always network)
  if (url.hostname.includes('script.google.com') || 
      url.hostname.includes('googleusercontent.com') ||
      url.pathname.includes('/exec') ||
      url.search.includes('google.script')) {
    // Network only for live data
    event.respondWith(fetch(event.request).catch(() => {
      return new Response(JSON.stringify({ success: false, error: 'Offline — ข้อมูลออนไลน์ไม่พร้อม' }), {
        headers: { 'Content-Type': 'application/json' }
      });
    }));
    return;
  }

  // For navigation / HTML — Network first, fallback to cache
  if (event.request.mode === 'navigate' || 
      (event.request.headers.get('accept') || '').includes('text/html')) {
    event.respondWith(
      fetch(event.request)
        .then((response) => {
          // Clone and cache successful responses
          if (response && response.status === 200) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
          }
          return response;
        })
        .catch(() => {
          return caches.match(event.request).then(cached => {
            return cached || caches.match(OFFLINE_URL) || new Response(
              '<!DOCTYPE html><html><head><meta charset="utf-8"><title>Offline</title></head><body style="background:#060810;color:#f4f6fc;font-family:sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;text-align:center;"><div><h1>⚡ SSKYMT</h1><p>คุณออฟไลน์อยู่<br>กรุณาเชื่อมต่ออินเทอร์เน็ตเพื่อใช้งานข้อมูลล่าสุด</p><button onclick="location.reload()" style="padding:12px 24px;background:#4d8bff;color:white;border:none;border-radius:8px;font-size:16px;cursor:pointer;">ลองใหม่</button></div></body></html>',
              { headers: { 'Content-Type': 'text/html; charset=utf-8' } }
            );
          });
        })
    );
    return;
  }

  // Static assets — Cache first
  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached;
      return fetch(event.request).then((response) => {
        if (response && response.status === 200 && response.type === 'basic') {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      }).catch(() => cached);
    })
  );
});

// Background sync placeholder (for future offline queue)
self.addEventListener('sync', (event) => {
  if (event.tag === 'sskymt-sync') {
    console.log('[SW] Background sync triggered');
    // Future: flush offline queue
  }
});

// Push notification support (optional, for future LINE-like alerts)
self.addEventListener('push', (event) => {
  const data = event.data ? event.data.json() : { title: 'SSKYMT', body: 'มีอัพเดทใหม่' };
  event.waitUntil(
    self.registration.showNotification(data.title || 'SSKYMT Analytics', {
      body: data.body || '',
      icon: './?page=icon192',
      badge: './?page=icon192',
      vibrate: [100, 50, 100]
    })
  );
});

self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  event.waitUntil(clients.openWindow('./'));
});

console.log('[SW] SSKYMT Service Worker loaded — v5.0.0');
