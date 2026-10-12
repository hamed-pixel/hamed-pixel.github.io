// Network first, cache as fallback: staff always get the newest version, and the app still opens offline.
const CACHE = 'smartech-stock-v1';
self.addEventListener('install', e => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET') return;
  // Firebase data traffic manages itself; only cache the app files and libraries.
  if (u.hostname.endsWith('googleapis.com') && !u.hostname.startsWith('fonts')) return;
  e.respondWith(fetch(e.request).then(r => {
    if (r.ok || r.type === 'opaque') { const c = r.clone(); caches.open(CACHE).then(k => k.put(e.request, c)); }
    return r;
  }).catch(() => caches.match(e.request)));
});
