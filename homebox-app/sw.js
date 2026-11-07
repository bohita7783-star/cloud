// HomeBox offline service worker. Scope: this folder only. A worker registered
// for a deeper scope wins over a parent-folder worker, so the parent app's
// cache never serves HomeBox files once this is installed.
// NOTE: this separates CACHING only. IndexedDB is shared by every page of the same
// origin, so for business use host HomeBox on its own origin (see ARCHITECTURE.md).
const CACHE = 'homebox-v0.1.4';
const ASSETS = ['./', './index.html', './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png', './icons/maskable-512.png'];

self.addEventListener('install', e => e.waitUntil(
  caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting())));

self.addEventListener('activate', e => e.waitUntil(caches.keys()
  .then(ks => Promise.all(ks.filter(k => k.startsWith('homebox-') && k !== CACHE).map(k => caches.delete(k))))
  .then(() => self.clients.claim())));

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return;
  if (e.request.mode === 'navigate') {
    // Network first: updates arrive when online; the cached app when offline.
    e.respondWith(fetch(e.request)
      .then(r => { const copy = r.clone(); caches.open(CACHE).then(c => c.put('./index.html', copy)); return r; })
      .catch(() => caches.match('./index.html')));
    return;
  }
  e.respondWith(caches.match(e.request).then(hit => hit || fetch(e.request)));
});
