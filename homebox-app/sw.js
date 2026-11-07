// HomeBox offline service worker. Scope: this folder only. A worker registered
// for a deeper scope wins over a parent-folder worker, so the parent app's
// cache never serves HomeBox files once this is installed.
// NOTE: this separates CACHING only. IndexedDB is shared by every page of the same
// origin, so for business use host HomeBox on its own origin (see ARCHITECTURE.md).
const CACHE = 'homebox-v0.3.0';
// Pyodide (about 19 MB) lives in its own cache named after its version, so app
// updates do not download it again. Files under ./pyodide/<version>/ never change.
const PYODIDE_CACHE = 'homebox-pyodide-0.27.7';
const ASSETS = ['./', './index.html', './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png', './icons/maskable-512.png',
  './engine-worker.js', './engine/engine_policy.py', './engine/approval.py', './engine/converter_mobile.py',
  './engine/homebox_adapter.py', './engine/case_exports.py', './engine/report_layout.py', './trusted_approvers.json'];

self.addEventListener('install', e => e.waitUntil(
  caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting())));

self.addEventListener('activate', e => e.waitUntil(caches.keys()
  .then(ks => Promise.all(ks.filter(k => k.startsWith('homebox-') && k !== CACHE && k !== PYODIDE_CACHE).map(k => caches.delete(k))))
  .then(() => self.clients.claim())));

async function networkFirst(req, cacheName, key) {
  try {
    const r = await fetch(req);
    if (r.ok) { const copy = r.clone(); caches.open(cacheName).then(c => c.put(key || req, copy)); }
    return r;
  } catch (_) {
    return (await caches.match(key || req, { ignoreSearch: true })) || Response.error();
  }
}

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return;
  if (e.request.mode === 'navigate') return e.respondWith(networkFirst(e.request, CACHE, './index.html'));
  // The approver key list is the trust root: always try the latest from the origin first.
  if (url.pathname.endsWith('/trusted_approvers.json')) return e.respondWith(networkFirst(e.request, CACHE));
  if (url.pathname.includes('/pyodide/')) {
    return e.respondWith(caches.open(PYODIDE_CACHE).then(async c => {
      const hit = await c.match(e.request);
      if (hit) return hit;
      const r = await fetch(e.request);
      if (r.ok) c.put(e.request, r.clone());
      return r;
    }));
  }
  e.respondWith(caches.match(e.request).then(hit => hit || fetch(e.request)));
});
