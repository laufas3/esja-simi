// Geymir aðeins viðmótið (kóðann) svo síðan opnist án nets. Engin gögn eru geymd hér:
// köll á Microsoft Graph og innskráningu fara framhjá.
const C = 'esja-simi-v1';
const SKEL = ['./', 'index.html', 'app.js', 'simi.css', 'stilar.css', 'stillingar.js',
              'msal-browser.min.js', 'manifest.webmanifest', 'takn-180.png'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(C).then(c => c.addAll(SKEL)));
  self.skipWaiting();
});
self.addEventListener('activate', e => e.waitUntil(
  caches.keys().then(k => Promise.all(k.filter(n => n !== C).map(n => caches.delete(n))))
    .then(() => self.clients.claim())));
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  e.respondWith(fetch(e.request).then(r => {
    const k = r.clone(); caches.open(C).then(c => c.put(e.request, k)); return r;
  }).catch(() => caches.match(e.request)));
});
