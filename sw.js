/* แคชตัวแอปไว้ในเครื่อง เปิดได้ทันที ข้อมูลดึงจากเซิร์ฟเวอร์ทุกครั้ง */
var V = 'incident-v4';
var SHELL = ['./', 'index.html', 'config.js', 'manifest.webmanifest', 'apple-touch-icon.png', 'icon-192.png', 'icon-512.png', 'favicon-32.png'];
self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(V).then(function (c) { return c.addAll(SHELL); }).then(function () { return self.skipWaiting(); }));
});
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (ks) {
    return Promise.all(ks.filter(function (k) { return k !== V; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});
self.addEventListener('fetch', function (e) {
  var req = e.request, url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== self.location.origin) return;
  if (req.mode === 'navigate' || /\/(index\.html|config\.js)?$/.test(url.pathname)) {
    e.respondWith(fetch(req).then(function (r) {
      var cp = r.clone(); caches.open(V).then(function (c) { c.put(req, cp); }); return r;
    }).catch(function () { return caches.match(req).then(function (r) { return r || caches.match('index.html'); }); }));
    return;
  }
  e.respondWith(caches.match(req).then(function (r) { return r || fetch(req); }));
});
