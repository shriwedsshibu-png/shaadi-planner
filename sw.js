/* Shaadi Planner — makes the app open instantly and work without internet.
   The page itself: fresh copy when online (so updates arrive), saved copy when offline.
   Fonts and icons: saved copy. The Google Sheet link is never stored here (the page keeps its own saved plan). */
var CACHE = 'planner-v4';
var SHELL = ['/', '/index.html', '/manifest.webmanifest', '/icon-192.png', '/icon-512.png', '/icon-maskable-512.png', '/apple-touch-icon.png'];
self.addEventListener('install', function (e) { e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(SHELL); }).then(function () { return self.skipWaiting(); })); });
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (ks) { return Promise.all(ks.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); })); }).then(function () { return self.clients.claim(); }));
});
self.addEventListener('fetch', function (e) {
  var req = e.request, url = new URL(req.url);
  if (req.method !== 'GET' || /script\.google(usercontent)?\.com$/.test(url.hostname)) return;   // sheet requests go straight to Google
  if (req.mode === 'navigate' || (url.origin === location.origin && /\/(index\.html)?$/.test(url.pathname))) {
    e.respondWith(Promise.race([
      fetch(req).then(function (r) { var cp = r.clone(); caches.open(CACHE).then(function (c) { c.put('/', cp); }); return r; }),
      new Promise(function (res) { setTimeout(res, 3500); })      // slow network: show the saved page after 3.5 s
    ]).then(function (r) { return r || caches.match('/'); }).catch(function () { return caches.match('/'); }));
    return;
  }
  if (url.origin === location.origin || /fonts\.(googleapis|gstatic)\.com$/.test(url.hostname)) {
    e.respondWith(caches.match(req).then(function (hit) {
      var net = fetch(req).then(function (r) { if (r && (r.ok || r.type === 'opaque')) { var cp = r.clone(); caches.open(CACHE).then(function (c) { c.put(req, cp); }); } return r; }).catch(function () { return hit; });
      return hit || net;
    }));
  }
});
