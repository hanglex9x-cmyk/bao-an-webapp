var CACHE = 'baoan-demo-v1';
var SHELL = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png', './apple-touch-icon.png'];
self.addEventListener('install', function(e){ e.waitUntil(caches.open(CACHE).then(function(c){ return c.addAll(SHELL); }).then(function(){ return self.skipWaiting(); })); });
self.addEventListener('activate', function(e){ e.waitUntil(caches.keys().then(function(ks){ return Promise.all(ks.map(function(k){ if(k !== CACHE) return caches.delete(k); })); }).then(function(){ return self.clients.claim(); })); });
self.addEventListener('fetch', function(e){
  var url = new URL(e.request.url);
  if (url.origin !== self.location.origin) return; // API va tai nguyen ngoai: luon qua mang
  if (e.request.mode === 'navigate') {
    e.respondWith(fetch(e.request).then(function(r){ var cp = r.clone(); caches.open(CACHE).then(function(c){ c.put('./index.html', cp); }); return r; }).catch(function(){ return caches.match('./index.html'); }));
    return;
  }
  e.respondWith(caches.match(e.request).then(function(m){ return m || fetch(e.request); }));
});
