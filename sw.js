const CACHE_NAME = 'gp-app-v1';
const ASSETS = ['./index.html', './manifest.json', './notoDevanagariFont.js'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE_NAME).then(c => c.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))))
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  // API कॉल्स नेहमी नेटवर्कवरून; बाकी फाईल्स कॅशेवरून (ऑफलाईन सपोर्टसाठी)
  if (e.request.url.includes('script.google.com')) return;
  e.respondWith(
    caches.match(e.request).then(cached => cached || fetch(e.request))
  );
});
