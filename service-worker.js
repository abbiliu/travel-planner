const CACHE = 'travel-v8';
const SHELL = [
  './index.html',
  './manifest.json',
  './scene_01.jpg',
  './scene_02.jpg',
  './scene_03.jpg',
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const url = e.request.url;

  // APIs: network first, 離線時用快取
  if (url.includes('googleapis.com') || url.includes('accounts.google.com')) {
    e.respondWith(
      fetch(e.request)
        .then(r => { caches.open(CACHE).then(c => c.put(e.request, r.clone())); return r; })
        .catch(() => caches.match(e.request))
    );
    return;
  }

  // index.html: network first，確保每次都拿最新版
  if (url.endsWith('/') || url.includes('index.html') || !url.includes('.')) {
    e.respondWith(
      fetch(e.request)
        .then(r => { caches.open(CACHE).then(c => c.put(e.request, r.clone())); return r; })
        .catch(() => caches.match(e.request))
    );
    return;
  }

  // 其他靜態資源: cache first
  e.respondWith(caches.match(e.request).then(r => r || fetch(e.request)));
});
