const CACHE = 'travel-v14';
const SHARE_CACHE = 'pwa-share-v1';
const SHELL = [
  './index.html',
  './manifest.json',
  './scene_01.jpg',
  './scene_02.jpg',
  './scene_03.jpg',
  './countries.json',
  './flight.svg',
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE && k !== SHARE_CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Page tells SW which share token is active (or clears it)
self.addEventListener('message', event => {
  if (!event.data) return;
  caches.open(SHARE_CACHE).then(c => {
    if (event.data.type === 'SET_SHARE_TOKEN') {
      c.put('/_share_token', new Response(event.data.token, { headers: { 'Content-Type': 'text/plain' } }));
    } else if (event.data.type === 'CLEAR_SHARE_TOKEN') {
      c.delete('/_share_token');
    }
  });
});

self.addEventListener('fetch', e => {
  const url = e.request.url;

  // Dynamic manifest: inject share token so iOS uses correct start_url
  if (url.includes('manifest.json')) {
    e.respondWith(
      caches.open(SHARE_CACHE).then(async shareCache => {
        const tokenResp = await shareCache.match('/_share_token');
        const token = tokenResp ? await tokenResp.text() : null;
        const manifestResp = await fetch('./manifest.json');
        if (!token) return manifestResp;
        const manifest = await manifestResp.json();
        manifest.start_url = './index.html?share=' + encodeURIComponent(token);
        return new Response(JSON.stringify(manifest), {
          headers: { 'Content-Type': 'application/manifest+json' }
        });
      }).catch(() => caches.match(e.request).then(r => r || fetch(e.request)))
    );
    return;
  }

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
