const CACHE = 'travel-v37';
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
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = req.url;

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
      }).catch(() => caches.match(req).then(r => r || fetch(req)))
    );
    return;
  }

  // 其他所有資源（頁面、Sheets API、圖片、天氣/行事曆 CDN）：
  // network first，成功就存快取；離線或請求失敗時改用快取，
  // 讓「連線時開過的內容」離線也能看。cache.put 等 await 完才算
  // fetch 事件結束，避免 SW 被提前中止導致快取沒存到。
  e.respondWith((async () => {
    try {
      const res = await fetch(req);
      if (res && (res.ok || res.type === 'opaque')) {
        const cache = await caches.open(CACHE);
        await cache.put(req, res.clone());
      }
      return res;
    } catch (err) {
      const cached = await caches.match(req);
      if (cached) return cached;
      throw err;
    }
  })());
});
