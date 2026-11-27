/* 麥穗鍋貼所 Service Worker
   - 網頁本身：網路優先，離線時改用快取（確保一上線就拿到最新版）
   - 圖示、字型、圖片、Three.js：快取優先，背景更新
   更新網站內容時不需要改這裡；若改了快取策略，請把 VERSION 加一。 */
const VERSION = 'v2';
const PAGE_CACHE = `maisui-page-${VERSION}`;
const ASSET_CACHE = `maisui-assets-${VERSION}`;
const PRECACHE = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/apple-touch-icon.png',
  './icons/favicon-32.png',
];
const ASSET_HOSTS = [
  'fonts.googleapis.com',
  'fonts.gstatic.com',
  'images.unsplash.com',
  'i.postimg.cc',
  'cdn.jsdelivr.net',
  'unpkg.com',
];
const MAX_ASSETS = 120;

self.addEventListener('install', (event) => {
  // cache: 'reload'：略過瀏覽器的 HTTP 快取，確保存到的是最新檔案
  event.waitUntil(caches.open(PAGE_CACHE).then((c) => c.addAll(PRECACHE.map((u) => new Request(u, { cache: 'reload' })))).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const keep = [PAGE_CACHE, ASSET_CACHE];
    for (const key of await caches.keys()) if (!keep.includes(key)) await caches.delete(key);
    await self.clients.claim();
  })());
});

async function trim(cache) {
  const keys = await cache.keys();
  for (let i = 0; i < keys.length - MAX_ASSETS; i++) await cache.delete(keys[i]);
}

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // 網頁：網路優先
  if (req.mode === 'navigate') {
    event.respondWith((async () => {
      try {
        // cache: 'no-cache'：每次都向伺服器確認是否有新版（GitHub Pages 預設會讓瀏覽器快取 10 分鐘）
        const res = await fetch(req.url, { cache: 'no-cache', credentials: 'same-origin' });
        const cache = await caches.open(PAGE_CACHE);
        if (res.ok) cache.put('./index.html', res.clone());
        return res;
      } catch {
        return (await caches.match('./index.html')) || (await caches.match('./')) || Response.error();
      }
    })());
    return;
  }

  // 同網站的靜態檔與外部圖片／字型／程式庫：快取優先、背景更新
  const sameOrigin = url.origin === self.location.origin;
  if (sameOrigin || ASSET_HOSTS.includes(url.hostname)) {
    event.respondWith((async () => {
      const cache = await caches.open(ASSET_CACHE);
      const cached = await caches.match(req, { ignoreSearch: sameOrigin });
      const network = fetch(req).then((res) => {
        if (res && (res.ok || res.type === 'opaque')) { cache.put(req, res.clone()); trim(cache); }
        return res;
      }).catch(() => cached || Response.error());
      event.waitUntil(network.then(() => {}, () => {}));
      return cached || network;
    })());
  }
});
