// 離線快取。先用手機裡的版本(在倉庫訊號差也能立刻打開),有網路時背景更新。
const CACHE = 'dwt-951fc5d7e246';
const SHELL = ["icon-180.png","icon-192.png","icon-512.png","./","manifest.webmanifest"];
self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
// 只處理手機版自己的檔案。同一個網站底下還有電腦版看板(tracker/),不能被這裡攔下來換成手機版。
const SCOPE_PATH = new URL(self.registration.scope).pathname;
self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== self.location.origin || !url.pathname.startsWith(SCOPE_PATH)) return;
  const rel = url.pathname.slice(SCOPE_PATH.length);
  const isPage = rel === '' || rel === 'index.html';
  if (!isPage && SHELL.indexOf(rel) === -1) return;
  e.respondWith(caches.match(isPage ? './' : rel).then((hit) => hit || fetch(e.request)));
});
