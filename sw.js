// 離線快取。先用手機裡的版本(在倉庫訊號差也能立刻打開),有網路時背景更新。
const CACHE = 'dwt-951fc5d7e246';
const SHELL = ["icon-180.png","icon-192.png","icon-512.png","./","manifest.webmanifest"];
self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET' || new URL(e.request.url).origin !== self.location.origin) return;
  const key = e.request.mode === 'navigate' ? './' : e.request;
  e.respondWith(caches.match(key).then((hit) => hit || fetch(e.request)));
});
