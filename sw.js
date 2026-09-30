// Minimal service worker -- exists mainly so the browser treats the kiosk
// as an installable app (Chrome/Edge require an active SW with a fetch
// handler before showing the "Install" prompt). It caches the app shell so
// the kiosk still opens if wifi briefly drops, then gets out of the way --
// everything else (Supabase, fonts, mailer) goes straight to the network.

const CACHE = 'flexi-kiosk-v1';
const SHELL = ['./', './index.html', './manifest.json'];

self.addEventListener('install', (e) => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).catch(() => {}));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
  );
  self.clients.claim();
});

self.addEventListener('fetch', (e) => {
  // Only handle page navigations (the kiosk's own HTML) -- everything else
  // (Supabase calls, CDN scripts, fonts) passes straight through untouched.
  if (e.request.mode === 'navigate') {
    e.respondWith(
      fetch(e.request).catch(() => caches.match('./index.html'))
    );
  }
});
