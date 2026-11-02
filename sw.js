// Offline cache for the installed app. The build (vite.config.ts) writes its version and
// file list into this worker, so every deploy changes sw.js: the browser installs the new
// worker, which caches the new files and drops the old cache.
const PREFIX = "ip-";
const VERSION = "1793577886705";
const FILES = ["assets/index-7mlzYRgm.js","assets/index-De2w91Uz.css","assets/index.es-wkJu3uPF.js","assets/purify.es-Dqk5qvad.js"];
const at = (p) => new URL(p, self.location.href).href;
const EXTRA = ["./", "index.html", "manifest.webmanifest", "icons/icon.svg", "icons/apple-touch-icon.png", "icons/icon-192.png", "icons/icon-512.png"];

self.addEventListener("install", (event) => {
  event.waitUntil(
    (async () => {
      const cache = await caches.open(PREFIX + VERSION);
      await cache.addAll([...EXTRA, ...FILES].map((p) => new Request(at(p), { cache: "reload" })));
      await self.skipWaiting();
    })(),
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    (async () => {
      const keys = await caches.keys();
      await Promise.all(keys.filter((k) => k.startsWith(PREFIX) && k !== PREFIX + VERSION).map((k) => caches.delete(k)));
      await self.clients.claim();
    })(),
  );
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET" || new URL(req.url).origin !== self.location.origin) return;
  if (req.mode === "navigate") {
    // Network first for the page, so a new version shows up when online.
    // "no-cache": ask the server every time, so a new version is not hidden by the HTTP cache.
    event.respondWith(
      fetch(req.url, { cache: "no-cache", credentials: "same-origin" }).catch(() => caches.match(at("index.html"), { ignoreVary: true, ignoreSearch: true })),
    );
    return;
  }
  event.respondWith(caches.match(req, { ignoreVary: true }).then((hit) => hit || fetch(req)));
});
