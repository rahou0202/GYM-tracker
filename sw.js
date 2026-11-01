// Keeps a copy of the app on the phone so it opens offline (at the gym with no signal).
// Bump VERSION whenever the app files change, so phones pick up the new version.
const VERSION = "gym-tracker-v1";
const FILES = ["./", "index.html", "local-api.js", "xlsx.full.min.js", "manifest.webmanifest",
               "icon-192.png", "icon-512.png", "apple-touch-icon.png"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys()
    .then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener("fetch", (e) => {
  if (e.request.method !== "GET") return;
  const url = new URL(e.request.url);
  const isFont = url.hostname === "fonts.googleapis.com" || url.hostname === "fonts.gstatic.com";
  if (url.origin !== location.origin && !isFont) return;

  // open instantly from the saved copy (even with weak signal at the gym);
  // app files are refreshed in the background, so updates show up the next time the app opens
  e.respondWith(caches.match(e.request, { ignoreSearch: true }).then((hit) => {
    if (hit && isFont) return hit; // fonts never change
    const fresh = fetchAndKeep(e.request);
    if (hit) {
      fresh.catch(() => {});
      return hit;
    }
    return fresh.catch(() => caches.match("index.html"));
  }));
});

function fetchAndKeep(request) {
  return fetch(request).then((res) => {
    if (res.ok || res.type === "opaque") {
      const copy = res.clone();
      caches.open(VERSION).then((c) => c.put(request, copy));
    }
    return res;
  });
}
