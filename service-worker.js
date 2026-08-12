// Service worker di Jarvis.
// v2: usa strategia "network-first" per l'HTML, cosi' quando aggiorni
// index.html su GitHub il telefono prende SEMPRE la versione nuova
// (prima cache), e usa la cache solo come riserva se non c'e' internet.
const CACHE_NAME = "jarvis-cache-v2";
const ASSETS = [
  "./",
  "./index.html",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png",
  "./apple-touch-icon.png"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS))
  );
  self.skipWaiting(); // attiva subito la nuova versione, non aspetta la chiusura dell'app
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim(); // prende subito il controllo delle pagine aperte
});

self.addEventListener("fetch", (event) => {
  const req = event.request;

  // Per la pagina HTML: prova sempre la rete per prima cosa (versione fresca).
  // Se non c'e' internet, usa la copia salvata come riserva.
  if (req.mode === "navigate" || req.destination === "document") {
    event.respondWith(
      fetch(req)
        .then((res) => {
          const clone = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(req, clone));
          return res;
        })
        .catch(() => caches.match(req).then((c) => c || caches.match("./index.html")))
    );
    return;
  }

  // Per tutto il resto (icone, manifest...): cache prima, rete come riserva.
  event.respondWith(
    caches.match(req).then((cached) => {
      return (
        cached ||
        fetch(req)
          .then((response) => {
            if (response.ok && req.url.startsWith(self.location.origin)) {
              const clone = response.clone();
              caches.open(CACHE_NAME).then((cache) => cache.put(req, clone));
            }
            return response;
          })
          .catch(() => cached)
      );
    })
  );
});
