const CACHE_NAME = "inventario-ti-v22";
const ASSETS = ["./", "./index.html", "./manifest.json", "./icon-192.png", "./icon-512.png"];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS)).catch(() => {})
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;

  // O backend (Google Apps Script) serve dados dinâmicos (login, inventário,
  // chamados) — nunca deve ser cacheado. Se ficasse em cache, uma resposta
  // antiga (inclusive de erro) continuaria sendo servida para sempre, mesmo
  // depois do backend ser corrigido ou os dados mudarem.
  // Domínios *.googleapis.com (Firestore, Auth) também ficam de fora: a
  // conexão "Listen" do Firestore (sincronização ao vivo dos chamados) é uma
  // resposta contínua (stream), não uma resposta normal de uma vez só —
  // tentar interceptar/clonar ela pra cachear quebra a conexão (erro
  // "Response body is already used" no console) e derrubava o listener ao
  // vivo, mesmo sem nenhum benefício em cachear algo que nunca deveria ser
  // reaproveitado.
  const hostname = new URL(event.request.url).hostname;
  if (hostname === "script.google.com" || hostname.endsWith(".googleapis.com")) return;

  event.respondWith(
    caches.match(event.request).then((cached) => {
      const fetchPromise = fetch(event.request)
        .then((networkResponse) => {
          // Cross-origin CDN scripts (react, chart.js, xlsx...) come back as
          // opaque responses (status 0, ok === false) since they're loaded
          // without CORS, so opaque responses are cached too — otherwise
          // those libraries would never be cached and get re-downloaded on
          // every visit.
          if (networkResponse && (networkResponse.ok || networkResponse.type === "opaque")) {
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, networkResponse.clone()));
          }
          return networkResponse;
        })
        .catch(() => cached);
      return cached || fetchPromise;
    })
  );
});
