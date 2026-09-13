// Service worker mínimo, sin caché.
// No guarda nada offline a propósito: así el sitio nunca puede quedarse
// "atascado" sirviendo una versión vieja o una descarga a medias tras
// una actualización. El manifest.json sigue permitiendo "Añadir a
// pantalla de inicio" con normalidad.

self.addEventListener('install', () => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// No interceptamos fetch: cada petición va directa a la red,
// como si no hubiera service worker.
