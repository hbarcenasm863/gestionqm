const CACHE = 'hbmol-v2';
const ASSETS = [
  '/gestionqm/',
  '/gestionqm/index.html',
  '/gestionqm/manifest.json',
  '/gestionqm/hbmol.png'
];

// Tiempo máximo que se espera a la red antes de usar el caché como
// respaldo. Suficiente para wifi lenta de colegio, pero no tanto como
// para que la app se sienta congelada si no hay internet.
const NETWORK_TIMEOUT_MS = 3000;

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Red primero, con respaldo de caché: así la PWA instalada (celular o
// computador) siempre trae la última versión publicada cuando hay
// internet, y solo usa lo cacheado si la red falla o tarda demasiado.
self.addEventListener('fetch', e => {
  if (!e.request.url.startsWith(self.location.origin)) return;
  if (e.request.method !== 'GET') return;

  e.respondWith(
    new Promise(resolve => {
      let settled = false;
      const timer = setTimeout(() => {
        if (settled) return;
        settled = true;
        caches.match(e.request).then(cached => {
          resolve(cached || caches.match('/gestionqm/index.html'));
        });
      }, NETWORK_TIMEOUT_MS);

      fetch(e.request).then(res => {
        // Solo se cachean respuestas exitosas — un 404/500 pasajero (p.ej.
        // durante un redeploy) no debe reemplazar el último contenido bueno
        // conocido, o quedaría sirviéndose como "respaldo offline" hasta la
        // próxima vez que la red responda bien.
        if (res.ok) {
          const clone = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, clone));
        }
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        resolve(res);
      }).catch(() => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        caches.match(e.request).then(cached => {
          resolve(cached || caches.match('/gestionqm/index.html'));
        });
      });
    })
  );
});
