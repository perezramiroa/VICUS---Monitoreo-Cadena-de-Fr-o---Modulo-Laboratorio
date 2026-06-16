const CACHE_NAME = 'rsamio-v1.1.4';
const PREFIX = '/Monitoreo-Cadena-de-Fr-o---Farmacia';
const urlsToCache = [
  `${PREFIX}/`,
  `${PREFIX}/index.html`,
  `${PREFIX}/Vicus_farmacia.html`,
  `${PREFIX}/reporte_individual_farmacia.html`,
  `${PREFIX}/css/pwa-styles.css`,
  `${PREFIX}/js/pwa-utils.js`,
  `${PREFIX}/logos/logo_rih.jpg`,
  `${PREFIX}/logos/footer.jpg`,
  `${PREFIX}/logos/portal-vicus.png`
];

// Instalación del Service Worker
self.addEventListener('install', event => {
  console.log('[SW] Instalando nueva versión...');
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => {
        console.log('[SW] Pre-cacheando archivos esenciales...');
        return cache.addAll(urlsToCache);
      })
      .then(() => self.skipWaiting())
  );
});

// Activación y limpieza de caches antiguas
self.addEventListener('activate', event => {
  console.log('[SW] Activando nueva versión...');
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cacheName => {
          if (cacheName !== CACHE_NAME) {
            console.log('[SW] Eliminando cache antigua:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// Estrategia de peticiones mejorada
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // EXCLUIR Google Apps Script y ThingSpeak de la interceptación del SW
  // Estas peticiones deben ir directo a la red para evitar problemas de CORS
  if (url.hostname.includes('script.google.com') || 
      url.hostname.includes('script.googleusercontent.com') ||
      url.hostname.includes('api.thingspeak.com')) {
    console.log('[SW] Ignorando petición externa (CORS bypass):', url.hostname);
    return; // Dejar que el navegador maneje la petición normalmente
  }
  
  // 1. Estrategia NETWORK FIRST para archivos HTML (asegura última versión)
  if (event.request.mode === 'navigate' || url.pathname.endsWith('.html')) {
    event.respondWith(
      fetch(event.request)
        .then(response => {
          // Si la red responde, guardamos en caché y devolvemos
          const responseClone = response.clone();
          caches.open(CACHE_NAME).then(cache => {
            cache.put(event.request, responseClone);
          });
          return response;
        })
        .catch(() => {
          // Si falla la red, usamos la caché
          return caches.match(event.request);
        })
    );
    return;
  }

  // 2. Estrategia STALE-WHILE-REVALIDATE para el resto (CSS, JS, Imágenes)
  event.respondWith(
    caches.match(event.request)
      .then(cachedResponse => {
        const fetchPromise = fetch(event.request).then(networkResponse => {
          // Solo cachear si es exitoso y es GET
          if (networkResponse.status === 200 && event.request.method === 'GET') {
            const responseClone = networkResponse.clone();
            caches.open(CACHE_NAME).then(cache => {
              cache.put(event.request, responseClone);
            });
          }
          return networkResponse;
        });

        // Devolver la caché inmediatamente si existe, si no esperar a la red
        return cachedResponse || fetchPromise;
      }).catch(() => {
        // Fallback básico para imágenes si falla todo
        if (event.request.destination === 'image') {
          return caches.match(`${PREFIX}/logos/portal-vicus.png`);
        }
      })
  );
});

// Escuchar mensajes para forzar actualización
self.addEventListener('message', event => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});
