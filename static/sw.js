const CACHE_NAME = 'FeeRemind-v3';

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll([
        '/dashboard',
        '/login',
        '/static/banner.jpg'
      ]))
      .catch((error) => console.error('Cache error:', error))
  );
});

self.addEventListener('fetch', (event) => {
  console.log('Fetching:', event.request.url);
  event.respondWith(
    caches.match(event.request)
      .then((response) => response || fetch(event.request))
  );
});