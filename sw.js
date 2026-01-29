const CACHE_NAME = 'licznik-czasu-v1';

self.addEventListener('install', (event) => {
    // Force activation immediately
    self.skipWaiting();
});

self.addEventListener('activate', (event) => {
    // Take control of clients immediately
    event.waitUntil(clients.claim());
});

self.addEventListener('fetch', (event) => {
    // Always go to network, never cache
    // This ensures "no cache" behavior as requested
    event.respondWith(fetch(event.request));
});
