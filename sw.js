// Service worker mínimo: hace el panel instalable. Siempre va a la red (los leads tienen que estar al día).
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', () => {});
