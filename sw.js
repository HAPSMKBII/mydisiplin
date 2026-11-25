// MyDisiplin — service worker
// Naikkan VERSI setiap kali index.html dikemas kini.
const VERSI = 'mydisiplin-v1.0.6';
const WAJIB = ['./index.html'];
const PILIHAN = ['./', './manifest.webmanifest', './tetapan.js', './lencana.png', './ikon-192.png', './ikon-512.png',
  './ikon-maskable-512.png', './apple-touch-icon.png', './favicon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSI)
    .then(c => c.addAll(WAJIB).then(() => Promise.all(PILIHAN.map(f => c.add(f).catch(() => {})))))
    .then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k !== VERSI).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return; // Apps Script, fon: terus ke rangkaian
  if (req.mode === 'navigate' || url.pathname.endsWith('/tetapan.js')) {
    const kunci = req.mode === 'navigate' ? './index.html' : './tetapan.js';
    e.respondWith(fetch(req).then(r => {
      if (r.ok) { const k = r.clone(); caches.open(VERSI).then(c => c.put(kunci, k)); }
      return r;
    }).catch(() => caches.match(kunci)));
    return;
  }
  e.respondWith(caches.match(req).then(r => r || fetch(req)));
});
