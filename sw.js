'use strict';

const CACHE_NAME = 'pdf2ppt-ocr-v4';

// Transformers.js v3 本体(cdn.jsdelivr.net/@huggingface)と
// manga-ocr-base モデルファイル(huggingface.co/onnx-community)をキャッシュ対象とする
const CACHE_ORIGINS = [
  'cdn.jsdelivr.net/npm/@huggingface/transformers',
  'huggingface.co/onnx-community',
  'huggingface.co/datasets/onnx-community'
];

self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));

self.addEventListener('fetch', e => {
  if (!CACHE_ORIGINS.some(o => e.request.url.includes(o))) return;

  e.respondWith(
    caches.open(CACHE_NAME).then(cache =>
      cache.match(e.request).then(cached => {
        if (cached) return cached;
        return fetch(e.request).then(res => {
          if (res.ok) cache.put(e.request, res.clone());
          return res;
        });
      })
    )
  );
});
