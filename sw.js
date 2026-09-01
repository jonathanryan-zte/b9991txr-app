const CACHE = "b9991txr-shell-v2";
const SHELL = [
  "./",
  "./index.html",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png"
];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)));
  self.skipWaiting();
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (e) => {
  const url = new URL(e.request.url);

  // Jangan pernah sentuh Apps Script API atau CDN eksternal (Tesseract.js) —
  // itu harus selalu langsung ke jaringan, tidak boleh di-cache di sini.
  if (url.origin !== self.location.origin) return;
  if (e.request.method !== "GET") return;

  // HTML diambil dengan cache: "no-store" supaya HTTP cache Safari tidak
  // menyajikan index.html versi lama. Pernah terjadi: HP masih menjalankan
  // build lama berminggu-minggu (Invoice masih wajib diisi) walaupun versi
  // baru sudah lama tayang di GitHub Pages.
  const isHtml = e.request.mode === "navigate" ||
    e.request.destination === "document" ||
    url.pathname.endsWith(".html") || url.pathname.endsWith("/");

  e.respondWith(
    fetch(e.request, isHtml ? { cache: "no-store" } : undefined)
      .then((res) => {
        const copy = res.clone();
        caches.open(CACHE).then((c) => c.put(e.request, copy));
        return res;
      })
      .catch(() => caches.match(e.request))
  );
});
