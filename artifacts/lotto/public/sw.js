// ── SS앱 Service Worker ──────────────────────────────
const CACHE_NAME = "ssapp-v6";
const OFFLINE_URL = "./";

// 설치 시 캐시할 핵심 리소스
const PRECACHE_URLS = [
  "./",
  "./manifest.json",
  "./icon.png",
];

// ── Install: 핵심 리소스 사전 캐싱 ──────────────────────────
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(PRECACHE_URLS))
  );
  self.skipWaiting();
});

// ── Activate: 오래된 캐시 삭제 ──────────────────────────────
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((k) => k !== CACHE_NAME)
          .map((k) => caches.delete(k))
      )
    )
  );
  self.clients.claim();
});

// ── Fetch: 캐시 전략 분기 ────────────────────────────────────
self.addEventListener("fetch", (event) => {
  const { request } = event;
  const url = new URL(request.url);

  // API 요청 → Network-first (실패 시 캐시 fallback)
  if (
    request.method !== "GET" ||
    url.pathname.includes("/api/") ||
    url.hostname !== self.location.hostname
  ) {
    event.respondWith(networkFirst(request));
    return;
  }

  // HTML 네비게이션 → Network-first (오프라인 시 기본 화면)
  if (request.mode === "navigate") {
    event.respondWith(
      fetch(request).catch(() =>
        caches.match(OFFLINE_URL).then((r) => r || caches.match("./"))
      )
    );
    return;
  }

  // JS / CSS / 이미지 / 폰트 → Cache-first
  event.respondWith(cacheFirst(request));
});

async function cacheFirst(request) {
  const cached = await caches.match(request);
  if (cached) return cached;
  try {
    const response = await fetch(request);
    if (response.ok) {
      const cache = await caches.open(CACHE_NAME);
      cache.put(request, response.clone());
    }
    return response;
  } catch {
    return cached || new Response("Offline", { status: 503 });
  }
}

async function networkFirst(request) {
  try {
    const response = await fetch(request);
    if (response.ok) {
      const cache = await caches.open(CACHE_NAME);
      cache.put(request, response.clone());
    }
    return response;
  } catch {
    const cached = await caches.match(request);
    return cached || new Response("Offline", { status: 503 });
  }
}

// ── Push 알림 준비 ───────────────────────────────────────────
self.addEventListener("push", (event) => {
  const data = event.data?.json() ?? { title: "Check:Bite", body: "새 알림" };
  event.waitUntil(
    self.registration.showNotification(data.title, {
      body: data.body,
      icon: "./icons/icon-192.png",
      badge: "./icons/icon-192.png",
    })
  );
});

self.addEventListener("notificationclick", (event) => {
  event.notification.close();
  event.waitUntil(
    clients.matchAll({ type: "window" }).then((list) => {
      if (list.length) return list[0].focus();
      return clients.openWindow("./");
    })
  );
});
