// ── SS앱 Service Worker ──────────────────────────────
const CACHE_NAME = "ssapp-v13";
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

// ── Activate: 모든 기존 캐시 강제 삭제 후 핵심 리소스 재캐싱 ─────
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) =>
        // 현재 버전 포함 모든 캐시 삭제 → 구버전 파일 완전 제거
        Promise.all(keys.map((k) => caches.delete(k)))
      )
      .then(() =>
        // 삭제 후 핵심 리소스만 새로 캐싱
        caches.open(CACHE_NAME).then((cache) => cache.addAll(PRECACHE_URLS))
      )
      .then(() => self.clients.claim())
  );
});

// ── Fetch: 캐시 전략 분기 ────────────────────────────────────
self.addEventListener("fetch", (event) => {
  const { request } = event;
  const url = new URL(request.url);

  // API 요청 → Network-only (절대 캐시하지 않음)
  if (
    request.method !== "GET" ||
    url.pathname.includes("/api/") ||
    url.hostname !== self.location.hostname
  ) {
    event.respondWith(networkOnly(request));
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

// API 전용: 캐시 저장 없이 항상 네트워크 직접 호출
async function networkOnly(request) {
  try {
    return await fetch(request);
  } catch {
    return new Response(JSON.stringify({ error: "Offline" }), {
      status: 503,
      headers: { "Content-Type": "application/json" },
    });
  }
}

// ── 페이지에서 SKIP_WAITING 메시지 수신 시 즉시 활성화 ───────
self.addEventListener("message", (event) => {
  if (event.data && event.data.type === "SKIP_WAITING") {
    self.skipWaiting();
  }
});

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
