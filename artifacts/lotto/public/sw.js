// ── SS앱 Service Worker ──────────────────────────────
const CACHE_NAME = "ssapp-build-__PWA_BUILD_ID__";
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
    caches.open(CACHE_NAME).then((cache) => cache.addAll(
      PRECACHE_URLS.map((url) => new Request(new URL(url, self.registration.scope), { cache: "reload" }))
    ))
  );
});

// Keep old assets while other documents may still need them. Natural activation
// with no open windows can safely remove only this app's obsolete caches.
self.addEventListener("activate", (event) => {
  event.waitUntil(
    self.clients.matchAll({ type: "window", includeUncontrolled: true })
      .then((windows) => windows.length ? [] : caches.keys())
      .then((keys) =>
        Promise.all(keys.filter((k) => /^ssapp-(?:v\d+|build-.+)$/.test(k) && k !== CACHE_NAME).map((k) => caches.delete(k)))
      )
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
      networkHtml(request)
    );
    return;
  }

  // JS / CSS / 이미지 / 폰트 → Cache-first
  event.respondWith(cacheFirst(request));
});

async function cacheFirst(request) {
  const cache = await caches.open(CACHE_NAME);
  let cached = await cache.match(request);
  // Old open pages may request their immutable hashed chunks after activation.
  if (!cached && /\/assets\/[^/]+-[^/]+\.(js|css)$/.test(new URL(request.url).pathname)) {
    for (const key of await caches.keys()) {
      if (/^ssapp-(?:v\d+|build-.+)$/.test(key)) {
        cached = await (await caches.open(key)).match(request);
        if (cached) break;
      }
    }
  }
  if (cached) return cached;
  try {
    const response = await fetch(request);
    if (response.ok) {
      await cache.put(request, response.clone());
    }
    return response;
  } catch {
    return cached || new Response("Offline", { status: 503 });
  }
}

async function networkHtml(request) {
  const cache = await caches.open(CACHE_NAME);
  try {
    const response = await fetch(request, { cache: "no-cache" });
    if (response.ok && response.headers.get("content-type")?.includes("text/html")) {
      await cache.put(OFFLINE_URL, response.clone());
    }
    return response;
  } catch {
    return (await cache.match(OFFLINE_URL)) || new Response("Offline", { status: 503 });
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
    event.waitUntil(activateOnRequest(event.source));
  }
});

async function activateOnRequest(requester) {
  if (!requester) return;
  const windows = await self.clients.matchAll({ type: "window", includeUncontrolled: true });
  // Legacy pages reload unconditionally on controllerchange. Do not activate
  // underneath those pages: require their owners to save and close them first.
  const others = windows.filter((client) => client.id !== requester.id);
  const safe = await Promise.all(others.map((client) => new Promise((resolve) => {
    const channel = new MessageChannel();
    const timer = setTimeout(() => { channel.port1.close(); resolve(false); }, 1500);
    channel.port1.onmessage = (message) => {
      clearTimeout(timer);
      channel.port1.close();
      resolve(message.data === "PWA_UPDATE_SAFE");
    };
    client.postMessage({ type: "PWA_UPDATE_SAFETY_CHECK" }, [channel.port2]);
  })));
  if (safe.every(Boolean)) await self.skipWaiting();
  else requester.postMessage({ type: "PWA_UPDATE_BLOCKED" });
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
