const DEVICE_ID_KEY = "lotto_device_id";

export interface DeviceBinding {
  userName: string;
  registeredAt: string;
}

export interface DeviceAccessState {
  enabled: boolean;
  bindings: DeviceBinding[];
}

function createDeviceId(): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `${Date.now()}-${Math.random().toString(36).slice(2)}-${Math.random().toString(36).slice(2)}`;
}

export function getDeviceId(): string {
  const existing = localStorage.getItem(DEVICE_ID_KEY);
  if (existing) return existing;
  const created = createDeviceId();
  localStorage.setItem(DEVICE_ID_KEY, created);
  return created;
}

export async function checkDeviceAccess(userName: string): Promise<{ allowed: boolean; reason?: string }> {
  try {
    const response = await fetch("/api/device-access/check", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ userName, deviceId: getDeviceId() }),
    });
    const data = await response.json() as { allowed?: boolean; reason?: string };
    return {
      allowed: response.ok && data.allowed === true,
      reason: data.reason ?? (response.ok ? undefined : "기기 등록을 확인할 수 없습니다."),
    };
  } catch {
    return { allowed: false, reason: "서버에 연결할 수 없어 기기 등록을 확인하지 못했습니다." };
  }
}

export async function getDeviceAccessState(): Promise<DeviceAccessState> {
  const response = await fetch("/api/device-access", { cache: "no-store" });
  if (!response.ok) throw new Error("기기 제한 상태를 불러오지 못했습니다.");
  return response.json() as Promise<DeviceAccessState>;
}

export async function setDeviceAccessEnabled(enabled: boolean): Promise<void> {
  const response = await fetch("/api/device-access/settings", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ enabled }),
  });
  if (!response.ok) throw new Error("기기 제한 설정을 저장하지 못했습니다.");
}

export async function releaseDevice(userName: string): Promise<void> {
  const response = await fetch(`/api/device-access/${encodeURIComponent(userName)}`, { method: "DELETE" });
  if (!response.ok) throw new Error("기기 연결을 해제하지 못했습니다.");
}
