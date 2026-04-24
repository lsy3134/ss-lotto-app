// ─────────────────────────────────────────────────────────
// 허용 사용자 목록 — 이름: 권한("admin" | "user")
// 관리자(admin): 근무표 편집, 휴무 업로드, 순번표 편집 가능
// 일반(user)   : 근무표 조회만 가능
// ─────────────────────────────────────────────────────────
export const ALLOWED_USERS: Record<string, "admin" | "user"> = {
  "써니":   "admin",
  "관리자":  "admin",
};

export type Role = "admin" | "user";

export interface AuthUser {
  name: string;
  role: Role;
}

const STORAGE_KEY = "lotto_auth_user";

export function getStoredUser(): AuthUser | null {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as AuthUser;
    // 저장된 이름이 여전히 허용 목록에 있는지 재검증
    if (!ALLOWED_USERS[parsed.name]) return null;
    return parsed;
  } catch {
    return null;
  }
}

export function saveUser(user: AuthUser) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(user));
}

export function clearUser() {
  localStorage.removeItem(STORAGE_KEY);
}

export function authenticate(name: string): AuthUser | null {
  const trimmed = name.trim();
  const role = ALLOWED_USERS[trimmed];
  if (!role) return null;
  const user: AuthUser = { name: trimmed, role };
  saveUser(user);
  return user;
}
