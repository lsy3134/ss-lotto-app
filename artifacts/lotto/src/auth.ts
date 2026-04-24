// ─────────────────────────────────────────────────────────
// 관리자 계정 (이름 + 비밀번호) — 코드에서만 관리, UI 변경 불가
// ─────────────────────────────────────────────────────────
export const ADMIN_CREDENTIALS: Record<string, string> = {
  "이수예": "1004",
  "유미선": "1004",
};

// ─────────────────────────────────────────────────────────
// 기본 일반 사용자 목록 (관리자가 UI에서 추가/삭제 가능)
// ─────────────────────────────────────────────────────────
const DEFAULT_USERS: string[] = [
  "김승희", "도의지", "진유진", "박정민", "신현진",
];

const USERS_STORAGE_KEY = "lotto_allowed_users";

export type Role = "admin" | "user";

export interface AuthUser {
  name: string;
  role: Role;
}

// ── 일반 사용자 목록 관리 ───────────────────────────────
export function getUserList(): string[] {
  try {
    const raw = localStorage.getItem(USERS_STORAGE_KEY);
    if (!raw) return [...DEFAULT_USERS];
    const parsed = JSON.parse(raw) as string[];
    return Array.isArray(parsed) ? parsed : [...DEFAULT_USERS];
  } catch {
    return [...DEFAULT_USERS];
  }
}

export function saveUserList(list: string[]) {
  localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(list));
}

export function addUser(name: string): { ok: boolean; reason?: string } {
  const trimmed = name.trim();
  if (!trimmed) return { ok: false, reason: "이름을 입력해 주세요." };
  if (ADMIN_CREDENTIALS[trimmed]) return { ok: false, reason: "관리자 이름은 추가할 수 없습니다." };
  const list = getUserList();
  if (list.includes(trimmed)) return { ok: false, reason: "이미 등록된 사용자입니다." };
  saveUserList([...list, trimmed]);
  return { ok: true };
}

export function removeUser(name: string) {
  const list = getUserList().filter(n => n !== name);
  saveUserList(list);
}

// ── 인증 로직 ───────────────────────────────────────────
// step 1: 이름만으로 역할 판별 (비번 검증 전)
export type NameCheckResult = "admin" | "user" | "unknown";

export function checkName(name: string): NameCheckResult {
  const trimmed = name.trim();
  if (ADMIN_CREDENTIALS[trimmed]) return "admin";
  if (getUserList().includes(trimmed)) return "user";
  return "unknown";
}

// step 2: 최종 인증
export function authenticate(name: string, password?: string): AuthUser | null {
  const trimmed = name.trim();
  // 관리자 인증
  if (ADMIN_CREDENTIALS[trimmed]) {
    if (password === ADMIN_CREDENTIALS[trimmed]) {
      const user: AuthUser = { name: trimmed, role: "admin" };
      saveUser(user);
      return user;
    }
    return null;
  }
  // 일반 사용자 인증
  if (getUserList().includes(trimmed)) {
    const user: AuthUser = { name: trimmed, role: "user" };
    saveUser(user);
    return user;
  }
  return null;
}

// ── localStorage 세션 관리 ─────────────────────────────
const SESSION_KEY = "lotto_auth_user";

export function getStoredUser(): AuthUser | null {
  try {
    const raw = localStorage.getItem(SESSION_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as AuthUser;
    // 여전히 유효한 사용자인지 재검증
    if (ADMIN_CREDENTIALS[parsed.name]) {
      return parsed.role === "admin" ? parsed : null;
    }
    if (getUserList().includes(parsed.name)) {
      return parsed.role === "user" ? parsed : null;
    }
    return null;
  } catch {
    return null;
  }
}

export function saveUser(user: AuthUser) {
  localStorage.setItem(SESSION_KEY, JSON.stringify(user));
}

export function clearUser() {
  localStorage.removeItem(SESSION_KEY);
}
