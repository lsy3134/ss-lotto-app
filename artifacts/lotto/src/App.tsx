import { type CSSProperties, createContext, useContext, useEffect, useRef, useState } from "react";
import { Switch, Route, Link, Router as WouterRouter, useLocation } from "wouter";
import * as XLSX from "xlsx";
import SchedulePage from "./pages/Schedule";
import {
  type AuthUser,
  authenticate, clearUser, getStoredUser,
  checkName, getUserList, addUser, removeUser,
} from "./auth";

const BASE_URL = import.meta.env.BASE_URL;
const base = BASE_URL.replace(/\/$/, "");
const APP_VERSION = "v1.0.13";
const SW_VERSION  = "ssapp-v13";

const C = {
  bgPage:       "#eef1f8",
  cardBg:       "#ffffff",
  purple:       "#7c6ef7",
  purpleLight:  "#eeebff",
  blue:         "#5b8dee",
  blueLight:    "#eef3ff",
  green:        "#3db882",
  greenLight:   "#e6f9f1",
  yellow:       "#f0b429",
  yellowLight:  "#fff8e6",
  red:          "#f76e6e",
  redLight:     "#fff0f0",
  orange:       "#f7a55a",
  orangeLight:  "#fff5eb",
  textPrimary:   "#1a2035",
  textSecondary: "#5a6478",
  textMuted:     "#9aa3b5",
  border:        "#e0e5f0",
  borderMid:     "#c8d0e4",
  white:         "#ffffff",
  bgDark:        "#7c6ef7",
};

interface Game { type: string; nums: number[]; }

// ───────────────────────────────────────────
// 인증 컨텍스트 (앱 전역)
// ───────────────────────────────────────────
export const AuthContext = createContext<{
  user: AuthUser | null;
  logout: () => void;
}>({ user: null, logout: () => {} });

export function useAuth() { return useContext(AuthContext); }

// ───────────────────────────────────────────
// 로그인 화면 (2단계: 이름 → 비밀번호)
// ───────────────────────────────────────────
function LoginScreen({ onLogin }: { onLogin: (user: AuthUser) => void }) {
  const [step, setStep] = useState<"name" | "password">("name");
  const [name, setName] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [shake, setShake] = useState(false);

  function triggerShake(msg: string) {
    setError(msg);
    setShake(true);
    setTimeout(() => setShake(false), 450);
  }

  function handleNameNext() {
    const result = checkName(name);
    if (result === "unknown") {
      triggerShake("등록되지 않은 사용자입니다.");
    } else if (result === "admin") {
      setError("");
      setStep("password");
    } else {
      // 일반 사용자 — 바로 로그인
      const user = authenticate(name);
      if (user) onLogin(user);
      else triggerShake("접근 권한이 없습니다.");
    }
  }

  function handlePasswordSubmit() {
    const user = authenticate(name, password);
    if (user) {
      onLogin(user);
    } else {
      triggerShake("비밀번호가 올바르지 않습니다.");
      setPassword("");
    }
  }

  const inputStyle = (hasError: boolean): CSSProperties => ({
    width: "100%", boxSizing: "border-box",
    padding: "14px 16px", borderRadius: 14,
    border: `1.5px solid ${hasError ? C.red : C.border}`,
    fontSize: "1rem", outline: "none",
    color: C.textPrimary, background: "#f8fafc",
    marginBottom: 8, fontFamily: "inherit",
  });

  const btnStyle: CSSProperties = {
    width: "100%", padding: "14px 0", marginTop: 4,
    borderRadius: 14, border: "none",
    background: "linear-gradient(135deg, #7c6ef7 0%, #5b4de8 100%)",
    color: "white", fontWeight: 700, fontSize: "1rem",
    cursor: "pointer",
    boxShadow: "0 4px 16px rgba(124,110,247,0.35)",
    fontFamily: "inherit",
  };

  return (
    <div style={{
      backgroundColor: C.bgPage,
      minHeight: "100dvh",
      display: "flex", flexDirection: "column",
      alignItems: "center", justifyContent: "center",
      padding: "24px",
    }}>
      <img
        src={`${BASE_URL}char_cloud.png`}
        alt=""
        style={{
          width: 100, height: 100, borderRadius: "50%",
          objectFit: "cover", marginBottom: 24,
          boxShadow: `0 0 24px ${C.purple}44`,
          animation: "floatBob 3s ease-in-out infinite",
        }}
      />
      <div style={{
        background: "white", borderRadius: 24,
        padding: "32px 28px", width: "100%", maxWidth: 360,
        boxShadow: "0 8px 32px rgba(100,110,180,0.13)",
        animation: shake ? "shake 0.45s ease" : "none",
      }}>
        <h2 style={{ margin: "0 0 4px", fontSize: "1.3rem", fontWeight: 800, color: C.textPrimary, textAlign: "center" }}>
          SS앱
        </h2>

        {step === "name" ? (
          <>
            <p style={{ margin: "0 0 22px", fontSize: "0.85rem", color: C.textSecondary, textAlign: "center" }}>
              이름을 입력해 주세요
            </p>
            <input
              autoFocus
              value={name}
              onChange={e => { setName(e.target.value); setError(""); }}
              onKeyDown={e => e.key === "Enter" && handleNameNext()}
              placeholder="이름"
              style={inputStyle(!!error)}
            />
            {error && (
              <p style={{ margin: "0 0 10px", fontSize: "0.82rem", color: C.red, textAlign: "center" }}>{error}</p>
            )}
            <button onClick={handleNameNext} style={btnStyle}>다음</button>
          </>
        ) : (
          <>
            <p style={{ margin: "0 0 4px", fontSize: "0.85rem", color: C.textSecondary, textAlign: "center" }}>
              관리자 확인
            </p>
            <p style={{ margin: "0 0 18px", fontSize: "0.95rem", fontWeight: 700, color: C.purple, textAlign: "center" }}>
              {name}
            </p>
            <input
              autoFocus
              type="password"
              value={password}
              onChange={e => { setPassword(e.target.value); setError(""); }}
              onKeyDown={e => e.key === "Enter" && handlePasswordSubmit()}
              placeholder="비밀번호"
              style={inputStyle(!!error)}
            />
            {error && (
              <p style={{ margin: "0 0 10px", fontSize: "0.82rem", color: C.red, textAlign: "center" }}>{error}</p>
            )}
            <button onClick={handlePasswordSubmit} style={btnStyle}>로그인</button>
            <button
              onClick={() => { setStep("name"); setPassword(""); setError(""); }}
              style={{
                width: "100%", marginTop: 10, padding: "10px 0",
                borderRadius: 14, border: `1px solid ${C.border}`,
                background: "white", color: C.textSecondary,
                fontWeight: 600, fontSize: "0.9rem", cursor: "pointer",
                fontFamily: "inherit",
              }}
            >← 돌아가기</button>
          </>
        )}
      </div>

      <style>{`
        @keyframes shake {
          0%, 100% { transform: translateX(0); }
          20%       { transform: translateX(-8px); }
          40%       { transform: translateX(8px); }
          60%       { transform: translateX(-6px); }
          80%       { transform: translateX(6px); }
        }
        @keyframes floatBob {
          0%, 100% { transform: translateY(0); }
          50%       { transform: translateY(-8px); }
        }
      `}</style>
    </div>
  );
}

// ───────────────────────────────────────────
// 사용자 관리 페이지 (admin 전용)
// ───────────────────────────────────────────
function UserManagementPage() {
  const { user } = useAuth();
  const [, setLocation] = useLocation();
  const [users, setUsers] = useState<string[]>(() => getUserList());
  const [newName, setNewName] = useState("");
  const [addError, setAddError] = useState("");
  const [addOk, setAddOk] = useState(false);

  if (user?.role !== "admin") {
    setLocation(`${base}/`);
    return null;
  }

  function handleAdd() {
    const result = addUser(newName);
    if (result.ok) {
      setUsers(getUserList());
      setNewName("");
      setAddError("");
      setAddOk(true);
      setTimeout(() => setAddOk(false), 2000);
    } else {
      setAddError(result.reason ?? "오류가 발생했습니다.");
    }
  }

  function handleRemove(name: string) {
    removeUser(name);
    setUsers(getUserList());
  }

  return (
    <div style={{ backgroundColor: C.bgPage, minHeight: "100dvh", padding: "0 0 100px" }}>
      {/* 헤더 */}
      <div style={{
        display: "flex", alignItems: "center", gap: 10,
        padding: "16px 18px 12px",
        borderBottom: `1px solid ${C.border}`,
        background: "#fff",
      }}>
        <button
          onClick={() => setLocation(`${base}/`)}
          style={{
            padding: "8px 14px", background: "#f0f0f0",
            border: "none", borderRadius: 10, cursor: "pointer",
            fontSize: "0.9rem", color: "#555", fontWeight: 600,
          }}
        >← 홈</button>
        <div style={{ flex: 1, textAlign: "center", fontWeight: 800, fontSize: "1.05rem", color: C.textPrimary }}>
          사용자 관리
        </div>
        <div style={{ width: 60 }} />
      </div>

      <div style={{ padding: "20px 18px", display: "flex", flexDirection: "column", gap: 16 }}>
        {/* 사용자 추가 */}
        <div style={{
          background: "white", borderRadius: 18,
          padding: "20px", border: `1px solid ${C.border}`,
          boxShadow: "0 2px 10px rgba(100,110,180,0.07)",
        }}>
          <div style={{ fontWeight: 700, fontSize: "0.9rem", color: C.textPrimary, marginBottom: 14 }}>
            일반 사용자 추가
          </div>
          <input
            value={newName}
            onChange={e => { setNewName(e.target.value); setAddError(""); setAddOk(false); }}
            onKeyDown={e => e.key === "Enter" && handleAdd()}
            placeholder="이름 입력"
            style={{
              width: "100%", boxSizing: "border-box",
              padding: "12px 14px", borderRadius: 12,
              border: `1.5px solid ${addError ? C.red : C.border}`,
              fontSize: "1rem", outline: "none",
              color: C.textPrimary, background: "#f8fafc",
              marginBottom: 8, fontFamily: "inherit",
            }}
          />
          {addError && (
            <p style={{ margin: "0 0 8px", fontSize: "0.82rem", color: C.red }}>{addError}</p>
          )}
          {addOk && (
            <p style={{ margin: "0 0 8px", fontSize: "0.82rem", color: C.green }}>✅ 추가되었습니다.</p>
          )}
          <button
            onClick={handleAdd}
            style={{
              width: "100%", padding: "12px 0", borderRadius: 12, border: "none",
              background: "linear-gradient(135deg, #7c6ef7 0%, #5b4de8 100%)",
              color: "white", fontWeight: 700, fontSize: "0.95rem",
              cursor: "pointer", fontFamily: "inherit",
            }}
          >+ 추가</button>
        </div>

        {/* 사용자 목록 */}
        <div style={{
          background: "white", borderRadius: 18,
          padding: "20px", border: `1px solid ${C.border}`,
          boxShadow: "0 2px 10px rgba(100,110,180,0.07)",
        }}>
          <div style={{ fontWeight: 700, fontSize: "0.9rem", color: C.textPrimary, marginBottom: 14 }}>
            등록된 일반 사용자 ({users.length}명)
          </div>
          {users.length === 0 ? (
            <p style={{ color: C.textMuted, fontSize: "0.85rem", textAlign: "center", margin: "16px 0" }}>
              등록된 사용자가 없습니다.
            </p>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {users.map(u => (
                <div key={u} style={{
                  display: "flex", alignItems: "center",
                  padding: "10px 14px", borderRadius: 12,
                  background: "#f8fafc", border: `1px solid ${C.border}`,
                }}>
                  <span style={{ flex: 1, fontWeight: 600, fontSize: "0.95rem", color: C.textPrimary }}>{u}</span>
                  <span style={{
                    fontSize: "0.72rem", padding: "2px 8px", borderRadius: 10,
                    background: C.blueLight, color: C.blue,
                    border: `1px solid #c0d4f7`, fontWeight: 600, marginRight: 10,
                  }}>일반</span>
                  <button
                    onClick={() => handleRemove(u)}
                    style={{
                      padding: "4px 10px", borderRadius: 8,
                      border: `1px solid ${C.red}44`, background: C.redLight,
                      color: C.red, fontWeight: 700, fontSize: "0.8rem",
                      cursor: "pointer", fontFamily: "inherit",
                    }}
                  >삭제</button>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* 관리자 계정 안내 */}
        <div style={{
          background: C.purpleLight, borderRadius: 14,
          padding: "14px 16px", border: `1px solid #c5befa`,
        }}>
          <div style={{ fontWeight: 700, fontSize: "0.82rem", color: C.purple, marginBottom: 8 }}>
            👑 관리자 계정 (고정, 변경 불가)
          </div>
          {["이수예", "유미선"].map(n => (
            <div key={n} style={{
              fontSize: "0.88rem", color: C.textPrimary, fontWeight: 600,
              padding: "4px 0", borderBottom: `1px solid ${C.border}33`,
            }}>{n}</div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ───────────────────────────────────────────
// 인트로 화면
// ───────────────────────────────────────────
function IntroView({ onEnter }: { onEnter: () => void }) {
  return (
    <div style={{
      backgroundColor: C.bgPage,
      minHeight: "100dvh",
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      justifyContent: "center",
      color: C.textPrimary,
      textAlign: "center",
      padding: "20px",
    }}>
      <img
        src={`${BASE_URL}char_cloud.png`}
        alt="Welcome"
        style={{
          width: "160px",
          height: "160px",
          objectFit: "cover",
          borderRadius: "50%",
          marginBottom: "20px",
          boxShadow: `0 0 24px ${C.purple}`,
          animation: "floatBob 3s ease-in-out infinite",
        }}
      />

      {/* 말풍선 */}
      <div style={{ position: "relative", marginBottom: "36px" }}>
        <div style={{
          backgroundColor: "white",
          color: "#333",
          padding: "14px 22px",
          borderRadius: "18px",
          fontWeight: 700,
          fontSize: "1rem",
          lineHeight: 1.6,
          boxShadow: "0 4px 16px rgba(0,0,0,0.15)",
        }}>
          안녕! 오늘도 수고 많으세요!<br />
          제가 근무표 작성을 도와드릴까요?
        </div>
        {/* 말풍선 꼬리 */}
        <div style={{
          position: "absolute",
          bottom: "-10px",
          left: "50%",
          transform: "translateX(-50%)",
          width: 0,
          height: 0,
          borderLeft: "10px solid transparent",
          borderRight: "10px solid transparent",
          borderTop: "10px solid white",
        }} />
      </div>

      <button
        onClick={onEnter}
        style={{
          background: "linear-gradient(135deg, #7c6ef7 0%, #5b4de8 100%)",
          color: "white",
          border: "none",
          padding: "15px 52px",
          borderRadius: "30px",
          fontSize: "1.1rem",
          fontWeight: 700,
          cursor: "pointer",
          boxShadow: "0 4px 20px rgba(124,110,247,0.45)",
          transition: "transform 0.15s, box-shadow 0.15s",
        }}
        onMouseEnter={e => (e.currentTarget.style.transform = "scale(1.05)")}
        onMouseLeave={e => (e.currentTarget.style.transform = "scale(1)")}
      >
        입장하기
      </button>
    </div>
  );
}

// ───────────────────────────────────────────
// 홈 / 서비스 선택 화면
// ───────────────────────────────────────────
function HomePage() {
  const { user, logout } = useAuth();
  const [showIntro, setShowIntro] = useState(() => {
    return !sessionStorage.getItem("ss_entered");
  });

  function enter() {
    sessionStorage.setItem("ss_entered", "1");
    setShowIntro(false);
  }

  if (showIntro) return <IntroView onEnter={enter} />;

  return (
    <div style={{
      backgroundColor: C.bgPage,
      minHeight: "100dvh",
      padding: "48px 24px 100px",
      color: C.textPrimary,
    }}>
      <header style={{ marginBottom: "36px" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6 }}>
          <h2 style={{ fontSize: "1.5rem", margin: 0, fontWeight: 800, color: C.textPrimary }}>서비스 선택</h2>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{
              fontSize: "0.75rem", color: C.textSecondary,
              background: user?.role === "admin" ? C.purpleLight : C.blueLight,
              border: `1px solid ${user?.role === "admin" ? "#c5befa" : "#c0d4f7"}`,
              borderRadius: 20, padding: "3px 10px", fontWeight: 600,
            }}>
              {user?.name} {user?.role === "admin" ? "👑" : ""}
            </span>
            <button
              onClick={logout}
              style={{
                fontSize: "0.72rem", padding: "4px 10px",
                border: `1px solid ${C.border}`, borderRadius: 20,
                background: "white", color: C.textSecondary,
                cursor: "pointer", fontWeight: 600,
              }}
            >로그아웃</button>
          </div>
        </div>
        <p style={{ margin: 0, color: C.textSecondary, fontSize: "0.9rem" }}>필요한 업무를 골라주세요.</p>
      </header>

      {/* 근무표 카드 */}
      <Link href={`${base}/schedule`} style={{ textDecoration: "none" }}>
        <div style={{
          backgroundColor: C.cardBg,
          borderRadius: "20px",
          padding: "22px 20px",
          marginBottom: "18px",
          display: "flex",
          alignItems: "center",
          cursor: "pointer",
          border: `1px solid ${C.border}`,
          transition: "transform 0.15s, box-shadow 0.15s",
          boxShadow: `0 2px 12px rgba(100,110,180,0.08)`,
        }}
          onMouseEnter={e => {
            (e.currentTarget as HTMLDivElement).style.transform = "translateY(-2px)";
            (e.currentTarget as HTMLDivElement).style.boxShadow = `0 8px 24px rgba(91,141,238,0.2)`;
          }}
          onMouseLeave={e => {
            (e.currentTarget as HTMLDivElement).style.transform = "translateY(0)";
            (e.currentTarget as HTMLDivElement).style.boxShadow = `0 2px 12px rgba(100,110,180,0.08)`;
          }}
        >
          <img
            src={`${BASE_URL}char_dino.png`}
            alt="근무표"
            style={{ width: "60px", height: "60px", objectFit: "cover", borderRadius: "14px", marginRight: "18px", flexShrink: 0 }}
          />
          <div style={{ flex: 1, textAlign: "left" }}>
            <div style={{ fontWeight: 800, fontSize: "1.1rem", marginBottom: "4px", color: C.textPrimary }}>근무표 생성</div>
            <div style={{ color: C.textSecondary, fontSize: "0.83rem" }}>63명 캐디 배정 및 일정 관리</div>
          </div>
          <div style={{ fontSize: "1.5rem", color: C.blue, opacity: 0.6 }}>›</div>
        </div>
      </Link>

      {/* 로또 카드 — admin 전용 */}
      {user?.role === "admin" && (
        <Link href={`${base}/lotto`} style={{ textDecoration: "none" }}>
          <div style={{
            backgroundColor: C.yellowLight,
            borderRadius: "20px",
            padding: "22px 20px",
            display: "flex",
            alignItems: "center",
            cursor: "pointer",
            border: `1px solid #f0d080`,
            transition: "transform 0.15s, box-shadow 0.15s",
            boxShadow: `0 2px 12px rgba(240,180,41,0.10)`,
          }}
            onMouseEnter={e => {
              (e.currentTarget as HTMLDivElement).style.transform = "translateY(-2px)";
              (e.currentTarget as HTMLDivElement).style.boxShadow = `0 8px 24px rgba(240,180,41,0.22)`;
            }}
            onMouseLeave={e => {
              (e.currentTarget as HTMLDivElement).style.transform = "translateY(0)";
              (e.currentTarget as HTMLDivElement).style.boxShadow = `0 2px 12px rgba(240,180,41,0.10)`;
            }}
          >
            <img
              src={`${BASE_URL}char_dragon.png`}
              alt="로또"
              style={{ width: "60px", height: "60px", objectFit: "cover", borderRadius: "14px", marginRight: "18px", flexShrink: 0 }}
            />
            <div style={{ flex: 1, textAlign: "left" }}>
              <div style={{ fontWeight: 800, fontSize: "1.1rem", marginBottom: "4px", color: "#a07000" }}>나만의 로또 생성</div>
              <div style={{ color: C.textSecondary, fontSize: "0.83rem" }}>오늘의 행운 번호 추출</div>
            </div>
            <div style={{ fontSize: "1.5rem", color: C.yellow, opacity: 0.8 }}>›</div>
          </div>
        </Link>
      )}

      {/* 사용자 관리 카드 — admin 전용 */}
      {user?.role === "admin" && (
        <Link href={`${base}/users`} style={{ textDecoration: "none", marginTop: 10, display: "block" }}>
          <div style={{
            backgroundColor: C.purpleLight,
            borderRadius: "20px",
            padding: "16px 20px",
            display: "flex",
            alignItems: "center",
            cursor: "pointer",
            border: `1px solid #c5befa`,
            transition: "transform 0.15s, box-shadow 0.15s",
            boxShadow: `0 2px 10px rgba(124,110,247,0.08)`,
          }}
            onMouseEnter={e => {
              (e.currentTarget as HTMLDivElement).style.transform = "translateY(-2px)";
              (e.currentTarget as HTMLDivElement).style.boxShadow = `0 8px 20px rgba(124,110,247,0.18)`;
            }}
            onMouseLeave={e => {
              (e.currentTarget as HTMLDivElement).style.transform = "translateY(0)";
              (e.currentTarget as HTMLDivElement).style.boxShadow = `0 2px 10px rgba(124,110,247,0.08)`;
            }}
          >
            <span style={{ fontSize: "2rem", marginRight: "16px" }}>👥</span>
            <div style={{ flex: 1, textAlign: "left" }}>
              <div style={{ fontWeight: 800, fontSize: "0.95rem", marginBottom: "2px", color: C.purple }}>사용자 관리</div>
              <div style={{ color: C.textSecondary, fontSize: "0.78rem" }}>일반 사용자 추가 / 삭제</div>
            </div>
            <div style={{ fontSize: "1.5rem", color: C.purple, opacity: 0.5 }}>›</div>
          </div>
        </Link>
      )}

      {/* 하단 + 디버그 패널 */}
      <DebugPanel />
    </div>
  );
}

function DebugPanel() {
  const [tapCount, setTapCount] = useState(0);
  const [show, setShow] = useState(false);
  const [info, setInfo] = useState<Record<string, string>>({});
  const [serverStatus, setServerStatus] = useState<"idle"|"loading"|"done"|"error">("idle");

  function handleTap() {
    const next = tapCount + 1;
    setTapCount(next);
    if (next >= 3) { setShow(true); loadDebugInfo(); }
  }

  function loadDebugInfo() {
    const localHM = (() => {
      try { return JSON.parse(localStorage.getItem("lotto_holidayMap") ?? "{}"); } catch { return {}; }
    })();
    const localKeys = Object.keys(localHM).map(k => k.slice(0, 2)).filter((v, i, a) => a.indexOf(v) === i).sort();
    setInfo({
      appVersion: APP_VERSION,
      swVersion: SW_VERSION,
      url: window.location.href,
      localUpdatedAt: localStorage.getItem("lotto_holidayMapUpdatedAt") ?? "(없음)",
      localHolidayMonths: localKeys.length ? localKeys.join(", ") : "(없음)",
      localFileName: localStorage.getItem("lotto_holidayFileName") ?? "(없음)",
    });
    setServerStatus("loading");
    fetch(`/api/holiday-map?_=${Date.now()}`, { cache: "no-store" })
      .then(r => r.json())
      .then((data: { fileName: string; holidayMap: Record<string, string[]>; updatedAt: string | null }) => {
        const serverKeys = Object.keys(data.holidayMap ?? {}).map(k => k.slice(0, 2)).filter((v, i, a) => a.indexOf(v) === i).sort();
        setInfo(prev => ({
          ...prev,
          serverFileName: data.fileName || "(없음)",
          serverUpdatedAt: data.updatedAt ?? "(없음)",
          serverHolidayMonths: serverKeys.length ? serverKeys.join(", ") : "(없음)",
          serverKeyCount: String(Object.keys(data.holidayMap ?? {}).length),
        }));
        setServerStatus("done");
      })
      .catch(e => {
        setInfo(prev => ({ ...prev, serverError: String(e) }));
        setServerStatus("error");
      });
  }

  const rows: [string, string][] = [
    ["앱 버전",           info.appVersion ?? ""],
    ["SW 버전",           info.swVersion ?? ""],
    ["접속 URL",          info.url ?? ""],
    ["─ 로컬 ─",          ""],
    ["로컬 파일명",        info.localFileName ?? ""],
    ["로컬 updatedAt",    info.localUpdatedAt ?? ""],
    ["로컬 월 목록",       info.localHolidayMonths ?? ""],
    ["─ 서버 ─",          serverStatus === "loading" ? "로딩 중…" : ""],
    ["서버 파일명",        info.serverFileName ?? ""],
    ["서버 updatedAt",    info.serverUpdatedAt ?? ""],
    ["서버 월 목록",       info.serverHolidayMonths ?? ""],
    ["서버 날짜 수",       info.serverKeyCount ?? ""],
    ["서버 오류",          info.serverError ?? ""],
  ];

  return (
    <>
      <div style={{ marginTop: "48px", textAlign: "center" }}>
        <img
          src={`${BASE_URL}char_smile.png`}
          alt=""
          style={{ width: "36px", height: "36px", objectFit: "cover", borderRadius: "50%", display: "block", margin: "0 auto 8px", opacity: 0.35 }}
        />
        <p
          onClick={handleTap}
          style={{ fontSize: "0.72rem", margin: 0, opacity: 0.35, cursor: "default", userSelect: "none" }}
        >
          SS앱 — {APP_VERSION} {tapCount > 0 && tapCount < 3 ? `(${3 - tapCount}번 더)` : ""}
        </p>
      </div>

      {show && (
        <div style={{
          marginTop: 20, borderRadius: 14, border: "1.5px solid #1565c0",
          background: "#f0f4ff", padding: "14px 16px", fontSize: "0.72rem",
          fontFamily: "monospace", color: "#1a237e",
        }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
            <strong style={{ fontSize: "0.8rem" }}>🔍 디버그 정보</strong>
            <div style={{ display: "flex", gap: 8 }}>
              <button onClick={loadDebugInfo} style={{ fontSize: "0.7rem", padding: "3px 10px", border: "1px solid #1565c0", borderRadius: 8, background: "#fff", color: "#1565c0", cursor: "pointer", fontWeight: 700 }}>새로고침</button>
              <button onClick={() => { setShow(false); setTapCount(0); }} style={{ fontSize: "0.7rem", padding: "3px 10px", border: "1px solid #ccc", borderRadius: 8, background: "#fff", color: "#666", cursor: "pointer" }}>닫기</button>
            </div>
          </div>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <tbody>
              {rows.filter(([, v]) => v !== undefined).map(([label, value]) =>
                label.startsWith("─") ? (
                  <tr key={label}><td colSpan={2} style={{ paddingTop: 8, paddingBottom: 2, fontWeight: 800, color: "#1565c0", fontSize: "0.7rem" }}>{label}</td></tr>
                ) : (
                  <tr key={label}>
                    <td style={{ paddingRight: 8, paddingBottom: 3, color: "#555", whiteSpace: "nowrap", verticalAlign: "top" }}>{label}</td>
                    <td style={{ wordBreak: "break-all", paddingBottom: 3, fontWeight: value ? 700 : 400, color: value ? "#1a237e" : "#aaa" }}>{value || "(없음)"}</td>
                  </tr>
                )
              )}
            </tbody>
          </table>
        </div>
      )}
    </>
  );
}


// ───────────────────────────────────────────
// 로또 페이지
// ───────────────────────────────────────────
function LottoPage() {
  const [, setLocation] = useLocation();
  const { user } = useAuth();
  const isAdmin = user?.role === "admin";
  const allDraws = useRef<number[][]>([]);
  const pastWinners = useRef<Set<string>>(new Set());
  const [hotText, setHotText] = useState("");
  const [games, setGames] = useState<Game[]>([]);
  const [log, setLog] = useState("");
  const [input, setInput] = useState("");

  useEffect(() => {
    if (!isAdmin) { setLocation(`${base}/`); return; }
    async function init() { await loadExcel(); loadLatest(); }
    init();
  }, [isAdmin]);

  async function loadExcel() {
    try {
      const res = await fetch(`${BASE_URL}lotto.xlsx`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const buf = await res.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(wb.Sheets[wb.SheetNames[0]]);
      rows.forEach((row) => {
        const nums = [row["당첨번호"], row["Unnamed: 3"], row["Unnamed: 4"], row["Unnamed: 5"], row["Unnamed: 6"], row["Unnamed: 7"]]
          .map(Number).filter((v) => Number.isInteger(v) && v >= 1 && v <= 45).sort((a, b) => a - b);
        if (nums.length === 6) { pastWinners.current.add(nums.join(",")); allDraws.current.push(nums); }
      });
    } catch (e) {
      console.warn("lotto.xlsx 로드 실패 — 기본 모드로 실행:", e);
    }
  }

  function loadLatest() {
    const saved = localStorage.getItem("latestLotto");
    if (saved) {
      const arr = saved.split(",").map(Number);
      if (!pastWinners.current.has(saved)) { pastWinners.current.add(saved); allDraws.current.unshift(arr); }
    }
  }

  function addLatest() {
    const nums = input.split(/[\s,]+/).map((n) => Number(n.trim())).filter((n) => n >= 1 && n <= 45).sort((a, b) => a - b);
    if (nums.length < 6 || nums.length > 7) { alert("6개 또는 7개(보너스 포함) 숫자를 입력해주세요."); return; }
    nums.splice(6);
    const key = nums.join(",");
    if (pastWinners.current.has(key)) { alert("이미 존재"); return; }
    pastWinners.current.add(key); allDraws.current.unshift(nums);
    localStorage.setItem("latestLotto", key); setInput(""); alert("추가 완료");
  }

  function rand() { return Math.floor(Math.random() * 45) + 1; }

  // 빈 배열 안전 처리: arr가 비어있으면 랜덤 유니크 숫자로 대체
  function pick(arr: number[], n: number): number[] {
    const src = arr.length >= n ? arr : Array.from({ length: 45 }, (_, i) => i + 1);
    const copy = [...src]; const res: number[] = [];
    while (res.length < n) {
      const i = Math.floor(Math.random() * copy.length);
      res.push(copy.splice(i, 1)[0]);
    }
    return res;
  }

  // 기존 nums와 중복 없이 source에서 count개 선택
  function safePick(base: number[], source: number[], count: number): number[] {
    const pool = source.length >= count
      ? source
      : Array.from({ length: 45 }, (_, i) => i + 1);
    const result: number[] = [];
    let guard = 0;
    while (result.length < count && guard++ < 10000) {
      const n = pool[Math.floor(Math.random() * pool.length)];
      if (!base.includes(n) && !result.includes(n)) result.push(n);
    }
    return result;
  }

  // ── 공통 유효성 검사 ─────────────────────────────────────────
  // noDuplast: 끝자리 중복 3개 이상 금지 (균형형만 적용)
  function isValid(nums: number[], opts: { checkOddEven?: boolean; checkConsec?: boolean; checkLastDigit?: boolean } = {}): boolean {
    const { checkOddEven = true, checkConsec = true, checkLastDigit = true } = opts;
    if (nums.length !== 6) return false;
    if (nums.some((n) => !Number.isInteger(n) || n < 1 || n > 45)) return false;
    if (new Set(nums).size !== 6) return false;                        // 중복 숫자 금지
    if (pastWinners.current.has(nums.join(","))) return false;         // 과거 당첨 조합 제외
    if (checkLastDigit) {
      const last: Record<number, number> = {};
      nums.forEach((n) => { const d = n % 10; last[d] = (last[d] || 0) + 1; });
      if (Object.values(last).some((v) => v >= 3)) return false;       // 끝자리 3개 이상 중복 금지
    }
    if (checkConsec) {
      let cnt = 1;
      for (let i = 1; i < nums.length; i++) {
        if (nums[i] === nums[i - 1] + 1) { cnt++; if (cnt > 3) return false; } else cnt = 1;
      }
    }
    if (checkOddEven) {
      const odd = nums.filter((n) => n % 2).length;
      if (odd < 2 || odd > 4) return false;                            // 홀짝 2~4 유지
    }
    return true;
  }

  // ── Hot / Cold 계산 ──────────────────────────────────────────
  // 최근 10회 기준, 1~45 전체 번호를 등장 횟수로 정렬
  // → 상위 6개 = Hot(자주 나온 번호), 하위 6개 = Cold(거의 안 나온 번호)
  // → 1~45 전체를 대상으로 하므로 hot ∩ cold = ∅ 보장
  function getHotCold() {
    if (allDraws.current.length < 10) return { hot: [] as number[], cold: [] as number[] };

    const freq: Record<number, number> = {};
    // 1~45 모든 번호를 0으로 초기화 → Cold 후보에 미등장 번호도 포함
    for (let i = 1; i <= 45; i++) freq[i] = 0;
    allDraws.current.slice(0, 10).forEach((draw) => {
      draw.forEach((n) => { freq[n] = (freq[n] || 0) + 1; });
    });

    const entries = Object.entries(freq)
      .map(([num, count]) => ({ num: Number(num), count: Number(count) }))
      .sort((a, b) => b.count - a.count);   // 빈도 내림차순

    const hot  = entries.slice(0, 6).map((x) => x.num);   // 상위 6개 (자주 등장)
    const cold = entries.slice(-6).map((x) => x.num);      // 하위 6개 (거의 미등장)

    return { hot, cold };
  }

  // ── 균형형 조합 ───────────────────────────────────────────────
  // Hot 2개 + Cold 2개 + 랜덤 2개 / 중복·연속3·홀짝2~4 조건 통과
  function generateBalanced(hot: number[], cold: number[]): number[] {
    const useHotCold = hot.length >= 2 && cold.length >= 2;
    for (let attempt = 0; attempt < 5000; attempt++) {
      const nums: number[] = [];
      if (useHotCold) {
        nums.push(...safePick(nums, hot, 2));   // Hot에서 2개
        nums.push(...safePick(nums, cold, 2));  // Cold에서 2개 (Hot과 중복 불가)
      }
      // 나머지 랜덤 채움
      let inner = 0;
      while (nums.length < 6 && inner++ < 500) {
        const n = rand();
        if (!nums.includes(n)) nums.push(n);
      }
      nums.sort((a, b) => a - b);
      if (isValid(nums, { checkOddEven: true, checkConsec: true, checkLastDigit: true })) return nums;
    }
    // 최종 fallback: 홀짝 조건만 유지
    for (let i = 0; i < 10000; i++) {
      const nums = pick([], 6).sort((a, b) => a - b);
      const odd = nums.filter((n) => n % 2).length;
      if (odd >= 2 && odd <= 4 && new Set(nums).size === 6) return nums;
    }
    return pick([], 6).sort((a, b) => a - b);
  }

  // ── 변형 전략 조합 ───────────────────────────────────────────────
  // Hot 3개 + Cold 1개 + 랜덤 2개 / 중복·연속3·홀짝2~4 조건 통과
  function generateVariant(hot: number[], cold: number[]): number[] {
    const useHotCold = hot.length >= 3 && cold.length >= 1;
    for (let attempt = 0; attempt < 5000; attempt++) {
      const nums: number[] = [];
      if (useHotCold) {
        nums.push(...safePick(nums, hot, 3));   // Hot에서 3개
        nums.push(...safePick(nums, cold, 1));  // Cold에서 1개 (Hot과 중복 불가)
      }
      // 나머지 랜덤 채움
      let inner = 0;
      while (nums.length < 6 && inner++ < 500) {
        const n = rand();
        if (!nums.includes(n)) nums.push(n);
      }
      nums.sort((a, b) => a - b);
      if (isValid(nums, { checkOddEven: true, checkConsec: true, checkLastDigit: true })) return nums;
    }
    // 최종 fallback: 홀짝 조건만 유지
    for (let i = 0; i < 10000; i++) {
      const nums = pick([], 6).sort((a, b) => a - b);
      const odd = nums.filter((n) => n % 2).length;
      if (odd >= 2 && odd <= 4 && new Set(nums).size === 6) return nums;
    }
    return pick([], 6).sort((a, b) => a - b);
  }

  function generate() {
    const { hot, cold } = getHotCold();
    const hasData = allDraws.current.length >= 10;
    setHotText(hasData
      ? `🔥 핫(상위6): ${hot.join(", ")}  |  ❄️ 콜드(하위6): ${cold.join(", ")}`
      : "📊 데이터 10회 미만 — 랜덤 모드로 생성"
    );
    const result: Game[] = [];
    for (let i = 0; i < 3; i++) result.push({ type: "균형형", nums: generateBalanced(hot, cold) });
    for (let i = 0; i < 2; i++) result.push({ type: "변형", nums: generateVariant(hot, cold) });
    setGames(result); setLog("생성 완료");
  }

  const ballColor = (n: number) => {
    if (n <= 10) return "#f9a825";
    if (n <= 20) return "#1e88e5";
    if (n <= 30) return "#e53935";
    if (n <= 40) return "#6d4c41";
    return "#43a047";
  };

  return (
    <div style={{
      fontFamily: "'Noto Sans KR', 'Pretendard', 'Apple SD Gothic Neo', sans-serif",
      background: "#eef1f8",
      minHeight: "100dvh",
      display: "flex",
      flexDirection: "column",
      padding: "0",
    }}>
      {/* 헤더 */}
      <div style={{
        display: "flex", alignItems: "center", gap: 10,
        padding: "16px 18px 12px",
        borderBottom: "1px solid #e5e7eb",
        background: "#fff",
      }}>
        <button
          onClick={() => setLocation(`${base}/`)}
          style={{
            padding: "8px 14px", background: "#f0f0f0",
            border: "none", borderRadius: 10, cursor: "pointer",
            fontSize: "0.9rem", color: "#555", fontWeight: 600,
          }}
        >← 홈</button>
        <div style={{ flex: 1, textAlign: "center" }}>
          <div style={{ fontSize: "1.1rem", fontWeight: 800, color: "#1a1a2e" }}>써니의 로또 추천🚀 </div>
        </div>
        <div style={{ width: 60 }} />
      </div>

      {/* 본문 */}
      <div style={{ flex: 1, padding: "20px 16px 100px", display: "flex", flexDirection: "column", gap: 14 }}>

        {/* 번호 입력 카드 */}
        <div style={{
          background: "#fff", borderRadius: 18,
          padding: "16px", border: "1px solid #e5e7eb",
          boxShadow: "0 2px 8px rgba(0,0,0,0.06)",
        }}>
          <div style={{ fontSize: "0.75rem", color: "#9ca3af", marginBottom: 8, fontWeight: 600 }}>
            이번주 당첨번호 등록
          </div>
          <input
            value={input}
            onChange={(e) => setInput(e.target.value)}
            placeholder="예: 3 11 15 29 35 44"
            style={{
              width: "100%", boxSizing: "border-box",
              padding: "13px 14px", borderRadius: 12,
              border: "1px solid #d1d5db",
              background: "#f9fafb",
              color: "#1a1a2e", fontSize: "1rem",
              outline: "none",
            }}
          />
          <button
            onClick={addLatest}
            style={{
              width: "100%", marginTop: 10,
              padding: "14px 0", borderRadius: 12, border: "none",
              background: "linear-gradient(135deg, #7c6ef7 0%, #5b4de8 100%)",
              color: "#fff", fontWeight: 700, fontSize: "1rem",
              cursor: "pointer",
              boxShadow: "0 4px 14px rgba(124,110,247,0.30)",
            }}
          >번호 추가</button>
        </div>

        {/* 핫/콜드 정보 */}
        {hotText && (
          <div style={{
            background: "#fffbeb", borderRadius: 12,
            padding: "10px 14px", border: "1px solid #fde68a",
            fontSize: "0.8rem", color: "#92400e", lineHeight: 1.6,
          }}>
            {hotText}
          </div>
        )}

        {/* 생성 버튼 */}
        <button
          onClick={generate}
          style={{
            width: "100%", padding: "18px 0", borderRadius: 16, border: "none",
            background: "linear-gradient(135deg, #f0b429 0%, #d08000 100%)",
            color: "#fff", fontWeight: 800, fontSize: "1.15rem",
            cursor: "pointer",
            letterSpacing: 1,
            boxShadow: "0 6px 20px rgba(240,180,41,0.38)",
          }}
        > 5게임 생성</button>

        {/* 결과 */}
        {games.length > 0 && (
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            {games.map((game, i) => (
              <div key={i} style={{
                background: "#fff", borderRadius: 16,
                padding: "14px 16px", border: "1px solid #e5e7eb",
                boxShadow: "0 2px 8px rgba(0,0,0,0.06)",
              }}>
                <div style={{
                  fontSize: "0.72rem", color: "#9ca3af",
                  fontWeight: 700, marginBottom: 10, letterSpacing: 0.5,
                }}>
                  게임 {i + 1} · {game.type}
                </div>
                <div style={{ overflowX: "auto" }}>
                  <div style={{ display: "flex", justifyContent: "center", gap: 8, flexWrap: "nowrap" }}>
                    {game.nums.map((n) => (
                      <div key={n} style={{
                        width: 40, height: 40, minWidth: 40, borderRadius: "50%",
                        display: "flex", alignItems: "center", justifyContent: "center",
                        color: "#fff", fontWeight: 800, fontSize: "1rem",
                        background: ballColor(n),
                        boxShadow: `0 2px 6px ${ballColor(n)}66`,
                        flexShrink: 0,
                      }}>{n}</div>
                    ))}
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}

        {log && (
          <div style={{ textAlign: "center", fontSize: "0.8rem", color: "#9ca3af", marginTop: 4 }}>
            {log}
          </div>
        )}
      </div>
    </div>
  );
}

// ───────────────────────────────────────────
// 하단 탭바
// ───────────────────────────────────────────
function BottomTabBar() {
  const [location, navigate] = useLocation();
  const { user } = useAuth();
  const isAdmin = user?.role === "admin";

  const allTabs = [
    { path: `${base}/`,         icon: "🏠", label: "홈",   adminOnly: false },
    { path: `${base}/schedule`, icon: "📋", label: "근무표", adminOnly: false },
    { path: `${base}/lotto`,    icon: "🎰", label: "로또",  adminOnly: true  },
  ];
  const tabs = allTabs.filter(t => !t.adminOnly || isAdmin);

  return (
    <nav style={{
      position: "fixed",
      bottom: 0, left: 0, right: 0,
      background: "#ffffff",
      borderTop: "1.5px solid #e8ecf4",
      display: "flex",
      zIndex: 100,
      boxShadow: "0 -2px 16px rgba(100,110,180,0.10)",
      paddingBottom: "env(safe-area-inset-bottom)",
    }}>
      {tabs.map((tab) => {
        const isActive = location === tab.path || (tab.path === `${base}/` && location === `${base}`);
        return (
          <button
            key={tab.path}
            onClick={() => navigate(tab.path)}
            style={{
              flex: 1,
              display: "flex", flexDirection: "column",
              alignItems: "center", justifyContent: "center",
              gap: "3px", padding: "10px 0 8px",
              border: "none", background: "transparent",
              cursor: "pointer", transition: "all 0.15s",
            }}
          >
            <span style={{
              fontSize: "22px", lineHeight: 1,
              filter: isActive ? "none" : "grayscale(1) opacity(0.35)",
              transition: "all 0.2s",
            }}>
              {tab.icon}
            </span>
            <span style={{
              fontSize: "10px", fontWeight: 700,
              color: isActive ? "#7c6ef7" : "#b0b8ca",
              transition: "all 0.2s",
              fontFamily: "'Noto Sans KR', sans-serif",
            }}>
              {tab.label}
            </span>
            {isActive && (
              <div style={{
                width: "4px", height: "4px",
                borderRadius: "50%",
                background: "#7c6ef7",
                marginTop: "1px",
              }} />
            )}
          </button>
        );
      })}
    </nav>
  );
}

// ───────────────────────────────────────────
// 라우터 (인증 게이트 포함)
// ───────────────────────────────────────────
export default function App() {
  const [user, setUser] = useState<AuthUser | null>(() => getStoredUser());

  function handleLogin(u: AuthUser) {
    setUser(u);
  }

  function handleLogout() {
    clearUser();
    setUser(null);
  }

  // 인증 안 된 경우 로그인 화면
  if (!user) {
    return <LoginScreen onLogin={handleLogin} />;
  }

  return (
    <AuthContext.Provider value={{ user, logout: handleLogout }}>
      <WouterRouter base={base}>
        <Switch>
          <Route path="/" component={HomePage} />
          <Route path="/schedule" component={SchedulePage} />
          <Route path="/lotto" component={LottoPage} />
          <Route path="/users" component={UserManagementPage} />
          <Route component={HomePage} />
        </Switch>
        <BottomTabBar />
      </WouterRouter>
    </AuthContext.Provider>
  );
}
