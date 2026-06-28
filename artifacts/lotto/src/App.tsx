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

interface ScoredGame {
  nums: number[];
  score: number;
  reasons: string[];
  hotCount: number;
  coldCount: number;
  oddCount: number;
  ranges: Record<string, number>;
  pairScore: number;
}

interface LottoStats {
  freq: Record<number, number>;
  hot10: number[];
  cold10: number[];
  hot30: number[];
  cold30: number[];
  pairFreq: Map<string, number>;
  tripleFreq: Map<string, number>;
  quadFreq: Map<string, number>;
  pairTop20: [string, number][];
  tripleTop20: [string, number][];
  quadTop20: [string, number][];
  sumMean: number;
  sumStd: number;
}

function lc(arr: number[], r: number): number[][] {
  const result: number[][] = [];
  function h(start: number, cur: number[]) {
    if (cur.length === r) { result.push([...cur]); return; }
    for (let i = start; i < arr.length; i++) { cur.push(arr[i]); h(i + 1, cur); cur.pop(); }
  }
  h(0, []);
  return result;
}

function computeLottoStats(draws: number[][]): LottoStats {
  const freq: Record<number, number> = {};
  for (let i = 1; i <= 45; i++) freq[i] = 0;
  const pairFreq = new Map<string, number>();
  const tripleFreq = new Map<string, number>();
  const quadFreq = new Map<string, number>();
  const sums: number[] = [];
  const addMap = (m: Map<string, number>, k: string) => m.set(k, (m.get(k) || 0) + 1);
  draws.forEach(d => {
    d.forEach(n => { freq[n] = (freq[n] || 0) + 1; });
    sums.push(d.reduce((a, b) => a + b, 0));
    lc(d, 2).forEach(c => addMap(pairFreq, c.join(",")));
    lc(d, 3).forEach(c => addMap(tripleFreq, c.join(",")));
    lc(d, 4).forEach(c => addMap(quadFreq, c.join(",")));
  });
  function topHC(n: number) {
    const f: Record<number, number> = {};
    for (let i = 1; i <= 45; i++) f[i] = 0;
    draws.slice(0, n).forEach(d => d.forEach(x => { f[x] = (f[x] || 0) + 1; }));
    const s = Object.entries(f).map(([k, v]) => ({ n: Number(k), c: v })).sort((a, b) => b.c - a.c);
    return { hot: s.slice(0, 6).map(x => x.n), cold: s.slice(-6).map(x => x.n) };
  }
  const { hot: hot10, cold: cold10 } = topHC(10);
  const { hot: hot30, cold: cold30 } = topHC(30);
  const pairTop20 = [...pairFreq.entries()].sort((a, b) => b[1] - a[1]).slice(0, 20);
  const tripleTop20 = [...tripleFreq.entries()].sort((a, b) => b[1] - a[1]).slice(0, 20);
  const quadTop20 = [...quadFreq.entries()].sort((a, b) => b[1] - a[1]).slice(0, 20);
  const sumMean = sums.reduce((a, b) => a + b, 0) / (sums.length || 1);
  const sumStd = Math.sqrt(sums.reduce((a, b) => a + (b - sumMean) ** 2, 0) / (sums.length || 1));
  return { freq, hot10, cold10, hot30, cold30, pairFreq, tripleFreq, quadFreq, pairTop20, tripleTop20, quadTop20, sumMean, sumStd };
}

function scoreLottoCandidate(
  nums: number[], stats: LottoStats,
  maxPairRaw: number, maxTripleRaw: number, maxQuadRaw: number,
): ScoredGame {
  let score = 0;
  const reasons: string[] = [];

  const hotCount = nums.filter(n => stats.hot30.includes(n)).length;
  const coldCount = nums.filter(n => stats.cold30.includes(n)).length;
  let hcPts = 0;
  if (hotCount >= 2) hcPts += 10; else if (hotCount >= 1) hcPts += 5;
  if (coldCount >= 2) hcPts += 10; else if (coldCount >= 1) hcPts += 5;
  score += Math.min(hcPts, 20);
  if (hotCount > 0 || coldCount > 0) reasons.push(`최근 30회 핫번호 ${hotCount}개 포함`);

  const pairRaw = lc(nums, 2).reduce((s, c) => s + (stats.pairFreq.get(c.join(",")) || 0), 0);
  const pairScore = maxPairRaw > 0 ? Math.round((pairRaw / maxPairRaw) * 20) : 10;
  score += pairScore;
  if (pairScore >= 15) reasons.push("동반출현(쌍) 점수 상위권");
  else if (pairScore >= 10) reasons.push("동반출현(쌍) 점수 양호");

  const tripleRaw = lc(nums, 3).reduce((s, c) => s + (stats.tripleFreq.get(c.join(",")) || 0), 0);
  const triplePts = maxTripleRaw > 0 ? Math.round((tripleRaw / maxTripleRaw) * 15) : 7;
  score += triplePts;
  if (triplePts >= 10) reasons.push("3개 조합 출현빈도 높음");

  const quadRaw = lc(nums, 4).reduce((s, c) => s + (stats.quadFreq.get(c.join(",")) || 0), 0);
  score += maxQuadRaw > 0 ? Math.round((quadRaw / maxQuadRaw) * 10) : 5;

  const oddCount = nums.filter(n => n % 2 !== 0).length;
  const evenCount = 6 - oddCount;
  if (oddCount === 3) { score += 10; reasons.push("홀짝 3:3 균형"); }
  else if (oddCount === 2 || oddCount === 4) { score += 7; reasons.push(`홀짝 ${oddCount}:${evenCount} 양호`); }
  else if (oddCount === 1 || oddCount === 5) score += 3;

  const ranges: Record<string, number> = { "1-10": 0, "11-20": 0, "21-30": 0, "31-40": 0, "41-45": 0 };
  nums.forEach(n => {
    if (n <= 10) ranges["1-10"]++;
    else if (n <= 20) ranges["11-20"]++;
    else if (n <= 30) ranges["21-30"]++;
    else if (n <= 40) ranges["31-40"]++;
    else ranges["41-45"]++;
  });
  const covered = Object.values(ranges).filter(v => v > 0).length;
  score += Math.round((covered / 5) * 10);
  if (covered >= 4) reasons.push(`구간 ${covered}개 분포 안정적`);

  const ldFreq: Record<number, number> = {};
  nums.forEach(n => { const d = n % 10; ldFreq[d] = (ldFreq[d] || 0) + 1; });
  const maxLd = Math.max(...Object.values(ldFreq));
  if (maxLd <= 1) { score += 5; reasons.push("끝자리 중복 없음"); }
  else if (maxLd === 2) score += 3;

  let maxC = 1, curC = 1;
  for (let i = 1; i < nums.length; i++) {
    if (nums[i] === nums[i - 1] + 1) { curC++; if (curC > maxC) maxC = curC; } else curC = 1;
  }
  if (maxC <= 2) { score += 5; reasons.push("연속번호 과다 없음"); }
  else if (maxC === 3) score += 2;

  const sum = nums.reduce((a, b) => a + b, 0);
  if (Math.abs(sum - stats.sumMean) <= stats.sumStd) { score += 5; reasons.push(`합계 ${sum} (통계 범위 내)`); }
  else if (Math.abs(sum - stats.sumMean) <= 2 * stats.sumStd) { score += 2; reasons.push(`합계 ${sum}`); }
  else reasons.push(`합계 ${sum}`);

  reasons.push("과거 당첨 조합과 중복 아님");
  return { nums, score: Math.min(100, score), reasons, hotCount, coldCount, oddCount, ranges, pairScore };
}

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
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap", justifyContent: "flex-end" }}>
              <button onClick={loadDebugInfo} style={{ fontSize: "0.7rem", padding: "3px 10px", border: "1px solid #1565c0", borderRadius: 8, background: "#fff", color: "#1565c0", cursor: "pointer", fontWeight: 700 }}>새로고침</button>
              <button onClick={() => {
                localStorage.removeItem("lotto_holidayMapUpdatedAt");
                localStorage.removeItem("lotto_holidayMap");
                localStorage.removeItem("lotto_holidayFileName");
                alert("로컬 캐시를 삭제했습니다. 근무표 페이지를 열면 서버 데이터가 자동으로 적용됩니다.");
                loadDebugInfo();
              }} style={{ fontSize: "0.7rem", padding: "3px 10px", border: "1px solid #d32f2f", borderRadius: 8, background: "#fff", color: "#d32f2f", cursor: "pointer", fontWeight: 700 }}>캐시 초기화</button>
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
  const [stats, setStats] = useState<LottoStats | null>(null);
  const [games, setGames] = useState<ScoredGame[]>([]);
  const [tab, setTab] = useState<"rec" | "stats">("rec");
  const [log, setLog] = useState("");
  const [input, setInput] = useState("");
  const [generating, setGenerating] = useState(false);

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
      const rows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[wb.SheetNames[0]], { header: 1 });
      rows.slice(1).forEach((row) => {
        if (!Array.isArray(row)) return;
        const nums = (row as unknown[]).slice(2, 8)
          .map(Number).filter((v) => Number.isInteger(v) && v >= 1 && v <= 45).sort((a, b) => a - b);
        if (nums.length === 6) { pastWinners.current.add(nums.join(",")); allDraws.current.push(nums); }
      });
      setStats(computeLottoStats(allDraws.current));
    } catch (e) {
      console.warn("lotto.xlsx 로드 실패 — 기본 모드로 실행:", e);
    }
  }

  function loadLatest() {
    const list: string[] = (() => {
      try { return JSON.parse(localStorage.getItem("lotto_userRegistered") ?? "[]") as string[]; } catch { return []; }
    })();
    const legacy = localStorage.getItem("latestLotto");
    if (legacy && !list.includes(legacy)) list.unshift(legacy);
    list.forEach(key => {
      const arr = key.split(",").map(Number);
      if (arr.length === 6 && !pastWinners.current.has(key)) {
        pastWinners.current.add(key);
        allDraws.current.unshift(arr);
      }
    });
  }

  function addLatest() {
    const nums = input.split(/[\s,]+/).map((n) => Number(n.trim())).filter((n) => n >= 1 && n <= 45).sort((a, b) => a - b);
    if (nums.length < 6 || nums.length > 7) { alert("6개 또는 7개(보너스 포함) 숫자를 입력해주세요."); return; }
    nums.splice(6);
    const key = nums.join(",");
    if (pastWinners.current.has(key)) { alert("이미 등록된 숫자입니다."); return; }
    pastWinners.current.add(key); allDraws.current.unshift(nums);
    const list: string[] = (() => {
      try { return JSON.parse(localStorage.getItem("lotto_userRegistered") ?? "[]") as string[]; } catch { return []; }
    })();
    if (!list.includes(key)) list.unshift(key);
    localStorage.setItem("lotto_userRegistered", JSON.stringify(list));
    localStorage.setItem("latestLotto", key);
    setInput(""); alert("추가 완료");
  }

  function rand() { return Math.floor(Math.random() * 45) + 1; }

  function safePick(base: number[], source: number[], count: number): number[] {
    const pool = source.length >= count ? source : Array.from({ length: 45 }, (_, i) => i + 1);
    const result: number[] = [];
    let guard = 0;
    while (result.length < count && guard++ < 10000) {
      const n = pool[Math.floor(Math.random() * pool.length)];
      if (!base.includes(n) && !result.includes(n)) result.push(n);
    }
    return result;
  }

  function isValid(nums: number[]): boolean {
    if (nums.length !== 6 || new Set(nums).size !== 6) return false;
    if (nums.some((n) => !Number.isInteger(n) || n < 1 || n > 45)) return false;
    if (pastWinners.current.has(nums.join(","))) return false;
    const ld: Record<number, number> = {};
    nums.forEach(n => { const d = n % 10; ld[d] = (ld[d] || 0) + 1; });
    if (Object.values(ld).some(v => v >= 3)) return false;
    let cnt = 1;
    for (let i = 1; i < nums.length; i++) {
      if (nums[i] === nums[i - 1] + 1) { cnt++; if (cnt > 3) return false; } else cnt = 1;
    }
    const odd = nums.filter(n => n % 2).length;
    return odd >= 2 && odd <= 4;
  }

  function makeCandidate(hot: number[], cold: number[], balanced: boolean): number[] | null {
    const nums: number[] = [];
    if (hot.length >= 2 && cold.length >= 2) {
      if (balanced) {
        nums.push(...safePick(nums, hot, 2));
        nums.push(...safePick(nums, cold, 2));
      } else {
        nums.push(...safePick(nums, hot, Math.min(3, hot.length)));
        nums.push(...safePick(nums, cold, 1));
      }
    }
    let inner = 0;
    while (nums.length < 6 && inner++ < 500) { const n = rand(); if (!nums.includes(n)) nums.push(n); }
    nums.sort((a, b) => a - b);
    return isValid(nums) ? nums : null;
  }

  function generateTopGames(count: 1 | 3 | 5) {
    if (generating) return;
    setGenerating(true);
    setLog("AI 분석 중...");
    const currentStats = stats ?? (allDraws.current.length >= 2 ? computeLottoStats(allDraws.current) : null);
    if (!currentStats) { setLog("데이터 로딩 중입니다. 잠시 후 다시 시도해주세요."); setGenerating(false); return; }
    const hot = currentStats.hot30;
    const cold = currentStats.cold30;

    const raw: { nums: number[]; pairRaw: number; tripleRaw: number; quadRaw: number }[] = [];
    let attempts = 0;
    while (raw.length < 1000 && attempts++ < 8000) {
      const candidate = makeCandidate(hot, cold, Math.random() > 0.4);
      if (!candidate) continue;
      const pairRaw = lc(candidate, 2).reduce((s, c) => s + (currentStats.pairFreq.get(c.join(",")) || 0), 0);
      const tripleRaw = lc(candidate, 3).reduce((s, c) => s + (currentStats.tripleFreq.get(c.join(",")) || 0), 0);
      const quadRaw = lc(candidate, 4).reduce((s, c) => s + (currentStats.quadFreq.get(c.join(",")) || 0), 0);
      raw.push({ nums: candidate, pairRaw, tripleRaw, quadRaw });
    }

    if (raw.length === 0) { setLog("후보 생성 실패. 다시 시도해주세요."); setGenerating(false); return; }

    const maxPR = Math.max(...raw.map(c => c.pairRaw));
    const maxTR = Math.max(...raw.map(c => c.tripleRaw));
    const maxQR = Math.max(...raw.map(c => c.quadRaw));
    const scored = raw.map(c => scoreLottoCandidate(c.nums, currentStats, maxPR, maxTR, maxQR));
    scored.sort((a, b) => b.score - a.score);

    const selected: ScoredGame[] = [];
    const used = new Set<string>();
    for (const g of scored) {
      const key = g.nums.join(",");
      if (used.has(key)) continue;
      used.add(key);
      selected.push(g);
      if (selected.length >= count) break;
    }

    setGames(selected);
    setLog(`${raw.length}개 후보 분석 → AI 점수 상위 ${count}게임 추천`);
    setGenerating(false);
  }

  const ballColor = (n: number) => {
    if (n <= 10) return "#f9a825";
    if (n <= 20) return "#1e88e5";
    if (n <= 30) return "#e53935";
    if (n <= 40) return "#6d4c41";
    return "#43a047";
  };

  const BallRow = ({ nums }: { nums: number[] }) => (
    <div style={{ display: "flex", justifyContent: "center", gap: 7, flexWrap: "nowrap", overflowX: "auto" }}>
      {nums.map(n => (
        <div key={n} style={{
          width: 38, height: 38, minWidth: 38, borderRadius: "50%",
          display: "flex", alignItems: "center", justifyContent: "center",
          color: "#fff", fontWeight: 800, fontSize: "0.88rem",
          background: ballColor(n), boxShadow: `0 2px 6px ${ballColor(n)}88`,
          flexShrink: 0,
        }}>{n}</div>
      ))}
    </div>
  );

  const BTN_COLORS = [
    "linear-gradient(135deg,#f0b429 0%,#d08000 100%)",
    "linear-gradient(135deg,#f7a55a 0%,#d06010 100%)",
    "linear-gradient(135deg,#43a047 0%,#2e7d32 100%)",
  ];

  return (
    <div style={{
      fontFamily: "'Noto Sans KR','Pretendard','Apple SD Gothic Neo',sans-serif",
      background: "#eef1f8", minHeight: "100dvh",
      display: "flex", flexDirection: "column",
    }}>
      {/* 헤더 */}
      <div style={{
        display: "flex", alignItems: "center", gap: 10,
        padding: "16px 18px 12px", borderBottom: "1px solid #e5e7eb", background: "#fff",
      }}>
        <button onClick={() => setLocation(`${base}/`)} style={{
          padding: "8px 14px", background: "#f0f0f0", border: "none",
          borderRadius: 10, cursor: "pointer", fontSize: "0.9rem", color: "#555", fontWeight: 600,
        }}>← 홈</button>
        <div style={{ flex: 1, textAlign: "center", fontSize: "1.1rem", fontWeight: 800, color: "#1a1a2e" }}>
          써니의 로또 추천🚀
        </div>
        <div style={{ width: 60 }} />
      </div>

      {/* 탭 */}
      <div style={{ display: "flex", background: "#fff", borderBottom: "1px solid #e5e7eb" }}>
        {(["rec", "stats"] as const).map(t => (
          <button key={t} onClick={() => setTab(t)} style={{
            flex: 1, padding: "12px 0", border: "none", background: "transparent",
            cursor: "pointer", fontWeight: 700, fontSize: "0.92rem",
            color: tab === t ? "#7c6ef7" : "#9ca3af",
            borderBottom: tab === t ? "2.5px solid #7c6ef7" : "2.5px solid transparent",
            transition: "all 0.15s",
          }}>{t === "rec" ? "🎰 번호 추천" : "📊 통계"}</button>
        ))}
      </div>

      {/* 본문 */}
      <div style={{ flex: 1, padding: "16px 16px 100px", display: "flex", flexDirection: "column", gap: 14, overflowY: "auto" }}>

        {/* ── 추천 탭 ── */}
        {tab === "rec" && (<>

          {/* 번호 입력 */}
          <div style={{ background: "#fff", borderRadius: 18, padding: "16px", border: "1px solid #e5e7eb", boxShadow: "0 2px 8px rgba(0,0,0,0.06)" }}>
            <div style={{ fontSize: "0.75rem", color: "#9ca3af", marginBottom: 8, fontWeight: 600 }}>이번주 당첨번호 등록</div>
            <input value={input} onChange={(e) => setInput(e.target.value)}
              placeholder="예: 3 11 15 29 35 44"
              style={{
                width: "100%", boxSizing: "border-box", padding: "13px 14px", borderRadius: 12,
                border: "1px solid #d1d5db", background: "#f9fafb", color: "#1a1a2e", fontSize: "1rem", outline: "none",
              }} />
            <button onClick={addLatest} style={{
              width: "100%", marginTop: 10, padding: "14px 0", borderRadius: 12, border: "none",
              background: "linear-gradient(135deg,#7c6ef7 0%,#5b4de8 100%)",
              color: "#fff", fontWeight: 700, fontSize: "1rem", cursor: "pointer",
              boxShadow: "0 4px 14px rgba(124,110,247,0.30)",
            }}>번호 추가</button>
          </div>

          {/* AI 통계 요약 */}
          {stats && (
            <div style={{
              background: "#fffbeb", borderRadius: 12, padding: "10px 14px",
              border: "1px solid #fde68a", fontSize: "0.78rem", color: "#92400e", lineHeight: 1.8,
            }}>
              <div>🔥 핫(30회): {stats.hot30.join(", ")}</div>
              <div>❄️ 콜드(30회): {stats.cold30.join(", ")}</div>
              <div style={{ marginTop: 4, color: "#64748b" }}>
                📈 합계 평균 {Math.round(stats.sumMean)} ± {Math.round(stats.sumStd)} · 데이터 {allDraws.current.length}회차
              </div>
            </div>
          )}

          {/* 생성 버튼 */}
          <div style={{ display: "flex", gap: 10 }}>
            {([5, 3, 1] as const).map((n, idx) => (
              <button key={n} onClick={() => generateTopGames(n)} disabled={generating} style={{
                flex: 1, padding: "16px 0", borderRadius: 16, border: "none",
                background: generating ? "#ccc" : BTN_COLORS[idx],
                color: "#fff", fontWeight: 800, fontSize: "1.1rem",
                cursor: generating ? "not-allowed" : "pointer", letterSpacing: 1,
                boxShadow: generating ? "none" : "0 4px 14px rgba(0,0,0,0.15)",
                transition: "all 0.15s",
              }}>{generating ? "분석중..." : `${n}게임`}</button>
            ))}
          </div>

          {/* 결과 카드 */}
          {games.length > 0 && (
            <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
              {games.map((game, i) => (
                <div key={i} style={{
                  background: "#fff", borderRadius: 18, padding: "16px",
                  border: "1px solid #e5e7eb", boxShadow: "0 2px 10px rgba(0,0,0,0.07)",
                }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
                    <div style={{ fontSize: "0.75rem", color: "#9ca3af", fontWeight: 700 }}>게임 {i + 1}</div>
                    <div style={{
                      background: game.score >= 80 ? "#4caf50" : game.score >= 65 ? "#f0b429" : "#ef5350",
                      color: "#fff", fontWeight: 800, fontSize: "0.82rem",
                      padding: "3px 12px", borderRadius: 20,
                    }}>AI {game.score}점</div>
                  </div>
                  <BallRow nums={game.nums} />
                  <div style={{
                    marginTop: 12, padding: "10px 12px", background: "#f8f9ff",
                    borderRadius: 12, fontSize: "0.72rem", color: "#4b5563", lineHeight: 1.9,
                  }}>
                    {game.reasons.map((r, j) => <div key={j}>✓ {r}</div>)}
                  </div>
                  <div style={{ marginTop: 10, display: "flex", flexWrap: "wrap", gap: 6 }}>
                    {[
                      `🔥핫 ${game.hotCount}개`,
                      `❄️콜드 ${game.coldCount}개`,
                      `홀 ${game.oddCount}:짝 ${6 - game.oddCount}`,
                      `구간 ${Object.values(game.ranges).filter(v => v > 0).length}/5`,
                      `쌍점수 ${game.pairScore}pt`,
                    ].map((label, j) => (
                      <span key={j} style={{
                        background: "#eef1f8", borderRadius: 8,
                        padding: "3px 9px", fontSize: "0.68rem", color: "#374151", fontWeight: 600,
                      }}>{label}</span>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          )}

          {log && (
            <div style={{ textAlign: "center", fontSize: "0.78rem", color: "#9ca3af", marginTop: 2 }}>{log}</div>
          )}

          <div style={{ textAlign: "center", fontSize: "0.67rem", color: "#b0b8ca", padding: "6px 0", lineHeight: 1.6 }}>
            로또는 무작위 추첨이므로 본 추천은 통계 기반 참고용입니다.
          </div>
        </>)}

        {/* ── 통계 탭 ── */}
        {tab === "stats" && !stats && (
          <div style={{ textAlign: "center", color: "#9ca3af", paddingTop: 40 }}>통계 데이터를 불러오는 중...</div>
        )}

        {tab === "stats" && stats && (<>

          <div style={{ background: "#fff", borderRadius: 16, padding: "16px", border: "1px solid #e5e7eb" }}>
            <div style={{ fontWeight: 800, fontSize: "0.88rem", color: "#1a1a2e", marginBottom: 12 }}>🔥 최근 10회 핫번호</div>
            <BallRow nums={stats.hot10} />
          </div>

          <div style={{ background: "#fff", borderRadius: 16, padding: "16px", border: "1px solid #e5e7eb" }}>
            <div style={{ fontWeight: 800, fontSize: "0.88rem", color: "#1a1a2e", marginBottom: 12 }}>🔥 최근 30회 핫번호</div>
            <BallRow nums={stats.hot30} />
            <div style={{ fontWeight: 800, fontSize: "0.88rem", color: "#1a1a2e", margin: "14px 0 12px" }}>❄️ 최근 30회 콜드번호</div>
            <BallRow nums={stats.cold30} />
          </div>

          {[
            { title: "🔗 번호쌍 TOP 20", data: stats.pairTop20, color: "#7c6ef7" },
            { title: "🔗🔗 3개 조합 TOP 20", data: stats.tripleTop20, color: "#f0b429" },
            { title: "🔗🔗🔗 4개 조합 TOP 20", data: stats.quadTop20, color: "#43a047" },
          ].map(({ title, data, color }) => (
            <div key={title} style={{ background: "#fff", borderRadius: 16, padding: "16px", border: "1px solid #e5e7eb" }}>
              <div style={{ fontWeight: 800, fontSize: "0.88rem", color: "#1a1a2e", marginBottom: 12 }}>{title}</div>
              <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                {data.map(([key, cnt], i) => (
                  <div key={key} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: "0.78rem" }}>
                    <span style={{ color: "#9ca3af", width: 20, textAlign: "right", flexShrink: 0 }}>{i + 1}</span>
                    <div style={{ display: "flex", gap: 4, flexShrink: 0 }}>
                      {key.split(",").map(Number).map(n => (
                        <span key={n} style={{
                          width: 26, height: 26, borderRadius: "50%",
                          background: ballColor(n), color: "#fff",
                          display: "inline-flex", alignItems: "center", justifyContent: "center",
                          fontWeight: 800, fontSize: "0.68rem",
                        }}>{n}</span>
                      ))}
                    </div>
                    <div style={{ flex: 1, background: "#eef1f8", borderRadius: 6, height: 7, overflow: "hidden" }}>
                      <div style={{ height: "100%", width: `${(cnt / (data[0]?.[1] ?? 1)) * 100}%`, background: color, borderRadius: 6 }} />
                    </div>
                    <span style={{ color: "#374151", fontWeight: 700, minWidth: 30, textAlign: "right", flexShrink: 0 }}>{cnt}회</span>
                  </div>
                ))}
              </div>
            </div>
          ))}

          <div style={{ textAlign: "center", fontSize: "0.67rem", color: "#b0b8ca", padding: "6px 0", lineHeight: 1.6 }}>
            로또는 무작위 추첨이므로 본 통계는 참고용입니다.
          </div>
        </>)}
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
