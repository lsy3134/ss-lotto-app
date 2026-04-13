import { useEffect, useRef, useState } from "react";
import { Switch, Route, Link, Router as WouterRouter, useLocation } from "wouter";
import * as XLSX from "xlsx";
import SchedulePage from "./pages/Schedule";

const BASE_URL = import.meta.env.BASE_URL;
const base = BASE_URL.replace(/\/$/, "");

const C = {
  bgDark:    "#1a1a2e",
  cardBg:    "#16213e",
  blue:      "#4e89ae",
  green:     "#52de97",
  yellow:    "#f8b400",
  white:     "#ffffff",
};

interface Game { type: string; nums: number[]; }

// ───────────────────────────────────────────
// 인트로 화면
// ───────────────────────────────────────────
function IntroView({ onEnter }: { onEnter: () => void }) {
  return (
    <div style={{
      backgroundColor: C.bgDark,
      minHeight: "100dvh",
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      justifyContent: "center",
      color: C.white,
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
          boxShadow: `0 0 24px ${C.blue}`,
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
          backgroundColor: C.blue,
          color: "white",
          border: "none",
          padding: "15px 52px",
          borderRadius: "30px",
          fontSize: "1.1rem",
          fontWeight: 700,
          cursor: "pointer",
          boxShadow: "0 4px 20px rgba(78,137,174,0.5)",
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
      backgroundColor: C.bgDark,
      minHeight: "100vh",
      padding: "48px 24px 32px",
      color: C.white,
    }}>
      <header style={{ marginBottom: "36px" }}>
        <h2 style={{ fontSize: "1.5rem", margin: "0 0 6px", fontWeight: 800 }}>서비스 선택</h2>
        <p style={{ margin: 0, opacity: 0.55, fontSize: "0.9rem" }}>필요한 업무를 골라주세요.</p>
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
          border: "1px solid rgba(255,255,255,0.08)",
          transition: "transform 0.15s, box-shadow 0.15s",
          boxShadow: `0 0 0 1px rgba(78,137,174,0.2)`,
        }}
          onMouseEnter={e => {
            (e.currentTarget as HTMLDivElement).style.transform = "translateY(-2px)";
            (e.currentTarget as HTMLDivElement).style.boxShadow = `0 8px 24px rgba(78,137,174,0.3)`;
          }}
          onMouseLeave={e => {
            (e.currentTarget as HTMLDivElement).style.transform = "translateY(0)";
            (e.currentTarget as HTMLDivElement).style.boxShadow = `0 0 0 1px rgba(78,137,174,0.2)`;
          }}
        >
          <img
            src={`${BASE_URL}char_dino.png`}
            alt="근무표"
            style={{ width: "60px", height: "60px", objectFit: "cover", borderRadius: "14px", marginRight: "18px", flexShrink: 0 }}
          />
          <div style={{ flex: 1, textAlign: "left" }}>
            <div style={{ fontWeight: 800, fontSize: "1.1rem", marginBottom: "4px" }}>근무표 생성</div>
            <div style={{ opacity: 0.55, fontSize: "0.83rem" }}>63명 캐디 배정 및 일정 관리</div>
          </div>
          <div style={{ fontSize: "1.5rem", opacity: 0.35, color: C.blue }}>›</div>
        </div>
      </Link>

      {/* 로또 카드 */}
      <Link href={`${base}/lotto`} style={{ textDecoration: "none" }}>
        <div style={{
          backgroundColor: "#23234a",
          borderRadius: "20px",
          padding: "22px 20px",
          display: "flex",
          alignItems: "center",
          cursor: "pointer",
          border: `1px solid ${C.yellow}33`,
          transition: "transform 0.15s, box-shadow 0.15s",
        }}
          onMouseEnter={e => {
            (e.currentTarget as HTMLDivElement).style.transform = "translateY(-2px)";
            (e.currentTarget as HTMLDivElement).style.boxShadow = `0 8px 24px rgba(248,180,0,0.2)`;
          }}
          onMouseLeave={e => {
            (e.currentTarget as HTMLDivElement).style.transform = "translateY(0)";
            (e.currentTarget as HTMLDivElement).style.boxShadow = "none";
          }}
        >
          <img
            src={`${BASE_URL}char_dragon.png`}
            alt="로또"
            style={{ width: "60px", height: "60px", objectFit: "cover", borderRadius: "14px", marginRight: "18px", flexShrink: 0 }}
          />
          <div style={{ flex: 1, textAlign: "left" }}>
            <div style={{ fontWeight: 800, fontSize: "1.1rem", marginBottom: "4px", color: C.yellow }}>나만의 로또 생성</div>
            <div style={{ opacity: 0.55, fontSize: "0.83rem" }}>오늘의 행운 번호 추출</div>
          </div>
          <div style={{ fontSize: "1.5rem", color: C.yellow, opacity: 0.5 }}>›</div>
        </div>
      </Link>

      {/* 하단 */}
      <div style={{ marginTop: "56px", textAlign: "center", opacity: 0.35 }}>
        <img
          src={`${BASE_URL}char_smile.png`}
          alt=""
          style={{ width: "36px", height: "36px", objectFit: "cover", borderRadius: "50%", marginBottom: "8px", display: "block", margin: "0 auto 8px" }}
        />
        <p style={{ fontSize: "0.72rem", margin: 0 }}>SS앱 — Caddie's Cloud Helper v1.0</p>
      </div>
    </div>
  );
}


// ───────────────────────────────────────────
// 로또 페이지
// ───────────────────────────────────────────
function LottoPage() {
  const [, setLocation] = useLocation();
  const allDraws = useRef<number[][]>([]);
  const pastWinners = useRef<Set<string>>(new Set());
  const [hotText, setHotText] = useState("");
  const [games, setGames] = useState<Game[]>([]);
  const [log, setLog] = useState("");
  const [input, setInput] = useState("");

  useEffect(() => {
    async function init() { await loadExcel(); loadLatest(); }
    init();
  }, []);

  async function loadExcel() {
    const res = await fetch(`${BASE_URL}lotto.xlsx`);
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(wb.Sheets[wb.SheetNames[0]]);
    rows.forEach((row) => {
      const nums = [row["당첨번호"], row["Unnamed: 3"], row["Unnamed: 4"], row["Unnamed: 5"], row["Unnamed: 6"], row["Unnamed: 7"]]
        .map(Number).filter((v) => Number.isInteger(v) && v >= 1 && v <= 45).sort((a, b) => a - b);
      if (nums.length === 6) { pastWinners.current.add(nums.join(",")); allDraws.current.push(nums); }
    });
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
  function pick(arr: number[], n: number): number[] {
    const copy = [...arr]; const res: number[] = [];
    while (res.length < n) { const i = Math.floor(Math.random() * copy.length); res.push(copy.splice(i, 1)[0]); }
    return res;
  }
  function isValid(nums: number[]): boolean {
    if (pastWinners.current.has(nums.join(","))) return false;
    const last: Record<number, number> = {};
    nums.forEach((n) => { const d = n % 10; last[d] = (last[d] || 0) + 1; });
    if (Object.values(last).some((v) => v >= 3)) return false;
    let cnt = 1;
    for (let i = 1; i < nums.length; i++) {
      if (nums[i] === nums[i - 1] + 1) { cnt++; if (cnt >= 3) return false; } else cnt = 1;
    }
    const odd = nums.filter((n) => n % 2).length;
    if (odd < 2 || odd > 4) return false;
    return true;
  }
  function getHotCold() {
    const freq: Record<number, number> = {};
    allDraws.current.slice(0, 10).forEach((draw) => { draw.forEach((n) => { freq[n] = (freq[n] || 0) + 1; }); });
    const sorted = Object.entries(freq).sort((a, b) => Number(b[1]) - Number(a[1])).map((x) => Number(x[0]));
    return { hot: sorted.slice(0, 6), cold: sorted.slice(-6) };
  }
  function generateBalanced(hot: number[], cold: number[]): number[] {
    while (true) {
      const nums: number[] = []; nums.push(...pick(hot, 2)); nums.push(...pick(cold, 2));
      while (nums.length < 6) { const n = rand(); if (!nums.includes(n)) nums.push(n); }
      nums.sort((a, b) => a - b); if (isValid(nums)) return nums;
    }
  }
  function generateGreedy(cold: number[]): number[] {
    while (true) {
      const nums: number[] = []; nums.push(...pick(cold, 4));
      while (nums.length < 6) { const n = rand(); if (n >= 31 && !nums.includes(n)) nums.push(n); }
      nums.sort((a, b) => a - b); if (isValid(nums)) return nums;
    }
  }
  function generate() {
    const { hot, cold } = getHotCold();
    setHotText("🔥 핫: " + hot.join(", ") + " / ❄ 콜드: " + cold.join(", "));
    const result: Game[] = [];
    for (let i = 0; i < 3; i++) result.push({ type: "균형형", nums: generateBalanced(hot, cold) });
    for (let i = 0; i < 2; i++) result.push({ type: "독식형", nums: generateGreedy(cold) });
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
      fontFamily: "'Pretendard', 'Apple SD Gothic Neo', sans-serif",
      background: "linear-gradient(160deg, #1a1a2e 0%, #16213e 60%, #0f3460 100%)",
      minHeight: "100vh",
      display: "flex",
      flexDirection: "column",
      padding: "0",
    }}>
      {/* 헤더 */}
      <div style={{
        display: "flex", alignItems: "center", gap: 10,
        padding: "16px 18px 12px",
        borderBottom: "1px solid rgba(255,255,255,0.08)",
      }}>
        <button
          onClick={() => setLocation(`${base}/`)}
          style={{
            padding: "8px 14px", background: "rgba(255,255,255,0.1)",
            border: "none", borderRadius: 10, cursor: "pointer",
            fontSize: "0.9rem", color: "#fff", fontWeight: 600,
          }}
        >← 홈</button>
        <div style={{ flex: 1, textAlign: "center" }}>
          <div style={{ fontSize: "1.1rem", fontWeight: 800, color: "#fff" }}>🚀 끝판왕 로또 추천기</div>
        </div>
        <div style={{ width: 60 }} />
      </div>

      {/* 본문 */}
      <div style={{ flex: 1, padding: "20px 16px", display: "flex", flexDirection: "column", gap: 14 }}>

        {/* 번호 입력 카드 */}
        <div style={{
          background: "rgba(255,255,255,0.06)", borderRadius: 18,
          padding: "16px", border: "1px solid rgba(255,255,255,0.1)",
        }}>
          <div style={{ fontSize: "0.75rem", color: "rgba(255,255,255,0.5)", marginBottom: 8, fontWeight: 600 }}>
            이번주 당첨번호 등록
          </div>
          <input
            value={input}
            onChange={(e) => setInput(e.target.value)}
            placeholder="예: 3 11 15 29 35 44"
            style={{
              width: "100%", boxSizing: "border-box",
              padding: "13px 14px", borderRadius: 12,
              border: "1.5px solid rgba(255,255,255,0.15)",
              background: "rgba(255,255,255,0.08)",
              color: "#fff", fontSize: "1rem",
              outline: "none",
            }}
          />
          <button
            onClick={addLatest}
            style={{
              width: "100%", marginTop: 10,
              padding: "14px 0", borderRadius: 12, border: "none",
              background: "rgba(255,255,255,0.15)",
              color: "#fff", fontWeight: 700, fontSize: "1rem",
              cursor: "pointer",
            }}
          >번호 추가</button>
        </div>

        {/* 핫/콜드 정보 */}
        {hotText && (
          <div style={{
            background: "rgba(248,180,0,0.12)", borderRadius: 12,
            padding: "10px 14px", border: "1px solid rgba(248,180,0,0.25)",
            fontSize: "0.8rem", color: "#f8b400", lineHeight: 1.6,
          }}>
            {hotText}
          </div>
        )}

        {/* 생성 버튼 */}
        <button
          onClick={generate}
          style={{
            width: "100%", padding: "18px 0", borderRadius: 16, border: "none",
            background: "linear-gradient(135deg, #f8b400 0%, #e65100 100%)",
            color: "#fff", fontWeight: 800, fontSize: "1.15rem",
            cursor: "pointer", boxShadow: "0 6px 24px rgba(248,180,0,0.35)",
            letterSpacing: 1,
          }}
        >✨ 5게임 생성</button>

        {/* 결과 */}
        {games.length > 0 && (
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            {games.map((game, i) => (
              <div key={i} style={{
                background: "rgba(255,255,255,0.07)", borderRadius: 16,
                padding: "14px 16px", border: "1px solid rgba(255,255,255,0.1)",
              }}>
                <div style={{
                  fontSize: "0.72rem", color: "rgba(255,255,255,0.4)",
                  fontWeight: 700, marginBottom: 10, letterSpacing: 0.5,
                }}>
                  게임 {i + 1} · {game.type}
                </div>
                <div style={{ display: "flex", justifyContent: "center", gap: 8, flexWrap: "wrap" }}>
                  {game.nums.map((n) => (
                    <div key={n} style={{
                      width: 46, height: 46, borderRadius: "50%",
                      display: "flex", alignItems: "center", justifyContent: "center",
                      color: "#fff", fontWeight: 800, fontSize: "1rem",
                      background: ballColor(n),
                      boxShadow: `0 3px 10px ${ballColor(n)}88`,
                    }}>{n}</div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        )}

        {log && (
          <div style={{ textAlign: "center", fontSize: "0.8rem", color: "rgba(255,255,255,0.35)", marginTop: 4 }}>
            {log}
          </div>
        )}
      </div>
    </div>
  );
}

// ───────────────────────────────────────────
// 라우터
// ───────────────────────────────────────────
export default function App() {
  return (
    <WouterRouter base={base}>
      <Switch>
        <Route path="/" component={HomePage} />
        <Route path="/schedule" component={SchedulePage} />
        <Route path="/lotto" component={LottoPage} />
        <Route component={HomePage} />
      </Switch>
    </WouterRouter>
  );
}
