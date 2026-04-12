import { useEffect, useRef, useState } from "react";
import { Switch, Route, Link, Router as WouterRouter, useLocation } from "wouter";
import * as XLSX from "xlsx";
import SchedulePage from "./pages/Schedule";

const BASE_URL = import.meta.env.BASE_URL;
const base = BASE_URL.replace(/\/$/, "");

interface Game {
  type: string;
  nums: number[];
}

// ───────────────────────────────────────────
// 홈 화면
// ───────────────────────────────────────────
function HomePage() {
  return (
    <div style={styles.page}>
      <div style={styles.homeWrap}>
        <p style={styles.homeLabel}>서비스 선택</p>

        {/* 근무표 — 메인 강조 카드 */}
        <Link href={`${base}/schedule`} style={{ textDecoration: "none" }}>
          <div style={{ ...styles.card, ...styles.cardPrimary }}>
            <span style={styles.cardIcon}>📅</span>
            <div>
              <div style={styles.cardTitle}>근무표</div>
              <div style={styles.cardDesc}>나의 근무 일정을 관리하세요</div>
            </div>
            <span style={styles.cardArrow}>›</span>
          </div>
        </Link>

        {/* 로또 생성기 — 보조 카드 */}
        <Link href={`${base}/lotto`} style={{ textDecoration: "none" }}>
          <div style={{ ...styles.card, ...styles.cardSecondary }}>
            <span style={styles.cardIcon}>🎱</span>
            <div>
              <div style={{ ...styles.cardTitle, color: "#333" }}>로또 생성기</div>
              <div style={styles.cardDesc}>끝판왕 번호를 추천해 드립니다</div>
            </div>
            <span style={{ ...styles.cardArrow, color: "#aaa" }}>›</span>
          </div>
        </Link>
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
    async function init() {
      await loadExcel();
      loadLatest();
    }
    init();
  }, []);

  async function loadExcel() {
    const res = await fetch(`${BASE_URL}lotto.xlsx`);
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(
      wb.Sheets[wb.SheetNames[0]]
    );
    rows.forEach((row) => {
      const nums = [
        row["당첨번호"], row["Unnamed: 3"], row["Unnamed: 4"],
        row["Unnamed: 5"], row["Unnamed: 6"], row["Unnamed: 7"],
      ]
        .map(Number)
        .filter((v) => Number.isInteger(v) && v >= 1 && v <= 45)
        .sort((a, b) => a - b);
      if (nums.length === 6) {
        pastWinners.current.add(nums.join(","));
        allDraws.current.push(nums);
      }
    });
  }

  function loadLatest() {
    const saved = localStorage.getItem("latestLotto");
    if (saved) {
      const arr = saved.split(",").map(Number);
      if (!pastWinners.current.has(saved)) {
        pastWinners.current.add(saved);
        allDraws.current.unshift(arr);
      }
    }
  }

  function addLatest() {
    const nums = input
      .split(/[\s,]+/)
      .map((n) => Number(n.trim()))
      .filter((n) => n >= 1 && n <= 45)
      .sort((a, b) => a - b);
    if (nums.length < 6 || nums.length > 7) {
      alert("6개 또는 7개(보너스 포함) 숫자를 입력해주세요.");
      return;
    }
    nums.splice(6);
    const key = nums.join(",");
    if (pastWinners.current.has(key)) { alert("이미 존재"); return; }
    pastWinners.current.add(key);
    allDraws.current.unshift(nums);
    localStorage.setItem("latestLotto", key);
    setInput("");
    alert("추가 완료");
  }

  function rand() { return Math.floor(Math.random() * 45) + 1; }

  function pick(arr: number[], n: number): number[] {
    const copy = [...arr];
    const res: number[] = [];
    while (res.length < n) {
      const i = Math.floor(Math.random() * copy.length);
      res.push(copy.splice(i, 1)[0]);
    }
    return res;
  }

  function isValid(nums: number[]): boolean {
    if (pastWinners.current.has(nums.join(","))) return false;
    const last: Record<number, number> = {};
    nums.forEach((n) => { const d = n % 10; last[d] = (last[d] || 0) + 1; });
    if (Object.values(last).some((v) => v >= 3)) return false;
    let cnt = 1;
    for (let i = 1; i < nums.length; i++) {
      if (nums[i] === nums[i - 1] + 1) { cnt++; if (cnt >= 3) return false; }
      else cnt = 1;
    }
    const odd = nums.filter((n) => n % 2).length;
    if (odd < 2 || odd > 4) return false;
    return true;
  }

  function getHotCold() {
    const freq: Record<number, number> = {};
    allDraws.current.slice(0, 10).forEach((draw) => {
      draw.forEach((n) => { freq[n] = (freq[n] || 0) + 1; });
    });
    const sorted = Object.entries(freq).sort((a, b) => Number(b[1]) - Number(a[1])).map((x) => Number(x[0]));
    return { hot: sorted.slice(0, 6), cold: sorted.slice(-6) };
  }

  function generateBalanced(hot: number[], cold: number[]): number[] {
    while (true) {
      const nums: number[] = [];
      nums.push(...pick(hot, 2));
      nums.push(...pick(cold, 2));
      while (nums.length < 6) { const n = rand(); if (!nums.includes(n)) nums.push(n); }
      nums.sort((a, b) => a - b);
      if (isValid(nums)) return nums;
    }
  }

  function generateGreedy(cold: number[]): number[] {
    while (true) {
      const nums: number[] = [];
      nums.push(...pick(cold, 4));
      while (nums.length < 6) { const n = rand(); if (n >= 31 && !nums.includes(n)) nums.push(n); }
      nums.sort((a, b) => a - b);
      if (isValid(nums)) return nums;
    }
  }

  function generate() {
    const { hot, cold } = getHotCold();
    setHotText("🔥 핫: " + hot.join(", ") + " / ❄ 콜드: " + cold.join(", "));
    const result: Game[] = [];
    for (let i = 0; i < 3; i++) result.push({ type: "균형형", nums: generateBalanced(hot, cold) });
    for (let i = 0; i < 2; i++) result.push({ type: "독식형", nums: generateGreedy(cold) });
    setGames(result);
    setLog("생성 완료");
  }

  return (
    <div style={styles.page}>
      <div style={{ background: "white", padding: "20px", borderRadius: "15px", marginTop: "30px", display: "inline-block", minWidth: "360px", textAlign: "center" }}>
        <div style={{ display: "flex", alignItems: "center", gap: "8px", marginBottom: "12px" }}>
          <button onClick={() => setLocation(`${base}/`)} style={styles.backBtn}>← 홈</button>
          <h2 style={{ margin: 0, flex: 1 }}>🚀 끝판왕 로또 추천기</h2>
        </div>

        <input
          value={input}
          onChange={(e) => setInput(e.target.value)}
          placeholder="번호 입력 (예: 3 11 15 29 35 44 또는 보너스 포함 7자리)"
          style={{ padding: "10px", width: "80%", border: "1px solid #ccc", borderRadius: "6px" }}
        />
        <button onClick={addLatest} style={{ ...styles.btn, marginTop: "10px" }}>이번주 번호 추가</button>

        {hotText && <div style={{ margin: "12px 0", color: "#555", fontSize: "0.9rem" }}>{hotText}</div>}

        <button onClick={generate} style={{ ...styles.btn, marginTop: "6px" }}>5게임 생성</button>

        <div style={{ marginTop: "10px" }}>
          {games.map((game, i) => (
            <div key={i}>
              <span style={{ fontSize: "0.8rem", color: "#888" }}>[{game.type}]</span>
              <div style={{ display: "flex", justifyContent: "center", gap: "8px", margin: "6px 0 10px" }}>
                {game.nums.map((n) => (
                  <div key={n} style={{ width: "40px", height: "40px", borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center", color: "white", fontWeight: "bold", background: "#333" }}>
                    {n}
                  </div>
                ))}
              </div>
            </div>
          ))}
        </div>

        {log && <div style={{ marginTop: "8px", fontSize: "0.85rem", color: "#666" }}>{log}</div>}
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

// ───────────────────────────────────────────
// 공통 스타일
// ───────────────────────────────────────────
const styles: Record<string, React.CSSProperties> = {
  page: {
    fontFamily: "sans-serif",
    textAlign: "center",
    background: "#f4f7f9",
    minHeight: "100vh",
    padding: "20px",
  },
  homeWrap: {
    background: "white",
    padding: "28px 24px",
    borderRadius: "18px",
    marginTop: "50px",
    display: "inline-block",
    minWidth: "340px",
    maxWidth: "440px",
    width: "100%",
    textAlign: "left",
  },
  homeLabel: {
    fontSize: "0.75rem",
    color: "#aaa",
    letterSpacing: "0.08em",
    textTransform: "uppercase",
    marginBottom: "14px",
  },
  card: {
    display: "flex",
    alignItems: "center",
    gap: "14px",
    padding: "18px 16px",
    borderRadius: "14px",
    marginBottom: "12px",
    cursor: "pointer",
    transition: "transform 0.15s",
  },
  cardPrimary: {
    background: "linear-gradient(135deg, #1a1a2e 0%, #16213e 100%)",
    color: "white",
    boxShadow: "0 6px 20px rgba(26,26,46,0.25)",
  },
  cardSecondary: {
    background: "#f8f9fa",
    border: "1px solid #eee",
    color: "#333",
  },
  cardIcon: {
    fontSize: "1.8rem",
    flexShrink: 0,
  },
  cardTitle: {
    fontWeight: 700,
    fontSize: "1.05rem",
    color: "white",
    marginBottom: "2px",
  },
  cardDesc: {
    fontSize: "0.8rem",
    color: "rgba(255,255,255,0.6)",
  },
  cardArrow: {
    marginLeft: "auto",
    fontSize: "1.4rem",
    color: "rgba(255,255,255,0.5)",
    flexShrink: 0,
  },
  btn: {
    padding: "12px",
    width: "100%",
    background: "#222",
    color: "white",
    border: "none",
    borderRadius: "8px",
    cursor: "pointer",
    fontSize: "1rem",
  },
  backBtn: {
    padding: "6px 14px",
    background: "#f0f0f0",
    border: "none",
    borderRadius: "8px",
    cursor: "pointer",
    fontSize: "0.85rem",
    color: "#555",
    flexShrink: 0,
  },
};
