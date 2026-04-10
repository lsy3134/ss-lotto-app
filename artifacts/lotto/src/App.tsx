import { useEffect, useRef, useState } from "react";
import * as XLSX from "xlsx";

const BASE_URL = import.meta.env.BASE_URL;

function getColor(n: number): string {
  if (n <= 10) return "#fbc400";
  if (n <= 20) return "#69c8f2";
  if (n <= 30) return "#ff7272";
  if (n <= 40) return "#aaa";
  return "#b0d840";
}

interface Game {
  type: string;
  nums: number[];
}

export default function App() {
  const pastWinners = useRef<Set<string>>(new Set());
  const recentHotNumbers = useRef<number[]>([]);
  const [hotText, setHotText] = useState("");
  const [games, setGames] = useState<Game[]>([]);
  const [log, setLog] = useState("엑셀 로딩이 완료되면 번호를 생성할 수 있습니다.");
  const [loaded, setLoaded] = useState(false);

  useEffect(() => {
    async function loadExcel() {
      try {
        const res = await fetch(`${BASE_URL}lotto.xlsx`);
        if (!res.ok) throw new Error("lotto.xlsx 파일을 찾을 수 없습니다.");
        const buf = await res.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(
          wb.Sheets[wb.SheetNames[0]]
        );

        const recent: number[] = [];

        rows.forEach((r, i) => {
          const nums = [
            r["당첨번호"], r["Unnamed: 3"], r["Unnamed: 4"],
            r["Unnamed: 5"], r["Unnamed: 6"], r["Unnamed: 7"],
          ]
            .map(Number)
            .filter((v) => Number.isInteger(v) && v >= 1 && v <= 45)
            .sort((a, b) => a - b);

          if (nums.length === 6) {
            pastWinners.current.add(nums.join(","));
            if (i < 10) recent.push(...nums);
          }
        });

        const freq: Record<number, number> = {};
        recent.forEach((n) => (freq[n] = (freq[n] || 0) + 1));
        recentHotNumbers.current = Object.entries(freq)
          .sort((a, b) => Number(b[1]) - Number(a[1]))
          .slice(0, 6)
          .map((x) => Number(x[0]));

        setHotText("🔥 최근 핫번호: " + recentHotNumbers.current.join(", "));
        setLoaded(true);
        setLog("엑셀 대조 준비 완료\n이제 과거 당첨 조합을 제외하고 번호를 생성합니다.");
      } catch (e) {
        setLog("엑셀 로딩 실패: " + (e as Error).message);
      }
    }
    loadExcel();
  }, []);

  function isValid(nums: number[]): boolean {
    if (pastWinners.current.has(nums.join(","))) return false;

    const last: Record<number, number> = {};
    nums.forEach((n) => {
      const d = n % 10;
      last[d] = (last[d] || 0) + 1;
    });
    if (Object.values(last).some((v) => v >= 3)) return false;

    let cnt = 1;
    for (let i = 1; i < nums.length; i++) {
      if (nums[i] === nums[i - 1] + 1) {
        cnt++;
        if (cnt >= 3) return false;
      } else {
        cnt = 1;
      }
    }

    const odd = nums.filter((n) => n % 2).length;
    if (odd < 2 || odd > 4) return false;

    return true;
  }

  function generateOne(): number[] | null {
    let tries = 0;
    while (tries < 5000) {
      tries++;
      const arr: number[] = [];
      while (arr.length < 6) {
        const n = Math.floor(Math.random() * 45) + 1;
        if (!arr.includes(n)) arr.push(n);
      }
      arr.sort((a, b) => a - b);
      if (isValid(arr)) return arr;
    }
    return null;
  }

  function generateBalanced(): number[] {
    while (true) {
      const nums = generateOne();
      if (!nums) continue;
      const odd = nums.filter((n) => n % 2).length;
      if (odd < 2 || odd > 4) continue;
      return nums;
    }
  }

  function generateGreedy(): number[] {
    while (true) {
      const nums = generateOne();
      if (!nums) continue;
      const high = nums.filter((n) => n >= 31).length;
      if (high < 4) continue;
      const last = new Set(nums.map((n) => n % 10));
      if (last.size < 6) continue;
      return nums;
    }
  }

  function generateFinalSet(): Game[] {
    const result: Game[] = [];
    for (let i = 0; i < 3; i++) result.push({ type: "균형형", nums: generateBalanced() });
    for (let i = 0; i < 2; i++) result.push({ type: "독식형", nums: generateGreedy() });
    return result;
  }

  function handleGenerate() {
    if (!loaded) return;
    const result = generateFinalSet();
    setGames(result);
    setLog(`완료: ${result.length}게임 생성 (균형형 3 + 독식형 2)`);
  }

  return (
    <div style={{ fontFamily: "sans-serif", textAlign: "center", background: "#f4f7f9", minHeight: "100vh", padding: "20px" }}>
      <div style={{ background: "white", padding: "20px", borderRadius: "15px", marginTop: "50px", display: "inline-block", minWidth: "340px" }}>
        <h2>🚀 스마트 2단계 추천기</h2>
        {hotText && <div style={{ margin: "8px 0", color: "#e65100", fontWeight: 600 }}>{hotText}</div>}

        <div style={{ marginTop: "12px" }}>
          {games.map((game, i) => (
            <div key={i} style={{ display: "flex", justifyContent: "center", alignItems: "center", gap: "8px", margin: "10px 0" }}>
              <span style={{ fontSize: "0.75rem", color: "#888", width: "40px", textAlign: "right" }}>{game.type}</span>
              {game.nums.map((n) => (
                <div
                  key={n}
                  style={{
                    width: "40px", height: "40px", borderRadius: "50%",
                    display: "flex", alignItems: "center", justifyContent: "center",
                    color: "white", fontWeight: "bold", background: getColor(n),
                  }}
                >
                  {n}
                </div>
              ))}
            </div>
          ))}
        </div>

        <button
          onClick={handleGenerate}
          disabled={!loaded}
          style={{
            padding: "12px", marginTop: "15px", width: "100%",
            background: loaded ? "#222" : "#999",
            color: "white", border: "none", borderRadius: "8px",
            cursor: loaded ? "pointer" : "not-allowed", fontSize: "1rem",
          }}
        >
          {loaded ? "5게임 생성" : "엑셀 로딩 중..."}
        </button>

        <div style={{ marginTop: "12px", fontSize: "0.85rem", color: "#666", whiteSpace: "pre-line" }}>
          {log}
        </div>
      </div>
    </div>
  );
}
