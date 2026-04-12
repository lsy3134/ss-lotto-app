import { useEffect, useRef, useState } from "react";
import * as XLSX from "xlsx";

const BASE_URL = import.meta.env.BASE_URL;

interface Game {
  type: string;
  nums: number[];
}

export default function App() {
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
      .split(",")
      .map((n) => Number(n.trim()))
      .filter((n) => n >= 1 && n <= 45)
      .sort((a, b) => a - b);

    if (nums.length !== 6) {
      alert("입력 오류");
      return;
    }

    const key = nums.join(",");

    if (pastWinners.current.has(key)) {
      alert("이미 존재");
      return;
    }

    pastWinners.current.add(key);
    allDraws.current.unshift(nums);
    localStorage.setItem("latestLotto", key);
    setInput("");
    alert("추가 완료");
  }

  function rand() {
    return Math.floor(Math.random() * 45) + 1;
  }

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

  function getHotCold(): { hot: number[]; cold: number[] } {
    const freq: Record<number, number> = {};
    allDraws.current.slice(0, 10).forEach((draw) => {
      draw.forEach((n) => {
        freq[n] = (freq[n] || 0) + 1;
      });
    });

    const sorted = Object.entries(freq)
      .sort((a, b) => Number(b[1]) - Number(a[1]))
      .map((x) => Number(x[0]));

    return {
      hot: sorted.slice(0, 6),
      cold: sorted.slice(-6),
    };
  }

  function generateBalanced(hot: number[], cold: number[]): number[] {
    while (true) {
      const nums: number[] = [];
      nums.push(...pick(hot, 2));
      nums.push(...pick(cold, 2));
      while (nums.length < 6) {
        const n = rand();
        if (!nums.includes(n)) nums.push(n);
      }
      nums.sort((a, b) => a - b);
      if (isValid(nums)) return nums;
    }
  }

  function generateGreedy(cold: number[]): number[] {
    while (true) {
      const nums: number[] = [];
      nums.push(...pick(cold, 4));
      while (nums.length < 6) {
        const n = rand();
        if (n >= 31 && !nums.includes(n)) nums.push(n);
      }
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
    <div style={{ fontFamily: "sans-serif", textAlign: "center", background: "#f4f7f9", minHeight: "100vh", padding: "20px" }}>
      <div style={{ background: "white", padding: "20px", borderRadius: "15px", marginTop: "50px", display: "inline-block", minWidth: "360px" }}>
        <h2>🚀 끝판왕 로또 추천기</h2>

        <input
          value={input}
          onChange={(e) => setInput(e.target.value)}
          placeholder="이번주 번호 입력 (예: 3,11,15,29,35,44)"
          style={{ padding: "10px", width: "80%", marginTop: "10px", border: "1px solid #ccc", borderRadius: "6px" }}
        />
        <button
          onClick={addLatest}
          style={{ padding: "12px", marginTop: "10px", width: "100%", background: "#222", color: "white", border: "none", borderRadius: "8px", cursor: "pointer" }}
        >
          이번주 번호 추가
        </button>

        {hotText && (
          <div style={{ margin: "12px 0", color: "#555", fontSize: "0.9rem" }}>{hotText}</div>
        )}

        <button
          onClick={generate}
          style={{ padding: "12px", marginTop: "10px", width: "100%", background: "#222", color: "white", border: "none", borderRadius: "8px", cursor: "pointer" }}
        >
          5게임 생성
        </button>

        <div style={{ marginTop: "10px" }}>
          {games.map((game, i) => (
            <div key={i}>
              <span style={{ fontSize: "0.8rem", color: "#888" }}>[{game.type}]</span>
              <div style={{ display: "flex", justifyContent: "center", gap: "8px", margin: "6px 0 10px" }}>
                {game.nums.map((n) => (
                  <div
                    key={n}
                    style={{
                      width: "40px", height: "40px", borderRadius: "50%",
                      display: "flex", alignItems: "center", justifyContent: "center",
                      color: "white", fontWeight: "bold", background: "#333",
                    }}
                  >
                    {n}
                  </div>
                ))}
              </div>
            </div>
          ))}
        </div>

        {log && (
          <div style={{ marginTop: "8px", fontSize: "0.85rem", color: "#666" }}>{log}</div>
        )}
      </div>
    </div>
  );
}
