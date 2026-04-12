import { useState } from "react";
import { useLocation } from "wouter";

const BASE = import.meta.env.BASE_URL.replace(/\/$/, "");

// ── 타입 ──────────────────────────────────────────
type StatusType =
  | "조출" | "후출" | "찾근"
  | "당번" | "병가" | "휴무" | "하우스"
  | "스페어" | null;

const STATUS_BUTTONS: StatusType[] = [
  "조출", "후출", "찾근", "당번", "병가", "휴무", "하우스", "스페어",
];

const EXCLUDED_SET = new Set(["당번", "병가", "휴무", "하우스"]);

const STATUS_COLOR: Record<string, { bg: string; color: string }> = {
  조출:   { bg: "#ff6b35", color: "#fff" },
  후출:   { bg: "#2196f3", color: "#fff" },
  찾근:   { bg: "#00bcd4", color: "#fff" },
  당번:   { bg: "#e53935", color: "#fff" },
  병가:   { bg: "#9e9e9e", color: "#fff" },
  휴무:   { bg: "#bdbdbd", color: "#555" },
  하우스: { bg: "#f9a825", color: "#fff" },
  스페어: { bg: "#7b1fa2", color: "#fff" },
};

interface DayResult {
  shift1: string[];
  shift2: string[];
  spares: string[];
  excluded: string[];
}

// ── 배정 엔진 ─────────────────────────────────────
function assignShifts(
  names: string[],
  statuses: Record<string, StatusType>,
  teamSize: number
): DayResult {
  const shift1: string[] = [];
  const shift2: string[] = [];
  const spares: string[] = [];
  const excluded: string[] = [];
  const autoQueue: string[] = [];

  let cho = 0; // 조출 카운트
  let hu = 0;  // 후출 카운트

  for (const name of names) {
    const s = statuses[name] ?? null;

    if (s === "조출") {
      if (cho < 4) { shift1.push(name); cho++; }
      else shift2.push(name); // 초과 시 2부
    } else if (s === "후출") {
      if (hu < 4) { shift2.push(name); hu++; }
      else shift1.push(name);
    } else if (s === "찾근") {
      shift1.push(name);
      shift2.push(name);
    } else if (EXCLUDED_SET.has(s ?? "")) {
      excluded.push(name);
    } else if (s === "스페어") {
      shift2.push(name);
      spares.push(name);
    } else {
      autoQueue.push(name);
    }
  }

  const slots = Math.max(0, teamSize - shift1.length);
  autoQueue.forEach((n, i) => {
    if (i < slots) shift1.push(n);
    else shift2.push(n);
  });

  return { shift1, shift2, spares, excluded };
}

// 마지막 스페어 기준으로 순번 회전
function rotateByLastSpare(names: string[], spares: string[]): string[] {
  if (!spares.length) return names;
  const lastSpare = spares[spares.length - 1];
  const idx = names.indexOf(lastSpare);
  if (idx === -1) return names;
  const next = (idx + 1) % names.length;
  return [...names.slice(next), ...names.slice(0, next)];
}

// ── 메인 컴포넌트 ─────────────────────────────────
export default function SchedulePage() {
  const [, setLocation] = useLocation();

  // 입력 단계
  const [nameText, setNameText] = useState("");
  const [teamSize, setTeamSize] = useState(5);
  const [names, setNames] = useState<string[]>([]);
  const [statuses, setStatuses] = useState<Record<string, StatusType>>({});

  // 결과
  const [dayResult, setDayResult] = useState<DayResult | null>(null);
  const [weekly, setWeekly] = useState<{ day: string; result: DayResult }[]>([]);

  const [view, setView] = useState<"input" | "result">("input");

  // 이름 확정
  function confirmNames() {
    const parsed = nameText.split("\n").map((n) => n.trim()).filter(Boolean);
    if (!parsed.length) return;
    setNames(parsed);
    setStatuses({});
    setDayResult(null);
    setWeekly([]);
    setView("result");
  }

  // 상태 토글
  function toggleStatus(name: string, s: StatusType) {
    setStatuses((prev) => ({ ...prev, [name]: prev[name] === s ? null : s }));
  }

  // 찾근 가능 여부 (본인 순번이 팀수 초과 또는 어제 스페어 2명 이상)
  function canChakgeun(name: string): boolean {
    const idx = names.indexOf(name); // 0-based
    const active = names.filter((n) => !EXCLUDED_SET.has(statuses[n] ?? ""));
    const myRank = active.indexOf(name); // 제외자 제외 순번
    if (myRank >= teamSize) return true;
    // 전날 스페어 2명 이상 체크
    if (weekly.length > 0) {
      const yesterday = weekly[weekly.length - 1];
      if (yesterday.result.spares.length >= 2) return true;
    }
    return false;
  }

  // 1일 배정
  function assign() {
    const result = assignShifts(names, statuses, teamSize);
    setDayResult(result);
    setWeekly([]);
  }

  // 일주일 생성
  function generateWeek() {
    const DAY_LABELS = ["월", "화", "수", "목", "금", "토", "일"];
    const results: { day: string; result: DayResult }[] = [];
    let currentNames = [...names];

    for (let d = 0; d < 7; d++) {
      // 첫날만 수동 상태 적용, 이후는 자동 배정
      const dayStatuses = d === 0 ? statuses : {};
      const res = assignShifts(currentNames, dayStatuses, teamSize);
      results.push({ day: DAY_LABELS[d], result: res });
      currentNames = rotateByLastSpare(currentNames, res.spares);
    }

    setWeekly(results);
  }

  // ── 렌더 ────────────────────────────────────────
  return (
    <div style={S.page}>
      {/* 헤더 */}
      <div style={S.header}>
        <button onClick={() => setLocation(`${BASE}/`)} style={S.backBtn}>←</button>
        <span style={S.headerTitle}>📅 근무표</span>
        {view === "result" && (
          <button onClick={() => setView("input")} style={S.smallBtn}>다시 입력</button>
        )}
      </div>

      {/* ─── 입력 단계 ─── */}
      {view === "input" && (
        <div style={S.card}>
          <label style={S.label}>이름 목록 (한 줄에 한 명, 순번 순서)</label>
          <textarea
            value={nameText}
            onChange={(e) => setNameText(e.target.value)}
            placeholder={"홍길동\n김철수\n이영희"}
            rows={8}
            style={S.textarea}
          />
          <label style={S.label}>1부 인원 수 (팀수)</label>
          <input
            type="number"
            value={teamSize}
            min={1}
            onChange={(e) => setTeamSize(Number(e.target.value))}
            style={S.numInput}
          />
          <button onClick={confirmNames} style={S.primaryBtn}>다음 →</button>
        </div>
      )}

      {/* ─── 배정 단계 ─── */}
      {view === "result" && (
        <>
          {/* 인원 리스트 + 상태 버튼 */}
          <div style={S.card}>
            <div style={S.rowBetween}>
              <span style={S.label}>인원 {names.length}명 · 1부 {teamSize}명</span>
            </div>

            {names.map((name, idx) => {
              const s = statuses[name] ?? null;
              const sc = s ? STATUS_COLOR[s] : null;
              return (
                <div key={name} style={S.personRow}>
                  <span style={S.personNum}>{idx + 1}</span>
                  <span style={S.personName}>{name}</span>
                  <div style={S.btnGroup}>
                    {STATUS_BUTTONS.map((btn) => {
                      const active = s === btn;
                      const disabled = btn === "찾근" && !canChakgeun(name) && s !== "찾근";
                      const color = active ? STATUS_COLOR[btn!] : null;
                      return (
                        <button
                          key={btn}
                          disabled={disabled}
                          onClick={() => toggleStatus(name, btn)}
                          style={{
                            ...S.statusBtn,
                            background: active ? color!.bg : "#f0f0f0",
                            color: active ? color!.color : disabled ? "#ccc" : "#555",
                            border: active ? "none" : "1px solid #e0e0e0",
                          }}
                        >
                          {btn}
                        </button>
                      );
                    })}
                  </div>
                </div>
              );
            })}

            <button onClick={assign} style={{ ...S.primaryBtn, marginTop: "16px" }}>
              배정하기
            </button>
          </div>

          {/* 1일 결과 */}
          {dayResult && (
            <div style={S.card}>
              <div style={S.sectionTitle}>📋 오늘 배정 결과</div>
              <ResultBlock label="1부" names={dayResult.shift1} color="#1565c0" />
              <ResultBlock label="2부" names={dayResult.shift2} color="#2e7d32" spares={dayResult.spares} />
              {dayResult.excluded.length > 0 && (
                <ResultBlock label="제외" names={dayResult.excluded} color="#9e9e9e" />
              )}

              <button onClick={generateWeek} style={{ ...S.primaryBtn, background: "#333", marginTop: "16px" }}>
                일주일 자동 생성
              </button>
            </div>
          )}

          {/* 주간 결과 */}
          {weekly.length > 0 && (
            <div style={S.card}>
              <div style={S.sectionTitle}>📅 주간 근무표</div>
              {weekly.map(({ day, result: r }) => (
                <div key={day} style={S.weekRow}>
                  <div style={S.dayLabel}>{day}</div>
                  <div style={S.weekShifts}>
                    <span style={S.shiftChip}>
                      <b style={{ color: "#1565c0" }}>1부</b> {r.shift1.join(" · ")}
                    </span>
                    <span style={S.shiftChip}>
                      <b style={{ color: "#2e7d32" }}>2부</b> {r.shift2.join(" · ")}
                      {r.spares.length > 0 && (
                        <span style={{ color: "#7b1fa2" }}> (스페어: {r.spares.join(", ")})</span>
                      )}
                    </span>
                    {r.excluded.length > 0 && (
                      <span style={{ ...S.shiftChip, color: "#999" }}>
                        제외: {r.excluded.join(", ")}
                      </span>
                    )}
                  </div>
                </div>
              ))}
            </div>
          )}
        </>
      )}
    </div>
  );
}

// ── 결과 블록 ─────────────────────────────────────
function ResultBlock({
  label, names, color, spares,
}: {
  label: string;
  names: string[];
  color: string;
  spares?: string[];
}) {
  return (
    <div style={{ marginBottom: "12px" }}>
      <div style={{ fontWeight: 700, color, fontSize: "0.85rem", marginBottom: "6px" }}>
        {label} ({names.length}명)
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: "6px" }}>
        {names.map((n) => {
          const isSpare = spares?.includes(n);
          return (
            <span
              key={n}
              style={{
                padding: "5px 10px",
                borderRadius: "20px",
                background: isSpare ? "#f3e5f5" : `${color}18`,
                color: isSpare ? "#7b1fa2" : color,
                fontSize: "0.85rem",
                fontWeight: isSpare ? 700 : 400,
                border: `1px solid ${isSpare ? "#ce93d8" : color}40`,
              }}
            >
              {n}{isSpare ? " ★" : ""}
            </span>
          );
        })}
        {names.length === 0 && <span style={{ color: "#bbb", fontSize: "0.8rem" }}>없음</span>}
      </div>
    </div>
  );
}

// ── 스타일 ────────────────────────────────────────
const S: Record<string, React.CSSProperties> = {
  page: {
    fontFamily: "sans-serif",
    background: "#f4f7f9",
    minHeight: "100vh",
    paddingBottom: "40px",
  },
  header: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
    padding: "14px 16px",
    background: "#1a1a2e",
    color: "white",
    position: "sticky",
    top: 0,
    zIndex: 10,
  },
  backBtn: {
    background: "rgba(255,255,255,0.15)",
    border: "none",
    color: "white",
    borderRadius: "8px",
    padding: "6px 12px",
    cursor: "pointer",
    fontSize: "1rem",
  },
  smallBtn: {
    marginLeft: "auto",
    background: "rgba(255,255,255,0.15)",
    border: "none",
    color: "white",
    borderRadius: "8px",
    padding: "6px 12px",
    cursor: "pointer",
    fontSize: "0.8rem",
  },
  headerTitle: {
    fontWeight: 700,
    fontSize: "1rem",
  },
  card: {
    background: "white",
    borderRadius: "14px",
    padding: "16px",
    margin: "12px",
    boxShadow: "0 2px 8px rgba(0,0,0,0.06)",
  },
  label: {
    display: "block",
    fontSize: "0.8rem",
    color: "#888",
    marginBottom: "8px",
    fontWeight: 600,
    textTransform: "uppercase",
    letterSpacing: "0.04em",
  },
  textarea: {
    width: "100%",
    padding: "10px",
    borderRadius: "8px",
    border: "1px solid #e0e0e0",
    fontSize: "0.95rem",
    resize: "vertical",
    fontFamily: "sans-serif",
    marginBottom: "14px",
    boxSizing: "border-box",
  },
  numInput: {
    width: "100%",
    padding: "10px",
    borderRadius: "8px",
    border: "1px solid #e0e0e0",
    fontSize: "1rem",
    marginBottom: "14px",
    boxSizing: "border-box",
  },
  primaryBtn: {
    width: "100%",
    padding: "14px",
    background: "#1a1a2e",
    color: "white",
    border: "none",
    borderRadius: "10px",
    fontSize: "1rem",
    fontWeight: 700,
    cursor: "pointer",
  },
  personRow: {
    display: "flex",
    alignItems: "flex-start",
    gap: "8px",
    padding: "10px 0",
    borderBottom: "1px solid #f5f5f5",
    flexWrap: "wrap",
  },
  personNum: {
    minWidth: "20px",
    fontSize: "0.8rem",
    color: "#bbb",
    paddingTop: "4px",
  },
  personName: {
    minWidth: "60px",
    fontWeight: 600,
    fontSize: "0.95rem",
    paddingTop: "3px",
  },
  btnGroup: {
    display: "flex",
    flexWrap: "wrap",
    gap: "4px",
    flex: 1,
  },
  statusBtn: {
    padding: "4px 8px",
    borderRadius: "6px",
    fontSize: "0.78rem",
    cursor: "pointer",
    fontWeight: 600,
    transition: "all 0.15s",
  },
  sectionTitle: {
    fontWeight: 700,
    fontSize: "0.95rem",
    marginBottom: "14px",
    color: "#333",
  },
  rowBetween: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: "8px",
  },
  weekRow: {
    display: "flex",
    gap: "10px",
    padding: "10px 0",
    borderBottom: "1px solid #f5f5f5",
    alignItems: "flex-start",
  },
  dayLabel: {
    minWidth: "24px",
    fontWeight: 700,
    color: "#1a1a2e",
    fontSize: "0.9rem",
    paddingTop: "2px",
  },
  weekShifts: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    flex: 1,
  },
  shiftChip: {
    fontSize: "0.82rem",
    color: "#444",
    lineHeight: 1.5,
  },
};
