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

// 6개 카테고리
interface DayResult {
  twoRound: string[];  // 투라운드 (찾근)
  shift1: string[];    // 1부 (순번 자동배정)
  spare1: string[];    // 1부 스페어 (조출)
  shift2: string[];    // 2부 (순번 자동배정 + 후출)
  spare2: string[];    // 2부 스페어 (스페어 체크)
  excluded: string[];  // 제외 (당번/병가/휴무/하우스)
}

// ── 배정 엔진 ─────────────────────────────────────
function assignShifts(
  names: string[],
  statuses: Record<string, StatusType>,
  teamSize: number
): DayResult {
  const twoRound: string[] = [];
  const shift1: string[] = [];
  const spare1: string[] = [];
  const shift2: string[] = [];
  const spare2: string[] = [];
  const excluded: string[] = [];
  const autoQueue: string[] = [];

  let choCount = 0;

  for (const name of names) {
    const s = statuses[name] ?? null;

    if (s === "찾근") {
      twoRound.push(name); // 투라운드 (1부+2부 모두 표시)
    } else if (s === "조출") {
      if (choCount < 4) { spare1.push(name); choCount++; }
      else shift1.push(name); // 초과 시 일반 1부로
    } else if (s === "후출") {
      shift2.push(name);
    } else if (EXCLUDED_SET.has(s ?? "")) {
      excluded.push(name);
    } else if (s === "스페어") {
      spare2.push(name);
    } else {
      autoQueue.push(name);
    }
  }

  // 순번 자동배정: 팀수 - (찾근수 + 조출수) 만큼 1부로
  const usedSlots = twoRound.length + spare1.length + shift1.length;
  const remaining = Math.max(0, teamSize - usedSlots);
  autoQueue.forEach((n, i) => {
    if (i < remaining) shift1.push(n);
    else shift2.push(n);
  });

  return { twoRound, shift1, spare1, shift2, spare2, excluded };
}

// ── 메인 컴포넌트 ─────────────────────────────────
export default function SchedulePage() {
  const [, setLocation] = useLocation();

  const [nameText, setNameText] = useState("");
  const [teamSize, setTeamSize] = useState(5);
  const [names, setNames] = useState<string[]>([]);
  const [statuses, setStatuses] = useState<Record<string, StatusType>>({});

  const [dayResult, setDayResult] = useState<DayResult | null>(null);
  const [weekly, setWeekly] = useState<{ day: string; result: DayResult }[]>([]);
  const [view, setView] = useState<"input" | "assign">("input");

  // 이름 확정
  function confirmNames() {
    const parsed = nameText.split("\n").map((n) => n.trim()).filter(Boolean);
    if (!parsed.length) return;
    setNames(parsed);
    setStatuses({});
    setDayResult(null);
    setWeekly([]);
    setView("assign");
  }

  // 상태 토글
  function toggleStatus(name: string, s: StatusType) {
    setStatuses((prev) => ({ ...prev, [name]: prev[name] === s ? null : s }));
  }

  // 찾근 가능 여부
  function canChakgeun(name: string): boolean {
    const active = names.filter((n) => !EXCLUDED_SET.has(statuses[n] ?? ""));
    const myRank = active.indexOf(name);
    return myRank >= teamSize;
  }

  // 1일 배정
  function assign() {
    const result = assignShifts(names, statuses, teamSize);
    setDayResult(result);
    setWeekly([]);
  }

  // 일주일 생성 — 같은 순번·상태 7일 반복
  function generateWeek() {
    const DAY_LABELS = ["월", "화", "수", "목", "금", "토", "일"];
    const results = DAY_LABELS.map((day) => ({
      day,
      result: assignShifts(names, statuses, teamSize),
    }));
    setWeekly(results);
    setDayResult(null);
  }

  // ── 렌더 ────────────────────────────────────────
  return (
    <div style={S.page}>
      {/* 헤더 */}
      <div style={S.header}>
        <button onClick={() => setLocation(`${BASE}/`)} style={S.backBtn}>←</button>
        <span style={S.headerTitle}>📅 근무표</span>
        {view === "assign" && (
          <button onClick={() => { setView("input"); setDayResult(null); setWeekly([]); }} style={S.smallBtn}>
            다시 입력
          </button>
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
      {view === "assign" && (
        <>
          <div style={S.card}>
            <div style={S.infoRow}>
              <span style={S.label}>인원 {names.length}명</span>
              <span style={S.label}>1부 팀수 {teamSize}명</span>
            </div>

            {names.map((name, idx) => {
              const s = statuses[name] ?? null;
              return (
                <div key={name} style={S.personRow}>
                  <span style={S.personNum}>{idx + 1}</span>
                  <span style={S.personName}>{name}</span>
                  <div style={S.btnGroup}>
                    {STATUS_BUTTONS.map((btn) => {
                      const active = s === btn;
                      const isChakgeun = btn === "찾근";
                      const disabled = isChakgeun && !canChakgeun(name) && s !== "찾근";
                      const col = active ? STATUS_COLOR[btn!] : null;
                      return (
                        <button
                          key={btn}
                          disabled={disabled}
                          onClick={() => toggleStatus(name, btn)}
                          style={{
                            ...S.statusBtn,
                            background: active ? col!.bg : "#f0f0f0",
                            color: active ? col!.color : disabled ? "#ccc" : "#555",
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

            <div style={{ display: "flex", gap: "8px", marginTop: "16px" }}>
              <button onClick={assign} style={{ ...S.primaryBtn, flex: 1 }}>배정하기</button>
              <button onClick={generateWeek} style={{ ...S.primaryBtn, flex: 1, background: "#374151" }}>
                일주일 생성
              </button>
            </div>
          </div>

          {/* 1일 결과 */}
          {dayResult && weekly.length === 0 && (
            <div style={S.card}>
              <div style={S.sectionTitle}>📋 오늘 배정 결과</div>
              <DayResultView result={dayResult} />
            </div>
          )}

          {/* 주간 결과 */}
          {weekly.length > 0 && (
            <div style={S.card}>
              <div style={S.sectionTitle}>📅 주간 근무표</div>
              {weekly.map(({ day, result: r }) => (
                <div key={day} style={S.weekDay}>
                  <div style={S.dayChip}>{day}</div>
                  <div style={S.weekContent}>
                    <DayResultView result={r} compact />
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

// ── 6카테고리 결과 컴포넌트 ─────────────────────────
const CATEGORIES = [
  {
    key: "twoRound" as const,
    label: "투라운드",
    badge: { bg: "#e0f7fa", color: "#00838f" },
  },
  {
    key: "shift1" as const,
    label: "1부",
    badge: { bg: "#e3f2fd", color: "#1565c0" },
  },
  {
    key: "spare1" as const,
    label: "1부 스페어",
    badge: { bg: "#fff3e0", color: "#e65100" },
  },
  {
    key: "shift2" as const,
    label: "2부",
    badge: { bg: "#e8f5e9", color: "#2e7d32" },
  },
  {
    key: "spare2" as const,
    label: "2부 스페어",
    badge: { bg: "#f3e5f5", color: "#6a1b9a" },
  },
  {
    key: "excluded" as const,
    label: "제외",
    badge: { bg: "#f5f5f5", color: "#9e9e9e" },
  },
] as const;

function DayResultView({ result, compact = false }: { result: DayResult; compact?: boolean }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: compact ? "4px" : "10px" }}>
      {CATEGORIES.map(({ key, label, badge }) => {
        const people = result[key];
        if (people.length === 0) return null;
        return (
          <div key={key} style={{ display: "flex", alignItems: "flex-start", gap: "8px" }}>
            <span
              style={{
                display: "inline-block",
                padding: compact ? "1px 7px" : "3px 10px",
                borderRadius: "12px",
                fontSize: compact ? "0.7rem" : "0.75rem",
                fontWeight: 700,
                background: badge.bg,
                color: badge.color,
                flexShrink: 0,
                marginTop: "1px",
                whiteSpace: "nowrap",
              }}
            >
              {label}
            </span>
            <span
              style={{
                fontSize: compact ? "0.8rem" : "0.88rem",
                color: "#333",
                lineHeight: 1.6,
                flex: 1,
              }}
            >
              {people.join("  ·  ")}
            </span>
          </div>
        );
      })}
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
  headerTitle: { fontWeight: 700, fontSize: "1rem" },
  card: {
    background: "white",
    borderRadius: "14px",
    padding: "16px",
    margin: "12px",
    boxShadow: "0 2px 8px rgba(0,0,0,0.06)",
  },
  label: {
    display: "block",
    fontSize: "0.75rem",
    color: "#888",
    marginBottom: "6px",
    fontWeight: 600,
    textTransform: "uppercase",
    letterSpacing: "0.04em",
  },
  infoRow: {
    display: "flex",
    justifyContent: "space-between",
    marginBottom: "8px",
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
  weekDay: {
    display: "flex",
    gap: "10px",
    padding: "12px 0",
    borderBottom: "1px solid #f0f0f0",
    alignItems: "flex-start",
  },
  dayChip: {
    minWidth: "28px",
    height: "28px",
    borderRadius: "8px",
    background: "#1a1a2e",
    color: "white",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontWeight: 700,
    fontSize: "0.85rem",
    flexShrink: 0,
    marginTop: "1px",
  },
  weekContent: {
    flex: 1,
  },
};
