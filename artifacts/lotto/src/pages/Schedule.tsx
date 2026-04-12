import { useState } from "react";
import { useLocation } from "wouter";

const BASE = import.meta.env.BASE_URL.replace(/\/$/, "");

// ── 타입 ──────────────────────────────────────────
type StatusType =
  | "조출" | "후출" | "찾근"
  | "당번" | "병가" | "휴무" | "하우스"
  | null;

type Mode = "복부제" | "단부제";

const STATUS_BUTTONS: StatusType[] = [
  "조출", "후출", "찾근", "당번", "병가", "휴무", "하우스",
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
};

// 6개 카테고리 결과
interface DayResult {
  twoRound: string[];  // 투라운드 (찾근)
  shift1: string[];    // 1부 / 단부
  spare1: string[];    // 1부 스페어 (복부제: 1부팀수+1번째)
  shift2: string[];    // 2부 (복부제만)
  spare2: string[];    // 2부 스페어 (복부제: 총팀수 초과) / 단부제 스페어
  excluded: string[];  // 제외
}

// ── 배정 엔진: 복부제 ─────────────────────────────
function assignDouble(
  names: string[],
  statuses: Record<string, StatusType>,
  shift1Size: number,
  shift2Size: number
): DayResult {
  const twoRound: string[] = [];
  const shift1: string[] = [];
  const spare1: string[] = [];
  const shift2: string[] = [];
  const spare2: string[] = [];
  const excluded: string[] = [];
  const autoQueue: string[] = [];
  let choCount = 0, huCount = 0;

  for (const name of names) {
    const s = statuses[name] ?? null;
    if (s === "찾근") {
      twoRound.push(name);
    } else if (s === "조출") {
      if (choCount < 4) { shift1.push(name); choCount++; }
      else autoQueue.push(name);
    } else if (s === "후출") {
      if (huCount < 4) { shift2.push(name); huCount++; }
      else autoQueue.push(name);
    } else if (EXCLUDED_SET.has(s ?? "")) {
      excluded.push(name);
    } else {
      autoQueue.push(name);
    }
  }

  // 자동 배정: 순번대로
  // 1부 가용 슬롯 = shift1Size - 조출count - 찾근count
  const avail1 = Math.max(0, shift1Size - choCount - twoRound.length);
  // 2부 가용 슬롯 = shift2Size - 후출count - 찾근count
  const avail2 = Math.max(0, shift2Size - huCount - twoRound.length);

  autoQueue.forEach((n, i) => {
    if (i < avail1) {
      shift1.push(n);                    // 1부
    } else if (i === avail1) {
      spare1.push(n);                    // 1부 스페어 (딱 1명, 1부팀수+1번째)
    } else if (i <= avail1 + avail2) {
      shift2.push(n);                    // 2부
    } else {
      spare2.push(n);                    // 2부 스페어 (총팀수 초과)
    }
  });

  return { twoRound, shift1, spare1, shift2, spare2, excluded };
}

// ── 배정 엔진: 단부제 ─────────────────────────────
function assignSingle(
  names: string[],
  statuses: Record<string, StatusType>,
  teamSize: number
): DayResult {
  const twoRound: string[] = [];
  const shift1: string[] = [];
  const spare2: string[] = [];
  const excluded: string[] = [];
  const autoQueue: string[] = [];

  for (const name of names) {
    const s = statuses[name] ?? null;
    if (s === "찾근") {
      twoRound.push(name);
    } else if (EXCLUDED_SET.has(s ?? "")) {
      excluded.push(name);
    } else {
      // 조출/후출은 단부제에서 일반 순번으로 처리
      autoQueue.push(name);
    }
  }

  const avail = Math.max(0, teamSize - twoRound.length);
  autoQueue.forEach((n, i) => {
    if (i < avail) shift1.push(n);
    else spare2.push(n);
  });

  return { twoRound, shift1, spare1: [], shift2: [], spare2, excluded };
}

// ── 메인 컴포넌트 ─────────────────────────────────
export default function SchedulePage() {
  const [, setLocation] = useLocation();

  const [nameText, setNameText] = useState("");
  const [mode, setMode] = useState<Mode>("복부제");
  const [shift1Size, setShift1Size] = useState(35);
  const [shift2Size, setShift2Size] = useState(35);
  const [singleSize, setSingleSize] = useState(60);

  const [names, setNames] = useState<string[]>([]);
  const [statuses, setStatuses] = useState<Record<string, StatusType>>({});

  const [dayResult, setDayResult] = useState<DayResult | null>(null);
  const [weekly, setWeekly] = useState<{ day: string; result: DayResult }[]>([]);
  const [view, setView] = useState<"input" | "assign">("input");

  function getResult() {
    return mode === "복부제"
      ? assignDouble(names, statuses, shift1Size, shift2Size)
      : assignSingle(names, statuses, singleSize);
  }

  function confirmNames() {
    const parsed = nameText.split("\n").map((n) => n.trim()).filter(Boolean);
    if (!parsed.length) return;
    setNames(parsed);
    setStatuses({});
    setDayResult(null);
    setWeekly([]);
    setView("assign");
  }

  function toggleStatus(name: string, s: StatusType) {
    setStatuses((prev) => ({ ...prev, [name]: prev[name] === s ? null : s }));
  }

  function canChakgeun(name: string): boolean {
    if (mode === "단부제") return false;
    const active = names.filter((n) => !EXCLUDED_SET.has(statuses[n] ?? ""));
    const myRank = active.indexOf(name);
    return myRank >= shift1Size;
  }

  function assign() {
    setDayResult(getResult());
    setWeekly([]);
  }

  function generateWeek() {
    const DAY_LABELS = ["월", "화", "수", "목", "금", "토", "일"];
    setWeekly(DAY_LABELS.map((day) => ({ day, result: getResult() })));
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
          {/* 운영 모드 */}
          <label style={S.label}>운영 방식</label>
          <div style={{ display: "flex", gap: "8px", marginBottom: "16px" }}>
            {(["복부제", "단부제"] as Mode[]).map((m) => (
              <button
                key={m}
                onClick={() => setMode(m)}
                style={{
                  ...S.modeBtn,
                  background: mode === m ? "#1a1a2e" : "#f0f0f0",
                  color: mode === m ? "#fff" : "#555",
                }}
              >
                {m}
              </button>
            ))}
          </div>

          {/* 팀수 입력 */}
          {mode === "복부제" ? (
            <div style={{ display: "flex", gap: "10px", marginBottom: "14px" }}>
              <div style={{ flex: 1 }}>
                <label style={S.label}>1부 팀수</label>
                <input
                  type="number" value={shift1Size} min={1}
                  onChange={(e) => setShift1Size(Number(e.target.value))}
                  style={S.numInput}
                />
              </div>
              <div style={{ flex: 1 }}>
                <label style={S.label}>2부 팀수</label>
                <input
                  type="number" value={shift2Size} min={1}
                  onChange={(e) => setShift2Size(Number(e.target.value))}
                  style={S.numInput}
                />
              </div>
            </div>
          ) : (
            <div style={{ marginBottom: "14px" }}>
              <label style={S.label}>팀수</label>
              <input
                type="number" value={singleSize} min={1}
                onChange={(e) => setSingleSize(Number(e.target.value))}
                style={S.numInput}
              />
            </div>
          )}

          <label style={S.label}>이름 목록 (한 줄에 한 명, 순번 순서)</label>
          <textarea
            value={nameText}
            onChange={(e) => setNameText(e.target.value)}
            placeholder={"홍길동\n김철수\n이영희"}
            rows={8}
            style={S.textarea}
          />
          <button onClick={confirmNames} style={S.primaryBtn}>다음 →</button>
        </div>
      )}

      {/* ─── 배정 단계 ─── */}
      {view === "assign" && (
        <>
          <div style={S.card}>
            <div style={S.infoRow}>
              <span style={S.chip}>{mode}</span>
              {mode === "복부제"
                ? <span style={S.chip}>1부 {shift1Size}팀 · 2부 {shift2Size}팀</span>
                : <span style={S.chip}>{singleSize}팀</span>
              }
              <span style={S.chip}>총 {names.length}명</span>
            </div>

            {names.map((name, idx) => {
              const s = statuses[name] ?? null;
              return (
                <div key={name} style={S.personRow}>
                  <span style={S.personNum}>{idx + 1}</span>
                  <span style={S.personName}>{name}</span>
                  <div style={S.btnGroup}>
                    {STATUS_BUTTONS.filter((btn) => {
                      // 단부제에서 조출/후출은 숨김
                      if (mode === "단부제" && (btn === "조출" || btn === "후출")) return false;
                      return true;
                    }).map((btn) => {
                      const active = s === btn;
                      const disabled = btn === "찾근" && !canChakgeun(name) && s !== "찾근";
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
              <DayResultView result={dayResult} mode={mode} />
            </div>
          )}

          {/* 주간 결과 */}
          {weekly.length > 0 && (
            <div style={S.card}>
              <div style={S.sectionTitle}>📅 주간 근무표</div>
              {weekly.map(({ day, result: r }) => (
                <div key={day} style={S.weekDay}>
                  <div style={S.dayChip}>{day}</div>
                  <div style={{ flex: 1 }}>
                    <DayResultView result={r} mode={mode} compact />
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

// ── 결과 표시 컴포넌트 ─────────────────────────────
function DayResultView({ result, mode, compact = false }: {
  result: DayResult;
  mode: Mode;
  compact?: boolean;
}) {
  const cats =
    mode === "복부제"
      ? [
          { names: result.twoRound, label: "투라운드", badge: { bg: "#e0f7fa", color: "#00838f" } },
          { names: result.shift1,   label: "1부",      badge: { bg: "#e3f2fd", color: "#1565c0" } },
          { names: result.spare1,   label: "1부 스페어", badge: { bg: "#fff3e0", color: "#e65100" } },
          { names: result.shift2,   label: "2부",      badge: { bg: "#e8f5e9", color: "#2e7d32" } },
          { names: result.spare2,   label: "2부 스페어", badge: { bg: "#f3e5f5", color: "#6a1b9a" } },
          { names: result.excluded, label: "제외",     badge: { bg: "#f5f5f5", color: "#9e9e9e" } },
        ]
      : [
          { names: result.twoRound, label: "투라운드", badge: { bg: "#e0f7fa", color: "#00838f" } },
          { names: result.shift1,   label: "단부",     badge: { bg: "#e3f2fd", color: "#1565c0" } },
          { names: result.spare2,   label: "스페어",   badge: { bg: "#f3e5f5", color: "#6a1b9a" } },
          { names: result.excluded, label: "제외",     badge: { bg: "#f5f5f5", color: "#9e9e9e" } },
        ];

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: compact ? "4px" : "10px" }}>
      {cats.map(({ names, label, badge }) => {
        if (!names.length) return null;
        return (
          <div key={label} style={{ display: "flex", alignItems: "flex-start", gap: "8px" }}>
            <span style={{
              display: "inline-block",
              padding: compact ? "1px 7px" : "3px 10px",
              borderRadius: "12px",
              fontSize: compact ? "0.68rem" : "0.75rem",
              fontWeight: 700,
              background: badge.bg,
              color: badge.color,
              flexShrink: 0,
              marginTop: "1px",
              whiteSpace: "nowrap",
            }}>
              {label}
            </span>
            <span style={{
              fontSize: compact ? "0.8rem" : "0.88rem",
              color: "#333",
              lineHeight: 1.6,
              flex: 1,
            }}>
              {names.join("  ·  ")}
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
  modeBtn: {
    flex: 1,
    padding: "10px",
    border: "none",
    borderRadius: "8px",
    fontSize: "0.9rem",
    fontWeight: 700,
    cursor: "pointer",
  },
  infoRow: {
    display: "flex",
    gap: "6px",
    marginBottom: "12px",
    flexWrap: "wrap",
  },
  chip: {
    padding: "3px 10px",
    borderRadius: "20px",
    background: "#f0f0f0",
    color: "#555",
    fontSize: "0.78rem",
    fontWeight: 600,
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
};
