import { useState } from "react";
import { useLocation } from "wouter";
import { ROSTER, ROSTER_MAP, isAutoOff, type GroupType } from "../data/roster";

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

const GROUP_STYLE: Record<GroupType, { bg: string; color: string; label: string }> = {
  하우스: { bg: "#e8f5e9", color: "#2e7d32", label: "하우스" },
  주중:   { bg: "#e3f2fd", color: "#1565c0", label: "주중" },
  주말:   { bg: "#fce4ec", color: "#c62828", label: "주말" },
};

const DAY_LABELS = ["월", "화", "수", "목", "금", "토", "일"];

interface DayResult {
  twoRound: string[];
  shift1: string[];
  spare1: string[];
  shift2: string[];
  spare2: string[];
  excluded: string[];
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
    if (s === "찾근") { twoRound.push(name); }
    else if (s === "조출") {
      if (choCount < 4) { shift1.push(name); choCount++; } else autoQueue.push(name);
    } else if (s === "후출") {
      if (huCount < 4) { shift2.push(name); huCount++; } else autoQueue.push(name);
    } else if (EXCLUDED_SET.has(s ?? "")) { excluded.push(name); }
    else { autoQueue.push(name); }
  }

  const avail1 = Math.max(0, shift1Size - choCount - twoRound.length);
  const avail2 = Math.max(0, shift2Size - huCount - twoRound.length);

  autoQueue.forEach((n, i) => {
    if (i < avail1)                      shift1.push(n);
    else if (i === avail1)               spare1.push(n);
    else if (i <= avail1 + avail2)       shift2.push(n);
    else                                 spare2.push(n);
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
    if (s === "찾근") { twoRound.push(name); }
    else if (EXCLUDED_SET.has(s ?? "")) { excluded.push(name); }
    else { autoQueue.push(name); }
  }

  const avail = Math.max(0, teamSize - twoRound.length);
  autoQueue.forEach((n, i) => {
    if (i < avail) shift1.push(n); else spare2.push(n);
  });

  return { twoRound, shift1, spare1: [], shift2: [], spare2, excluded };
}

// ── 자동 휴무 적용 ──────────────────────────────────
function applyAutoOff(
  names: string[],
  manualStatuses: Record<string, StatusType>,
  dayIdx: number
): Record<string, StatusType> {
  const result: Record<string, StatusType> = { ...manualStatuses };
  names.forEach((name) => {
    const person = ROSTER_MAP[name];
    // 수동 설정이 없는 경우에만 자동 휴무 적용
    if (person && !(name in manualStatuses) && isAutoOff(person.group, dayIdx)) {
      result[name] = "휴무";
    }
  });
  return result;
}

// ── 메인 컴포넌트 ─────────────────────────────────
export default function SchedulePage() {
  const [, setLocation] = useLocation();

  // 설정
  const [mode, setMode] = useState<Mode>("복부제");
  const [shift1Size, setShift1Size] = useState(35);
  const [shift2Size, setShift2Size] = useState(35);
  const [singleSize, setSingleSize] = useState(60);
  const [nameText, setNameText] = useState("");
  const [dayOfWeek, setDayOfWeek] = useState(0); // 0=월

  // 인원
  const [names, setNames] = useState<string[]>([]);
  const [rosterLoaded, setRosterLoaded] = useState(false);

  // 수동 상태 (사용자가 직접 클릭한 것)
  const [manualStatuses, setManualStatuses] = useState<Record<string, StatusType>>({});

  // 결과
  const [dayResult, setDayResult] = useState<DayResult | null>(null);
  const [weekly, setWeekly] = useState<{ day: string; result: DayResult }[]>([]);
  const [view, setView] = useState<"input" | "assign">("input");

  // 현재 요일의 유효 상태 반환 (auto 포함)
  function effectiveStatus(name: string, dayIdx: number = dayOfWeek): StatusType {
    if (name in manualStatuses) return manualStatuses[name];
    const person = ROSTER_MAP[name];
    if (person && isAutoOff(person.group, dayIdx)) return "휴무";
    return null;
  }

  // 상태 토글 (수동 오버라이드)
  function toggleStatus(name: string, btn: StatusType) {
    setManualStatuses((prev) => {
      const cur = effectiveStatus(name);
      if (cur === btn && name in prev) {
        // 수동 설정 제거 → auto로 복귀
        const next = { ...prev };
        delete next[name];
        return next;
      } else if (cur === btn && !(name in prev)) {
        // auto 상태를 클릭 → 명시적으로 null 설정 (auto 해제)
        return { ...prev, [name]: null };
      } else {
        return { ...prev, [name]: btn };
      }
    });
  }

  // 순번표 불러오기
  function loadRoster() {
    const loaded = ROSTER.map((p) => p.name);
    setNames(loaded);
    setRosterLoaded(true);
    setManualStatuses({});
    setDayResult(null);
    setWeekly([]);
    setView("assign");
  }

  // 직접 입력으로 다음 단계
  function confirmNames() {
    const parsed = nameText.split("\n").map((n) => n.trim()).filter(Boolean);
    if (!parsed.length) return;
    setNames(parsed);
    setRosterLoaded(false);
    setManualStatuses({});
    setDayResult(null);
    setWeekly([]);
    setView("assign");
  }

  function getEffective(dayIdx: number = dayOfWeek) {
    const base: Record<string, StatusType> = {};
    names.forEach((n) => { base[n] = effectiveStatus(n, dayIdx); });
    return base;
  }

  function assign() {
    const statuses = getEffective(dayOfWeek);
    const result = mode === "복부제"
      ? assignDouble(names, statuses, shift1Size, shift2Size)
      : assignSingle(names, statuses, singleSize);
    setDayResult(result);
    setWeekly([]);
  }

  function generateWeek() {
    const results = DAY_LABELS.map((day, di) => {
      const statuses = getEffective(di); // 각 요일별 auto-휴무 적용
      const result = mode === "복부제"
        ? assignDouble(names, statuses, shift1Size, shift2Size)
        : assignSingle(names, statuses, singleSize);
      return { day, result };
    });
    setWeekly(results);
    setDayResult(null);
  }

  function canChakgeun(name: string): boolean {
    if (mode === "단부제") return false;
    const active = names.filter((n) => !EXCLUDED_SET.has(effectiveStatus(n) ?? ""));
    return active.indexOf(name) >= shift1Size;
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
              <button key={m} onClick={() => setMode(m)}
                style={{ ...S.modeBtn, background: mode === m ? "#1a1a2e" : "#f0f0f0", color: mode === m ? "#fff" : "#555" }}>
                {m}
              </button>
            ))}
          </div>

          {/* 팀수 */}
          {mode === "복부제" ? (
            <div style={{ display: "flex", gap: "10px", marginBottom: "14px" }}>
              <div style={{ flex: 1 }}>
                <label style={S.label}>1부 팀수</label>
                <input type="number" value={shift1Size} min={1}
                  onChange={(e) => setShift1Size(Number(e.target.value))} style={S.numInput} />
              </div>
              <div style={{ flex: 1 }}>
                <label style={S.label}>2부 팀수</label>
                <input type="number" value={shift2Size} min={1}
                  onChange={(e) => setShift2Size(Number(e.target.value))} style={S.numInput} />
              </div>
            </div>
          ) : (
            <div style={{ marginBottom: "14px" }}>
              <label style={S.label}>팀수</label>
              <input type="number" value={singleSize} min={1}
                onChange={(e) => setSingleSize(Number(e.target.value))} style={S.numInput} />
            </div>
          )}

          {/* 순번표 불러오기 */}
          <button onClick={loadRoster} style={{ ...S.primaryBtn, background: "#1565c0", marginBottom: "12px" }}>
            📋 순번표 불러오기 (63명)
          </button>

          <div style={{ textAlign: "center", color: "#aaa", fontSize: "0.8rem", marginBottom: "12px" }}>또는 직접 입력</div>

          <label style={S.label}>이름 목록 (한 줄에 한 명)</label>
          <textarea value={nameText} onChange={(e) => setNameText(e.target.value)}
            placeholder={"홍길동\n김철수"} rows={5} style={S.textarea} />
          <button onClick={confirmNames} style={S.primaryBtn}>다음 →</button>
        </div>
      )}

      {/* ─── 배정 단계 ─── */}
      {view === "assign" && (
        <>
          <div style={S.card}>
            {/* 요일 선택 */}
            <label style={{ ...S.label, marginBottom: "8px" }}>오늘 요일</label>
            <div style={{ display: "flex", gap: "6px", marginBottom: "14px", flexWrap: "wrap" }}>
              {DAY_LABELS.map((d, i) => (
                <button key={d} onClick={() => setDayOfWeek(i)}
                  style={{
                    ...S.dayBtn,
                    background: dayOfWeek === i ? "#1a1a2e" : "#f0f0f0",
                    color: dayOfWeek === i ? "#fff" : "#555",
                  }}>
                  {d}
                </button>
              ))}
            </div>

            {/* 설정 정보 */}
            <div style={S.infoRow}>
              <span style={S.chip}>{mode}</span>
              {mode === "복부제"
                ? <span style={S.chip}>1부 {shift1Size}팀 · 2부 {shift2Size}팀</span>
                : <span style={S.chip}>{singleSize}팀</span>}
              <span style={S.chip}>총 {names.length}명</span>
              {rosterLoaded && (
                <>
                  <span style={{ ...S.chip, ...GROUP_STYLE["하우스"] }}>하우스 52</span>
                  <span style={{ ...S.chip, ...GROUP_STYLE["주중"] }}>주중 2</span>
                  <span style={{ ...S.chip, ...GROUP_STYLE["주말"] }}>주말 9</span>
                </>
              )}
            </div>

            {/* 인원 리스트 */}
            {names.map((name, idx) => {
              const person = ROSTER_MAP[name];
              const manualS = manualStatuses[name];
              const effS = effectiveStatus(name, dayOfWeek);
              const isAutoH휴무 = !(name in manualStatuses) && effS === "휴무";

              return (
                <div key={name} style={{
                  ...S.personRow,
                  opacity: effS === "휴무" ? 0.55 : 1,
                }}>
                  <span style={S.personNum}>{idx + 1}</span>

                  {/* 이름 + 그룹 배지 */}
                  <div style={{ minWidth: "70px" }}>
                    <div style={S.personName}>{name}</div>
                    {person && (
                      <span style={{
                        fontSize: "0.62rem",
                        padding: "1px 5px",
                        borderRadius: "4px",
                        background: GROUP_STYLE[person.group].bg,
                        color: GROUP_STYLE[person.group].color,
                        fontWeight: 700,
                      }}>
                        {GROUP_STYLE[person.group].label}
                        {isAutoH휴무 ? " (자동)" : ""}
                      </span>
                    )}
                  </div>

                  {/* 상태 버튼 */}
                  <div style={S.btnGroup}>
                    {STATUS_BUTTONS.filter((btn) =>
                      mode !== "단부제" || (btn !== "조출" && btn !== "후출")
                    ).map((btn) => {
                      const active = effS === btn;
                      const isAutoActive = active && isAutoH휴무;
                      const disabled = btn === "찾근" && !canChakgeun(name) && effS !== "찾근";
                      const col = active ? STATUS_COLOR[btn!] : null;
                      return (
                        <button key={btn} disabled={disabled}
                          onClick={() => toggleStatus(name, btn)}
                          style={{
                            ...S.statusBtn,
                            background: active ? col!.bg : "#f0f0f0",
                            color: active ? col!.color : disabled ? "#ccc" : "#555",
                            border: isAutoActive ? `2px dashed ${col!.bg}` : active ? "none" : "1px solid #e0e0e0",
                            opacity: disabled ? 0.4 : 1,
                          }}>
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
              <div style={S.sectionTitle}>📋 {DAY_LABELS[dayOfWeek]}요일 배정 결과</div>
              <DayResultView result={dayResult} mode={mode} />
            </div>
          )}

          {/* 주간 결과 */}
          {weekly.length > 0 && (
            <div style={S.card}>
              <div style={S.sectionTitle}>📅 주간 근무표 (월~일)</div>
              {weekly.map(({ day, result: r }, di) => (
                <div key={day} style={S.weekDay}>
                  <div style={{
                    ...S.dayChip,
                    background: di === 5 || di === 6 ? "#c62828" : "#1a1a2e",
                  }}>
                    {day}
                  </div>
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
const CATS_DOUBLE = [
  { key: "twoRound" as const, label: "투라운드", badge: { bg: "#e0f7fa", color: "#00838f" } },
  { key: "shift1"   as const, label: "1부",      badge: { bg: "#e3f2fd", color: "#1565c0" } },
  { key: "spare1"   as const, label: "1부 스페어", badge: { bg: "#fff3e0", color: "#e65100" } },
  { key: "shift2"   as const, label: "2부",      badge: { bg: "#e8f5e9", color: "#2e7d32" } },
  { key: "spare2"   as const, label: "2부 스페어", badge: { bg: "#f3e5f5", color: "#6a1b9a" } },
  { key: "excluded" as const, label: "휴무/제외", badge: { bg: "#f5f5f5", color: "#9e9e9e" } },
];
const CATS_SINGLE = [
  { key: "twoRound" as const, label: "투라운드", badge: { bg: "#e0f7fa", color: "#00838f" } },
  { key: "shift1"   as const, label: "단부",     badge: { bg: "#e3f2fd", color: "#1565c0" } },
  { key: "spare2"   as const, label: "스페어",   badge: { bg: "#f3e5f5", color: "#6a1b9a" } },
  { key: "excluded" as const, label: "휴무/제외", badge: { bg: "#f5f5f5", color: "#9e9e9e" } },
];

function DayResultView({ result, mode, compact = false }: {
  result: DayResult; mode: Mode; compact?: boolean;
}) {
  const cats = mode === "복부제" ? CATS_DOUBLE : CATS_SINGLE;
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: compact ? "4px" : "10px" }}>
      {cats.map(({ key, label, badge }) => {
        const people = result[key];
        if (!people.length) return null;
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
            }}>{label}</span>
            <span style={{
              fontSize: compact ? "0.8rem" : "0.88rem",
              color: "#333",
              lineHeight: 1.6,
              flex: 1,
            }}>{people.join("  ·  ")}</span>
          </div>
        );
      })}
    </div>
  );
}

// ── 스타일 ────────────────────────────────────────
const S: Record<string, React.CSSProperties> = {
  page: { fontFamily: "sans-serif", background: "#f4f7f9", minHeight: "100vh", paddingBottom: "40px" },
  header: {
    display: "flex", alignItems: "center", gap: "10px",
    padding: "14px 16px", background: "#1a1a2e", color: "white",
    position: "sticky", top: 0, zIndex: 10,
  },
  backBtn: {
    background: "rgba(255,255,255,0.15)", border: "none", color: "white",
    borderRadius: "8px", padding: "6px 12px", cursor: "pointer", fontSize: "1rem",
  },
  smallBtn: {
    marginLeft: "auto", background: "rgba(255,255,255,0.15)", border: "none", color: "white",
    borderRadius: "8px", padding: "6px 12px", cursor: "pointer", fontSize: "0.8rem",
  },
  headerTitle: { fontWeight: 700, fontSize: "1rem" },
  card: {
    background: "white", borderRadius: "14px", padding: "16px",
    margin: "12px", boxShadow: "0 2px 8px rgba(0,0,0,0.06)",
  },
  label: {
    display: "block", fontSize: "0.75rem", color: "#888",
    marginBottom: "6px", fontWeight: 600,
    textTransform: "uppercase", letterSpacing: "0.04em",
  },
  modeBtn: { flex: 1, padding: "10px", border: "none", borderRadius: "8px", fontSize: "0.9rem", fontWeight: 700, cursor: "pointer" },
  dayBtn: { padding: "6px 12px", border: "none", borderRadius: "8px", fontSize: "0.85rem", fontWeight: 700, cursor: "pointer", minWidth: "36px" },
  infoRow: { display: "flex", gap: "6px", marginBottom: "12px", flexWrap: "wrap" },
  chip: { padding: "3px 10px", borderRadius: "20px", background: "#f0f0f0", color: "#555", fontSize: "0.78rem", fontWeight: 600 },
  textarea: {
    width: "100%", padding: "10px", borderRadius: "8px", border: "1px solid #e0e0e0",
    fontSize: "0.95rem", resize: "vertical", fontFamily: "sans-serif",
    marginBottom: "14px", boxSizing: "border-box",
  },
  numInput: { width: "100%", padding: "10px", borderRadius: "8px", border: "1px solid #e0e0e0", fontSize: "1rem", boxSizing: "border-box" },
  primaryBtn: {
    width: "100%", padding: "14px", background: "#1a1a2e", color: "white",
    border: "none", borderRadius: "10px", fontSize: "1rem", fontWeight: 700, cursor: "pointer",
  },
  personRow: {
    display: "flex", alignItems: "flex-start", gap: "8px",
    padding: "10px 0", borderBottom: "1px solid #f5f5f5", flexWrap: "wrap",
    transition: "opacity 0.2s",
  },
  personNum: { minWidth: "22px", fontSize: "0.78rem", color: "#bbb", paddingTop: "4px" },
  personName: { fontWeight: 600, fontSize: "0.9rem", marginBottom: "2px" },
  btnGroup: { display: "flex", flexWrap: "wrap", gap: "4px", flex: 1 },
  statusBtn: { padding: "4px 8px", borderRadius: "6px", fontSize: "0.78rem", cursor: "pointer", fontWeight: 600, transition: "all 0.15s" },
  sectionTitle: { fontWeight: 700, fontSize: "0.95rem", marginBottom: "14px", color: "#333" },
  weekDay: { display: "flex", gap: "10px", padding: "12px 0", borderBottom: "1px solid #f0f0f0", alignItems: "flex-start" },
  dayChip: {
    minWidth: "28px", height: "28px", borderRadius: "8px",
    color: "white", display: "flex", alignItems: "center", justifyContent: "center",
    fontWeight: 700, fontSize: "0.85rem", flexShrink: 0, marginTop: "1px",
  },
};
