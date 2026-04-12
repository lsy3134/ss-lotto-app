import { useState, useEffect } from "react";
import { useLocation } from "wouter";
import * as XLSX from "xlsx";
import { ROSTER, ROSTER_MAP, isAutoOff, type GroupType } from "../data/roster";

const BASE = import.meta.env.BASE_URL.replace(/\/$/, "");

// ── 타입 ──────────────────────────────────────────
type StatusType =
  | "조출" | "후출" | "찾근"
  | "당번" | "병가" | "휴무" | "하우스"
  | null;

type Mode = "2부제" | "단부제";

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
const DAY_MAP: Record<string, number> = { 월: 0, 화: 1, 수: 2, 목: 3, 금: 4, 토: 5, 일: 6 };

// ── 엑셀 날짜 데이터 타입 ──────────────────────────
interface ExcelDayData {
  dateLabel: string;  // "04.01 (수)"
  dayName: string;    // "수"
  dayIdx: number;     // 2 (수=2)
  당번: number;
  휴무: number;
  병가: number;
  가용인원: number;
  예약팀수: number;
}

interface DayResult {
  twoRound: string[];
  shift1: string[];
  spare1: string[];
  shift2: string[];
  spare2: string[];
  excluded: string[];
  // 위치 표시용 메타 (후출: 2부 뒤에서3번째, 조출: 1부 앞)
  조출List?: string[];
  후출List?: string[];
}

// ── 엑셀 파싱 훅 ──────────────────────────────────
function useExcelData() {
  const [excelDays, setExcelDays] = useState<ExcelDayData[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  async function load() {
    setLoading(true);
    setError("");
    try {
      const res = await fetch(`${BASE}/schedule.xlsx`);
      if (!res.ok) throw new Error("파일을 찾을 수 없습니다");
      const buf = await res.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });

      // 시트 찾기 (설정 시트 제외, 투입계산 시트 사용)
      const sheetName = wb.SheetNames.find((n) => n.includes("투입계산")) ?? wb.SheetNames[1];
      const ws = wb.Sheets[sheetName];
      const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1");

      const days: ExcelDayData[] = [];
      for (let r = range.s.r + 2; r <= range.e.r; r++) {
        const getVal = (c: number) => {
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          return cell ? cell.v : undefined;
        };

        const dateLabel = getVal(0);
        const dayName = getVal(1);
        if (!dateLabel || !dayName || typeof dayName !== "string" || !DAY_MAP.hasOwnProperty(dayName)) continue;

        const 당번 = Number(getVal(5)) || 0;
        const 휴무 = Number(getVal(6)) || 0;
        const 병가 = Number(getVal(7)) || 0;
        const 가용인원 = Number(getVal(9)) || 0;
        const 예약팀수 = Number(getVal(10)) || 0;

        days.push({
          dateLabel: String(dateLabel),
          dayName,
          dayIdx: DAY_MAP[dayName],
          당번,
          휴무,
          병가,
          가용인원,
          예약팀수,
        });
      }
      setExcelDays(days);
    } catch (e: any) {
      setError(e.message ?? "오류");
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => { load(); }, []);

  return { excelDays, loading, error, reload: load };
}

// ── 배정 엔진: 2부제 (PDF 규정 기준) ──────────────
// 배치 순서:
//   1부: 찾근자(선발3번~) → 조출자 → 일반순번 → [1부스페어]
//   2부: 일반순번 → 후출자(뒤에서 3번째) → 일반순번(막팀) → [2부스페어]
function assignDouble(
  names: string[],
  statuses: Record<string, StatusType>,
  shift1Size: number,
  shift2Size: number
): DayResult {
  const twoRound: string[] = [];   // 찾근 (1부+2부 투라운드)
  const 조출List: string[] = [];   // 조출 (1부 앞 고정, 최대 4명)
  const 후출List: string[] = [];   // 후출 (2부 뒤에서 3번째, 최대 4명)
  const excluded: string[] = [];
  const autoQueue: string[] = [];  // 일반 순번 대기열

  for (const name of names) {
    const s = statuses[name] ?? null;
    if (s === "찾근")  { twoRound.push(name); }
    else if (s === "조출") {
      if (조출List.length < 4) 조출List.push(name); else autoQueue.push(name);
    } else if (s === "후출") {
      if (후출List.length < 4) 후출List.push(name); else autoQueue.push(name);
    } else if (EXCLUDED_SET.has(s ?? "")) { excluded.push(name); }
    else { autoQueue.push(name); }
  }

  // ── 1부 배치: 찾근 → 조출 → 일반순번 ──
  const fixed1 = [...twoRound, ...조출List];           // 고정 1부 인원
  const avail1 = Math.max(0, shift1Size - fixed1.length);
  const shift1 = [...fixed1, ...autoQueue.slice(0, avail1)];
  const spare1 = autoQueue.slice(avail1, avail1 + 1); // 1부스페어 1명
  const remaining = autoQueue.slice(avail1 + 1);       // 2부 후보 대기열

  // ── 2부 배치: 일반순번 + 후출자 뒤에서 3번째 위치 ──
  // 찾근자는 투라운드이므로 2부에도 참여 (twoRound 표시로 커버)
  // 후출자: 2부 팀수 기준 뒤에서 3번째 고정
  // 예) 2부 35팀 → 33번째 위치에 후출자 삽입
  const avail2 = Math.max(0, shift2Size - 후출List.length);
  const normalFor2 = remaining.slice(0, avail2);
  const restQueue = remaining.slice(avail2);

  let shift2: string[];
  if (후출List.length > 0 && normalFor2.length >= 2) {
    // 뒤에서 3번째 = 끝에서 2개 자리를 막팀(1번째, 2번째)으로 남기고 후출자 삽입
    const insertAt = Math.max(0, normalFor2.length - 2);
    shift2 = [
      ...normalFor2.slice(0, insertAt),
      ...후출List,
      ...normalFor2.slice(insertAt),
    ];
  } else {
    // 일반순번이 2명 미만이면 그냥 뒤에 추가
    shift2 = [...normalFor2, ...후출List];
  }

  const spare2 = restQueue;

  return { twoRound, shift1, spare1, shift2, spare2, excluded, 조출List, 후출List };
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
    if (person && !(name in manualStatuses) && isAutoOff(person.group, dayIdx)) {
      result[name] = "휴무";
    }
  });
  return result;
}

// ── 조 순서로 정렬된 순번표 ────────────────────────
const SORTED_ROSTER = [...ROSTER].sort((a, b) => {
  if (a.조 !== b.조) return a.조 - b.조;
  return a.no - b.no;
});

// ── 메인 컴포넌트 ─────────────────────────────────
export default function SchedulePage() {
  const [, setLocation] = useLocation();
  const { excelDays, loading: xlLoading, error: xlError } = useExcelData();

  // 설정
  const [mode, setMode] = useState<Mode>("2부제");
  // 2부제: totalSize = 총팀수, shift1Size = 1부팀수, shift2Size = 총팀수 - 1부팀수
  const [totalSize, setTotalSize] = useState(70);
  const [shift1Size, setShift1Size] = useState(35);
  const shift2Size = Math.max(0, totalSize - shift1Size);
  const [singleSize, setSingleSize] = useState(60);
  const [nameText, setNameText] = useState("");
  const [dayOfWeek, setDayOfWeek] = useState(0);

  // 선택된 날짜 (엑셀 날짜)
  const [selectedDate, setSelectedDate] = useState<ExcelDayData | null>(null);

  // 인원
  const [names, setNames] = useState<string[]>([]);
  const [rosterLoaded, setRosterLoaded] = useState(false);

  // 수동 상태
  const [manualStatuses, setManualStatuses] = useState<Record<string, StatusType>>({});

  // 결과
  const [dayResult, setDayResult] = useState<DayResult | null>(null);
  const [weekly, setWeekly] = useState<{ day: string; result: DayResult }[]>([]);
  const [view, setView] = useState<"input" | "assign">("input");

  // 현재 요일의 유효 상태 반환
  function effectiveStatus(name: string, dayIdx: number = dayOfWeek): StatusType {
    if (name in manualStatuses) return manualStatuses[name];
    const person = ROSTER_MAP[name];
    if (person && isAutoOff(person.group, dayIdx)) return "휴무";
    return null;
  }

  // 상태 토글
  function toggleStatus(name: string, btn: StatusType) {
    setManualStatuses((prev) => {
      const cur = effectiveStatus(name);
      if (cur === btn && name in prev) {
        const next = { ...prev };
        delete next[name];
        return next;
      } else if (cur === btn && !(name in prev)) {
        return { ...prev, [name]: null };
      } else {
        return { ...prev, [name]: btn };
      }
    });
  }

  // 날짜 선택 → 팀수/요일 자동 설정
  function selectExcelDate(day: ExcelDayData) {
    setSelectedDate(day);
    setDayOfWeek(day.dayIdx);
    if (day.예약팀수 > 0) {
      const total = day.예약팀수;
      const half = Math.round(total / 2);
      setTotalSize(total);
      setShift1Size(half);
      setSingleSize(total);
    }
  }

  // 순번표 불러오기 (1조→2조→3조→4조 순서)
  function loadRoster() {
    const loaded = SORTED_ROSTER.map((p) => p.name);
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
    const result = mode === "2부제"
      ? assignDouble(names, statuses, shift1Size, shift2Size)
      : assignSingle(names, statuses, singleSize);
    setDayResult(result);
    setWeekly([]);
  }

  function generateWeek() {
    const results = DAY_LABELS.map((day, di) => {
      const statuses = getEffective(di);
      const result = mode === "2부제"
        ? assignDouble(names, statuses, shift1Size, shift2Size)
        : assignSingle(names, statuses, singleSize);
      return { day, result };
    });
    setWeekly(results);
    setDayResult(null);
  }

  // 활성 인원 대기열(제외·찾근 제외)에서의 순번 인덱스
  function activeQueueIndex(name: string): number {
    const active = names.filter((n) => {
      const s = effectiveStatus(n);
      return s !== "찾근" && !EXCLUDED_SET.has(s ?? "");
    });
    return active.indexOf(name);
  }

  // 찾근 가능 여부 (PDF 규정 기준)
  // 2부제: 1부팀수 ≥ 6 + 본인 순번이 1부스페어 이후(2부 배정 순번~) 또는 근무 안 될 때
  // 단부제: 스페어 이후(근무 안 될 때)만 가능
  function canChakgeun(name: string): boolean {
    const s = effectiveStatus(name);
    if (s === "찾근") return true; // 이미 찾근 상태 → 해제 허용
    if (mode === "단부제") {
      // 단부제: 근무 안 될 때(스페어 이후)만 가능
      const qi = activeQueueIndex(name);
      return qi > singleSize; // singleSize 이후 = 근무 안 됨
    }
    // 2부제: 1부 6팀 이상 필수
    if (shift1Size < 6) return false;
    // 본인 순번이 1부스페어(shift1Size번째) 이후여야 함
    // = 2부 배정 순번 or 근무 안 될 때
    const qi = activeQueueIndex(name);
    return qi >= shift1Size; // shift1Size 인덱스 = 1부스페어 위치 이후
  }

  // 조출 가능 여부 (1부 6팀 이상 필수)
  const cho가능 = shift1Size >= 6;
  const cho현재수 = names.filter((n) => effectiveStatus(n) === "조출").length;
  const hu현재수 = names.filter((n) => effectiveStatus(n) === "후출").length;

  // 순번 위치에 따른 배정 구간 레이블
  function getSlotLabel(name: string): { label: string; color: string } | null {
    if (mode !== "2부제") return null;
    const s = effectiveStatus(name);
    if (s === "찾근") return { label: "투라운드", color: "#00bcd4" };
    if (s === "조출") return { label: "1부↑", color: "#ff6b35" };
    if (s === "후출") return { label: "2부↑", color: "#2196f3" };
    if (EXCLUDED_SET.has(s ?? "")) return null;
    const qi = activeQueueIndex(name);
    if (qi < 0) return null;
    if (qi < shift1Size)          return { label: `1부 #${qi + 1}`, color: "#1565c0" };
    if (qi === shift1Size)        return { label: "1부스페어", color: "#e65100" };
    if (qi <= shift1Size + shift2Size) return { label: `2부 #${qi - shift1Size}`, color: "#2e7d32" };
    return { label: "2부스페어", color: "#6a1b9a" };
  }

  // 현재 선택된 날짜에 대한 체크 카운트
  const checkedCounts = {
    찾근: names.filter((n) => effectiveStatus(n) === "찾근").length,
    조출: names.filter((n) => effectiveStatus(n) === "조출").length,
    후출: names.filter((n) => effectiveStatus(n) === "후출").length,
    당번: names.filter((n) => effectiveStatus(n) === "당번").length,
    병가: names.filter((n) => effectiveStatus(n) === "병가").length,
    휴무: names.filter((n) => effectiveStatus(n) === "휴무").length,
  };

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
            {(["2부제", "단부제"] as Mode[]).map((m) => (
              <button key={m} onClick={() => setMode(m)}
                style={{ ...S.modeBtn, background: mode === m ? "#1a1a2e" : "#f0f0f0", color: mode === m ? "#fff" : "#555" }}>
                {m}
              </button>
            ))}
          </div>

          {/* ── 엑셀 날짜 선택 ── */}
          <label style={S.label}>
            📊 날짜 선택 (엑셀 자동 로드)
            {xlLoading && <span style={{ color: "#aaa", fontWeight: 400, marginLeft: "6px" }}>불러오는 중…</span>}
            {xlError && <span style={{ color: "#e53935", fontWeight: 400, marginLeft: "6px" }}>{xlError}</span>}
          </label>

          {excelDays.length > 0 && (
            <div style={S.dateGrid}>
              {excelDays.map((d) => {
                const isSelected = selectedDate?.dateLabel === d.dateLabel;
                const isWeekend = d.dayIdx === 5 || d.dayIdx === 6;
                const hasTeams = d.예약팀수 > 0;
                return (
                  <button
                    key={d.dateLabel}
                    onClick={() => selectExcelDate(d)}
                    style={{
                      ...S.dateBtn,
                      background: isSelected ? "#1a1a2e" : "#f8f9fa",
                      color: isSelected ? "#fff" : isWeekend ? "#c62828" : "#333",
                      border: isSelected ? "2px solid #1a1a2e" : hasTeams ? "2px solid #1565c0" : "1px solid #e0e0e0",
                    }}
                  >
                    <span style={{ fontSize: "0.72rem", fontWeight: 700 }}>{d.dateLabel.split(" ")[0]}</span>
                    <span style={{ fontSize: "0.65rem", opacity: 0.7 }}>{d.dayName}</span>
                    {hasTeams && (
                      <span style={{
                        fontSize: "0.6rem",
                        background: isSelected ? "rgba(255,255,255,0.25)" : "#e3f2fd",
                        color: isSelected ? "#fff" : "#1565c0",
                        borderRadius: "4px",
                        padding: "1px 4px",
                        fontWeight: 700,
                      }}>
                        {d.예약팀수}팀
                      </span>
                    )}
                  </button>
                );
              })}
            </div>
          )}

          {/* 선택된 날짜 정보 */}
          {selectedDate && (
            <div style={S.excelInfo}>
              <div style={S.excelInfoTitle}>
                📋 {selectedDate.dateLabel} 기준 데이터
                {selectedDate.예약팀수 > 0 && (
                  <span style={{ marginLeft: "8px", color: "#1565c0", fontWeight: 700 }}>
                    → 팀수 자동 입력 완료
                  </span>
                )}
              </div>
              <div style={S.excelStatRow}>
                <StatBadge label="가용인원" value={selectedDate.가용인원} color="#1565c0" />
                <StatBadge label="예약팀수" value={selectedDate.예약팀수 || "미입력"} color={selectedDate.예약팀수 > 0 ? "#2e7d32" : "#9e9e9e"} />
                <StatBadge label="당번" value={selectedDate.당번} color="#e53935" />
                <StatBadge label="휴무" value={selectedDate.휴무} color="#757575" />
                <StatBadge label="병가" value={selectedDate.병가} color="#9e9e9e" />
              </div>
              {selectedDate.예약팀수 === 0 && (
                <div style={{ fontSize: "0.72rem", color: "#ff8f00", marginTop: "4px" }}>
                  ⚠ 예약팀수 미입력 — 아래에서 직접 팀수를 입력해 주세요
                </div>
              )}
            </div>
          )}

          {/* 팀수 입력 */}
          {mode === "2부제" ? (
            <div style={{ marginBottom: "14px", marginTop: "14px" }}>
              <div style={{ display: "flex", gap: "10px", marginBottom: "8px" }}>
                <div style={{ flex: 1 }}>
                  <label style={S.label}>총 팀수</label>
                  <input type="number" value={totalSize} min={1}
                    onChange={(e) => setTotalSize(Number(e.target.value))} style={S.numInput} />
                </div>
                <div style={{ flex: 1 }}>
                  <label style={S.label}>1부 팀수</label>
                  <input type="number" value={shift1Size} min={1} max={totalSize}
                    onChange={(e) => setShift1Size(Number(e.target.value))} style={S.numInput} />
                </div>
              </div>
              {/* 자동 계산 결과 표시 */}
              <div style={S.calcBox}>
                <span style={{ color: "#1565c0" }}>1부 {shift1Size}팀</span>
                <span style={{ color: "#aaa" }}>+</span>
                <span style={{ color: "#2e7d32" }}>2부 {shift2Size}팀</span>
                <span style={{ color: "#aaa" }}>=</span>
                <span style={{ fontWeight: 700 }}>총 {totalSize}팀</span>
                <span style={{ color: "#aaa", margin: "0 4px" }}>│</span>
                <span style={{ color: "#e65100", fontSize: "0.75rem" }}>1부스페어: {shift1Size + 1}번째</span>
                <span style={{ color: "#6a1b9a", fontSize: "0.75rem" }}>2부스페어: {totalSize + 2}번째~</span>
              </div>
            </div>
          ) : (
            <div style={{ marginBottom: "14px", marginTop: "14px" }}>
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
            <div style={{ display: "flex", gap: "6px", marginBottom: "10px", flexWrap: "wrap" }}>
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

            {/* 선택 날짜 표시 */}
            {selectedDate && (
              <div style={S.dateBar}>
                <span style={{ fontWeight: 700, color: "#1a1a2e" }}>{selectedDate.dateLabel}</span>
                <span style={{ color: "#666", fontSize: "0.8rem", marginLeft: "6px" }}>
                  예약 {selectedDate.예약팀수}팀
                </span>
                <div style={{ marginLeft: "auto", display: "flex", gap: "6px" }}>
                  <StatBadge label="당번" value={selectedDate.당번} color="#e53935" small />
                  <StatBadge label="휴무" value={selectedDate.휴무} color="#757575" small />
                  <StatBadge label="병가" value={selectedDate.병가} color="#9e9e9e" small />
                </div>
              </div>
            )}

            {/* 조출·후출·찾근 규정 안내 */}
            {mode === "2부제" && (
              <div style={{
                background: "#f8f9ff", border: "1px solid #e0e4ff", borderRadius: "10px",
                padding: "10px 12px", marginBottom: "10px", fontSize: "0.75rem", color: "#444",
              }}>
                <div style={{ fontWeight: 700, color: "#3f51b5", marginBottom: "6px" }}>
                  📘 조출·후출·찾근 규정 (PDF 기준)
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "6px" }}>
                  <div style={{ background: "#fff3ee", borderRadius: "7px", padding: "5px 7px", borderLeft: "3px solid #ff6b35" }}>
                    <div style={{ fontWeight: 700, color: "#ff6b35" }}>조출</div>
                    <div>1부 배치 (앞)</div>
                    <div>최대 4명</div>
                    <div style={{ color: shift1Size >= 6 ? "#2e7d32" : "#c62828", fontWeight: 600 }}>
                      {shift1Size >= 6 ? `✓ 사용 가능` : `✗ 1부 6팀+`}
                    </div>
                  </div>
                  <div style={{ background: "#e8f4ff", borderRadius: "7px", padding: "5px 7px", borderLeft: "3px solid #2196f3" }}>
                    <div style={{ fontWeight: 700, color: "#2196f3" }}>후출</div>
                    <div>2부 배치 (뒤)</div>
                    <div>최대 4명</div>
                    <div style={{ color: "#2e7d32", fontWeight: 600 }}>뒤에서 3번째</div>
                  </div>
                  <div style={{ background: "#e0faf9", borderRadius: "7px", padding: "5px 7px", borderLeft: "3px solid #00bcd4" }}>
                    <div style={{ fontWeight: 700, color: "#00838f" }}>찾근</div>
                    <div>1부+2부 투라운드</div>
                    <div>조출자보다 앞</div>
                    <div style={{ color: shift1Size >= 6 ? "#2e7d32" : "#c62828", fontWeight: 600 }}>
                      {shift1Size >= 6 ? `✓ 2부순번~` : `✗ 1부 6팀+`}
                    </div>
                  </div>
                </div>
              </div>
            )}

            {/* 현재 체크 현황 */}
            {names.length > 0 && (
              <div style={{ display: "flex", gap: "5px", marginBottom: "10px", flexWrap: "wrap" }}>
                {checkedCounts.찾근 > 0 && <MiniCount label="찾근" count={checkedCounts.찾근} color="#00bcd4" />}
                {checkedCounts.조출 > 0 && <MiniCount label={`조출 (${cho현재수}/4)`} count={checkedCounts.조출} color="#ff6b35" />}
                {checkedCounts.후출 > 0 && <MiniCount label={`후출 (${hu현재수}/4)`} count={checkedCounts.후출} color="#2196f3" />}
                {checkedCounts.당번 > 0 && <MiniCount label="당번" count={checkedCounts.당번} color="#e53935" />}
                {checkedCounts.병가 > 0 && <MiniCount label="병가" count={checkedCounts.병가} color="#9e9e9e" />}
                {checkedCounts.휴무 > 0 && <MiniCount label="휴무" count={checkedCounts.휴무} color="#bdbdbd" />}
              </div>
            )}

            {/* 설정 정보 */}
            <div style={S.infoRow}>
              <span style={S.chip}>{mode}</span>
              {mode === "2부제"
                ? <>
                    <span style={S.chip}>1부 {shift1Size}팀</span>
                    <span style={S.chip}>2부 {shift2Size}팀</span>
                    <span style={{ ...S.chip, fontWeight: 700 }}>총 {totalSize}팀</span>
                  </>
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

            {/* 인원 리스트 (조별 구분) */}
            {(() => {
              const rows: React.ReactNode[] = [];
              let lastGroup: number | null = null;
              const JO_COLORS: Record<number, { bg: string; color: string }> = {
                1: { bg: "#fce4ec", color: "#c62828" },
                2: { bg: "#e8f5e9", color: "#2e7d32" },
                3: { bg: "#e3f2fd", color: "#1565c0" },
                4: { bg: "#fff8e1", color: "#f57f17" },
              };
              names.forEach((name, idx) => {
                const person = ROSTER_MAP[name];
                const joNum = person?.조;
                // 조 구분 헤더
                if (rosterLoaded && joNum !== undefined && joNum !== lastGroup) {
                  lastGroup = joNum;
                  const jc = JO_COLORS[joNum] ?? { bg: "#f5f5f5", color: "#555" };
                  rows.push(
                    <div key={`jo-${joNum}`} style={{
                      display: "flex", alignItems: "center", gap: "8px",
                      padding: "6px 0 4px", marginTop: idx === 0 ? "0" : "4px",
                    }}>
                      <div style={{
                        background: jc.bg, color: jc.color,
                        fontWeight: 700, fontSize: "0.78rem",
                        padding: "2px 12px", borderRadius: "20px",
                        border: `1px solid ${jc.color}44`,
                      }}>
                        {joNum}조
                      </div>
                      <div style={{ flex: 1, height: "1px", background: jc.color + "33" }} />
                    </div>
                  );
                }

                const effS = effectiveStatus(name, dayOfWeek);
                const isAutoHumu = !(name in manualStatuses) && effS === "휴무";
                const slotLabel = getSlotLabel(name);

                rows.push(
                  <div key={name} style={{
                    ...S.personRow,
                    opacity: effS === "휴무" ? 0.5 : 1,
                  }}>
                    <span style={S.personNum}>{person?.no ?? idx + 1}</span>

                    <div style={{ minWidth: "72px" }}>
                      <div style={S.personName}>{name}</div>
                      <div style={{ display: "flex", gap: "3px", flexWrap: "wrap" }}>
                        {person && (
                          <span style={{
                            fontSize: "0.6rem", padding: "1px 4px", borderRadius: "4px",
                            background: GROUP_STYLE[person.group].bg,
                            color: GROUP_STYLE[person.group].color, fontWeight: 700,
                          }}>
                            {GROUP_STYLE[person.group].label}{isAutoHumu ? "(자동)" : ""}
                          </span>
                        )}
                        {slotLabel && (
                          <span style={{
                            fontSize: "0.6rem", padding: "1px 4px", borderRadius: "4px",
                            background: slotLabel.color + "18",
                            color: slotLabel.color, fontWeight: 700,
                          }}>
                            {slotLabel.label}
                          </span>
                        )}
                      </div>
                    </div>

                    <div style={S.btnGroup}>
                      {STATUS_BUTTONS.filter((btn) =>
                        mode !== "단부제" || (btn !== "조출" && btn !== "후출")
                      ).map((btn) => {
                        const active = effS === btn;
                        const isAutoActive = active && isAutoHumu;
                        // 조출: 1부 6팀 미만이면 비활성화 / 찾근: canChakgeun 조건
                        const disabled =
                          (btn === "조출" && !cho가능 && effS !== "조출") ||
                          (btn === "찾근" && !canChakgeun(name));
                        const col = active ? STATUS_COLOR[btn!] : null;
                        // 조출/후출 최대 4명 초과 시 비활성화
                        const maxReached =
                          (btn === "조출" && cho현재수 >= 4 && effS !== "조출") ||
                          (btn === "후출" && hu현재수 >= 4 && effS !== "후출");
                        return (
                          <button key={btn} disabled={disabled || maxReached}
                            onClick={() => toggleStatus(name, btn)}
                            title={
                              btn === "조출" && !cho가능 ? "1부 6팀 이상일 때만 사용 가능" :
                              btn === "조출" && maxReached ? "조출 최대 4명" :
                              btn === "후출" && maxReached ? "후출 최대 4명" :
                              btn === "찾근" && !canChakgeun(name) ? "2부 배정 순번 이상이어야 사용 가능" : ""
                            }
                            style={{
                              ...S.statusBtn,
                              background: active ? col!.bg : "#f0f0f0",
                              color: active ? col!.color : (disabled || maxReached) ? "#ccc" : "#555",
                              border: isAutoActive ? `2px dashed ${col!.bg}` : active ? "none" : "1px solid #e0e0e0",
                              opacity: (disabled || maxReached) ? 0.35 : 1,
                            }}>
                            {btn}
                          </button>
                        );
                      })}
                    </div>
                  </div>
                );
              });
              return rows;
            })()}

            <div style={{ display: "flex", gap: "8px", marginTop: "16px" }}>
              <button onClick={assign} style={{ ...S.primaryBtn, flex: 1 }}>배정하기</button>
              <button onClick={generateWeek} style={{ ...S.primaryBtn, flex: 1, background: "#374151" }}>
                일주일 생성
              </button>
            </div>
          </div>

          {/* 2부제 배정 기준 안내 */}
          {mode === "2부제" && names.length > 0 && (
            <div style={S.cutoffBox}>
              <span style={{ fontWeight: 700, fontSize: "0.78rem", color: "#555", marginRight: "6px" }}>배정 기준</span>
              <span style={{ color: "#1565c0", fontSize: "0.75rem" }}>1~{shift1Size}번 → 1부</span>
              <span style={{ color: "#e65100", fontSize: "0.75rem" }}>{shift1Size + 1}번 → 1부스페어</span>
              <span style={{ color: "#2e7d32", fontSize: "0.75rem" }}>{shift1Size + 2}~{totalSize + 1}번 → 2부</span>
              <span style={{ color: "#6a1b9a", fontSize: "0.75rem" }}>{totalSize + 2}번~ → 2부스페어</span>
            </div>
          )}

          {/* 1일 결과 */}
          {dayResult && weekly.length === 0 && (
            <div style={S.card}>
              <div style={S.sectionTitle}>
                📋 {selectedDate ? selectedDate.dateLabel : DAY_LABELS[dayOfWeek] + "요일"} 배정 결과
              </div>
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

// ── 소형 카운트 배지 ──────────────────────────────
function MiniCount({ label, count, color }: { label: string; count: number; color: string }) {
  return (
    <span style={{
      padding: "2px 8px", borderRadius: "20px", fontSize: "0.72rem",
      background: color + "22", color, fontWeight: 700, border: `1px solid ${color}44`,
    }}>
      {label} {count}
    </span>
  );
}

// ── 통계 배지 ─────────────────────────────────────
function StatBadge({ label, value, color, small = false }: {
  label: string; value: number | string; color: string; small?: boolean;
}) {
  return (
    <div style={{
      display: "flex", flexDirection: "column", alignItems: "center",
      background: color + "15", borderRadius: "8px",
      padding: small ? "3px 7px" : "6px 10px",
      border: `1px solid ${color}33`,
      minWidth: small ? "44px" : "52px",
    }}>
      <span style={{ fontSize: small ? "0.6rem" : "0.65rem", color, fontWeight: 600 }}>{label}</span>
      <span style={{ fontSize: small ? "0.85rem" : "1rem", fontWeight: 700, color }}>{value}</span>
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
  const cats = mode === "2부제" ? CATS_DOUBLE : CATS_SINGLE;
  const 조출Set = new Set(result.조출List ?? []);
  const 후출Set = new Set(result.후출List ?? []);

  function renderPeople(people: string[], key: string) {
    if (compact) return people.join("  ·  ");
    // 2부 목록에서 후출자·조출자 강조 표시
    return (
      <span style={{ fontSize: "0.88rem", color: "#333", lineHeight: 1.7 }}>
        {people.map((n, i) => {
          const isCho = (key === "shift1") && 조출Set.has(n);
          const isHu = (key === "shift2") && 후출Set.has(n);
          const isTwoR = key === "twoRound";
          const suffix = isCho ? " [조출]" : isHu ? " [후출]" : "";
          return (
            <span key={n}>
              {i > 0 && <span style={{ color: "#ccc" }}> · </span>}
              <span style={{
                fontWeight: (isCho || isHu || isTwoR) ? 700 : 400,
                color: isCho ? "#ff6b35" : isHu ? "#2196f3" : isTwoR ? "#00838f" : "#333",
              }}>
                {n}{suffix}
              </span>
            </span>
          );
        })}
      </span>
    );
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: compact ? "4px" : "10px" }}>
      {cats.map(({ key, label, badge }) => {
        const people = result[key];
        if (!people.length) return null;
        // 2부 항목에 후출 위치 안내 추가
        const posNote = !compact && mode === "2부제" && key === "shift2" && (result.후출List?.length ?? 0) > 0
          ? ` (후출: 뒤에서 3번째)` : "";
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
            }}>{label}{posNote}</span>
            {compact
              ? <span style={{ fontSize: "0.8rem", color: "#333", lineHeight: 1.6, flex: 1 }}>
                  {renderPeople(people, key)}
                </span>
              : <div style={{ flex: 1 }}>{renderPeople(people, key)}</div>
            }
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
  calcBox: {
    display: "flex", gap: "8px", flexWrap: "wrap", alignItems: "center",
    background: "#f8f9fa", borderRadius: "8px", padding: "8px 12px",
    fontSize: "0.82rem", fontWeight: 600,
  },
  cutoffBox: {
    display: "flex", gap: "8px", flexWrap: "wrap", alignItems: "center",
    background: "#fafafa", borderRadius: "8px", padding: "8px 12px",
    margin: "0 0 8px", border: "1px solid #e0e0e0",
  },
  dateGrid: {
    display: "grid", gridTemplateColumns: "repeat(5, 1fr)",
    gap: "6px", marginBottom: "14px",
  },
  dateBtn: {
    display: "flex", flexDirection: "column", alignItems: "center", gap: "2px",
    padding: "7px 4px", borderRadius: "8px", cursor: "pointer",
    fontSize: "0.75rem", fontWeight: 600, border: "1px solid #e0e0e0",
  },
  excelInfo: {
    background: "#f0f7ff", borderRadius: "10px", padding: "10px 12px",
    marginBottom: "14px", border: "1px solid #bbdefb",
  },
  excelInfoTitle: { fontSize: "0.78rem", fontWeight: 700, color: "#1565c0", marginBottom: "8px" },
  excelStatRow: { display: "flex", gap: "6px", flexWrap: "wrap" },
  dateBar: {
    display: "flex", alignItems: "center", gap: "6px",
    background: "#f8f9fa", borderRadius: "8px", padding: "8px 12px",
    marginBottom: "10px", flexWrap: "wrap",
  },
};
