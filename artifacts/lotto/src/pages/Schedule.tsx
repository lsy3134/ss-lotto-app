import { useState, useEffect, useMemo } from "react";
import { useLocation } from "wouter";
import * as XLSX from "xlsx";
import { ROSTER, isAutoOff, type GroupType, type PersonData } from "../data/roster";

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
  // 다음날 예상 순번: [스페어(앞번호순)] → [찾근자] → [오늘 근무자]
  nextDayQueue?: string[];
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

  // ── 다음날 예상 순번 계산 ──────────────────────────
  // 규정: [오늘 스페어 앞번호순] → [오늘 찾근자] → [오늘 근무자 순서]
  // 스페어를 선 앞번호(= 낮은 큐번호)가 다음날 첫 대기가 됨
  const nextDayQueue = buildNextDayQueue(names, spare1, spare2, twoRound, excluded);

  return { twoRound, shift1, spare1, shift2, spare2, excluded, 조출List, 후출List, nextDayQueue };
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

  const nextDayQueue = buildNextDayQueue(names, [], spare2, twoRound, excluded);

  return { twoRound, shift1, spare1: [], shift2: [], spare2, excluded, nextDayQueue };
}

// ── 다음날 예상 순번 계산기 ───────────────────────────
// 규정: 오늘 스페어(앞번호 우선) → 찾근자 → 오늘 근무자 순서
// "뒷번호가 스페어를 스면 앞번호는 죽어" = 앞번호(낮은 큐)가 다음날 먼저
function buildNextDayQueue(
  allNames: string[],     // 전체 이름 (원래 큐 순서)
  spare1: string[],
  spare2: string[],
  twoRound: string[],
  excluded: string[]
): string[] {
  const spareSet   = new Set([...spare1, ...spare2]);
  const twoRndSet  = new Set(twoRound);
  const exclSet    = new Set(excluded);

  // ① 스페어를 선 사람: 1부스페어 → 2부스페어 (앞번호 = 낮은 큐 인덱스가 먼저)
  const spares = [...spare1, ...spare2];

  // ② 찾근자 (원래 큐 순서 유지)
  const twoRndOrdered = allNames.filter(n => twoRndSet.has(n));

  // ③ 오늘 실제 근무자 (스페어·찾근·제외 아닌 나머지, 원래 큐 순서)
  const workers = allNames.filter(
    n => !spareSet.has(n) && !twoRndSet.has(n) && !exclSet.has(n)
  );

  return [...spares, ...twoRndOrdered, ...workers];
}



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

  // 상태 선택 모달 (전체 순번표 표시)
  const [modalStatus, setModalStatus] = useState<StatusType | null>(null);
  const [modalSearch, setModalSearch] = useState("");

  // 명단 보기 모달 (해당 상태인 사람만 표시)
  const [viewStatusModal, setViewStatusModal] = useState<"당번" | "휴무" | "병가" | null>(null);

  // ── 사용자 정의 순번표 (localStorage 영구 저장) ──
  const [customRoster, setCustomRoster] = useState<PersonData[]>(() => {
    try {
      const saved = localStorage.getItem("lotto_customRoster");
      return saved ? (JSON.parse(saved) as PersonData[]) : ROSTER;
    } catch { return ROSTER; }
  });
  // 저장
  useEffect(() => {
    localStorage.setItem("lotto_customRoster", JSON.stringify(customRoster));
  }, [customRoster]);
  // 조 순 정렬
  const sortedCustomRoster = useMemo(() =>
    [...customRoster].sort((a, b) => a.조 !== b.조 ? a.조 - b.조 : a.no - b.no),
    [customRoster]
  );
  // 이름 → PersonData 맵
  const customRosterMap = useMemo(() =>
    Object.fromEntries(customRoster.map(p => [p.name, p])),
    [customRoster]
  );

  // ── 순번표 편집 모달 ──
  const [rosterEditorOpen, setRosterEditorOpen] = useState(false);
  const [rosterEditorSearch, setRosterEditorSearch] = useState("");
  const [rosterForm, setRosterForm] = useState<{ mode: "add"|"edit"; orig?: PersonData; name: string; 조: 1|2|3|4; group: GroupType } | null>(null);

  // 현재 요일의 유효 상태 반환
  function effectiveStatus(name: string, dayIdx: number = dayOfWeek): StatusType {
    if (name in manualStatuses) return manualStatuses[name];
    const person = customRosterMap[name];
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
    const loaded = sortedCustomRoster.map((p) => p.name);
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

  // ── 순번표 편집 함수들 ──────────────────────────
  function openAddForm() {
    setRosterForm({ mode: "add", name: "", 조: 1, group: "하우스" });
  }
  function openEditForm(p: PersonData) {
    setRosterForm({ mode: "edit", orig: p, name: p.name, 조: p.조, group: p.group });
  }
  function deletePerson(p: PersonData) {
    if (!confirm(`"${p.name}"을(를) 순번표에서 삭제할까요?`)) return;
    setCustomRoster(prev => prev.filter(x => x.name !== p.name));
  }
  function savePerson() {
    if (!rosterForm) return;
    const trimName = rosterForm.name.trim();
    if (!trimName) { alert("이름을 입력하세요."); return; }
    if (rosterForm.mode === "add") {
      // 같은 이름 중복 체크
      if (customRoster.some(x => x.name === trimName)) {
        alert("이미 같은 이름이 있습니다."); return;
      }
      // 해당 조에서 가장 큰 no + 1
      const sameJo = customRoster.filter(x => x.조 === rosterForm.조);
      const maxNo = sameJo.length > 0 ? Math.max(...sameJo.map(x => x.no)) : 0;
      const newPerson: PersonData = {
        no: maxNo + 1, name: trimName,
        조: rosterForm.조, group: rosterForm.group,
      };
      setCustomRoster(prev => [...prev, newPerson]);
    } else {
      // 수정 모드 (이름 변경 중복 체크)
      if (trimName !== rosterForm.orig?.name && customRoster.some(x => x.name === trimName)) {
        alert("이미 같은 이름이 있습니다."); return;
      }
      setCustomRoster(prev => prev.map(x =>
        x.name === rosterForm.orig?.name
          ? { ...x, name: trimName, 조: rosterForm.조, group: rosterForm.group }
          : x
      ));
    }
    setRosterForm(null);
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
                    {/* 당번·휴무 자동 표시 */}
                    {d.당번 > 0 && (
                      <span style={{
                        fontSize: "0.55rem", fontWeight: 700,
                        color: isSelected ? "#90caf9" : "#1565c0",
                        lineHeight: 1,
                      }}>
                        당{d.당번}
                      </span>
                    )}
                    {d.휴무 > 0 && (
                      <span style={{
                        fontSize: "0.55rem", fontWeight: 700,
                        color: isSelected ? "#ccc" : "#9e9e9e",
                        lineHeight: 1,
                      }}>
                        휴{d.휴무}
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
                {/* 클릭 가능한 상태 뱃지 */}
                {(["당번", "휴무", "병가"] as const).map((st) => {
                  const sc = STATUS_COLOR[st];
                  const assigned = names.filter(n => effectiveStatus(n) === st).length;
                  const canClick = names.length > 0;
                  return (
                    <div key={st}
                      onClick={() => { if (canClick) setViewStatusModal(st); }}
                      title={canClick ? `${st} 명단 보기` : "순번표를 먼저 불러오세요"}
                      style={{
                        display: "flex", flexDirection: "column", alignItems: "center",
                        background: assigned > 0 ? sc.bg + "22" : sc.bg + "12",
                        borderRadius: "8px", padding: "6px 10px",
                        border: `1.5px solid ${assigned > 0 ? sc.bg : sc.bg + "55"}`,
                        minWidth: "52px", cursor: canClick ? "pointer" : "default",
                        position: "relative",
                      }}>
                      <span style={{ fontSize: "0.65rem", color: sc.bg, fontWeight: 700 }}>{st}</span>
                      <span style={{ fontSize: "1rem", fontWeight: 700, color: sc.bg }}>
                        {assigned > 0 ? assigned : selectedDate[st as "당번"|"휴무"|"병가"]}
                      </span>
                      {canClick && (
                        <span style={{
                          position: "absolute", top: "2px", right: "3px",
                          fontSize: "0.5rem", color: sc.bg + "99",
                        }}>▼</span>
                      )}
                    </div>
                  );
                })}
              </div>

              {/* 조출·후출·찾근 빠른 배정 버튼 (순번표 로드 후 표시) */}
              {names.length > 0 && mode === "2부제" && (
                <div style={{ display: "flex", gap: "6px", marginTop: "8px", flexWrap: "wrap" }}>
                  {([
                    { st: "조출" as const, icon: "⬆", sub: "1부 앞", disabled: !cho가능, hint: !cho가능 ? "1부 6팀+" : `${cho현재수}/4명` },
                    { st: "후출" as const, icon: "⬇", sub: "2부 뒤", disabled: false, hint: `${hu현재수}/4명` },
                    { st: "찾근" as const, icon: "🔄", sub: "투라운드", disabled: !cho가능, hint: !cho가능 ? "1부 6팀+" : `${checkedCounts.찾근 ?? 0}명` },
                  ]).map(({ st, icon, sub, disabled, hint }) => {
                    const sc = STATUS_COLOR[st];
                    const assigned = names.filter(n => effectiveStatus(n) === st).length;
                    return (
                      <button key={st}
                        onClick={() => { if (!disabled) { setModalStatus(st); setModalSearch(""); } }}
                        disabled={disabled}
                        style={{
                          display: "flex", alignItems: "center", gap: "5px",
                          padding: "5px 10px", borderRadius: "20px", border: "none",
                          background: disabled ? "#f0f0f0" : (assigned > 0 ? sc.bg : sc.bg + "22"),
                          color: disabled ? "#bbb" : (assigned > 0 ? "#fff" : sc.bg),
                          fontWeight: 700, fontSize: "0.78rem",
                          cursor: disabled ? "not-allowed" : "pointer",
                          opacity: disabled ? 0.6 : 1,
                        }}>
                        <span>{icon}</span>
                        <span>{st}</span>
                        <span style={{ fontSize: "0.68rem", opacity: 0.8 }}>{hint}</span>
                      </button>
                    );
                  })}
                </div>
              )}
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

          {/* 순번표 불러오기 + 편집 */}
          <div style={{ display: "flex", gap: "8px", marginBottom: "12px" }}>
            <button onClick={loadRoster} style={{ ...S.primaryBtn, background: "#1565c0", flex: 1, marginBottom: 0 }}>
              📋 순번표 불러오기 ({customRoster.length}명)
            </button>
            <button
              onClick={() => { setRosterEditorSearch(""); setRosterForm(null); setRosterEditorOpen(true); }}
              style={{
                padding: "12px 14px", borderRadius: "12px",
                border: "1.5px solid #1565c0", background: "#fff",
                color: "#1565c0", fontWeight: 700, fontSize: "0.85rem",
                cursor: "pointer", whiteSpace: "nowrap",
              }}>
              ✏ 편집
            </button>
          </div>

        </div>
      )}

      {/* ── 순번표 편집 모달 ── */}
      {rosterEditorOpen && (
        <div style={{
          position: "fixed", inset: 0, zIndex: 300,
          background: "rgba(0,0,0,0.5)",
          display: "flex", flexDirection: "column", justifyContent: "flex-end",
        }}
          onClick={(e) => { if (e.target === e.currentTarget && !rosterForm) setRosterEditorOpen(false); }}
        >
          <div style={{
            background: "#fff", borderRadius: "18px 18px 0 0",
            maxHeight: "92vh", display: "flex", flexDirection: "column", overflow: "hidden",
          }}>
            {/* 헤더 */}
            <div style={{
              display: "flex", alignItems: "center", padding: "14px 16px",
              borderBottom: "1px solid #f0f0f0", background: "#1a1a2e",
            }}>
              <span style={{ fontWeight: 800, fontSize: "1rem", color: "#fff" }}>
                📋 순번표 편집
              </span>
              <span style={{ marginLeft: "8px", fontSize: "0.78rem", color: "#aaa" }}>
                총 {customRoster.length}명
              </span>
              <button onClick={() => { setRosterEditorOpen(false); setRosterForm(null); }}
                style={{
                  marginLeft: "auto", background: "rgba(255,255,255,0.15)", border: "none",
                  borderRadius: "50%", width: "30px", height: "30px",
                  cursor: "pointer", fontSize: "1rem", color: "#fff", fontWeight: 700,
                }}>✕</button>
            </div>

            {rosterForm ? (
              /* ── 추가/수정 폼 ── */
              <div style={{ padding: "20px 18px", overflowY: "auto" }}>
                <div style={{ fontWeight: 700, fontSize: "1rem", marginBottom: "16px", color: "#1a1a2e" }}>
                  {rosterForm.mode === "add" ? "새 직원 추가" : `"${rosterForm.orig?.name}" 수정`}
                </div>

                {/* 이름 */}
                <div style={{ marginBottom: "14px" }}>
                  <label style={{ display: "block", fontSize: "0.8rem", fontWeight: 700, color: "#555", marginBottom: "5px" }}>이름</label>
                  <input
                    value={rosterForm.name}
                    onChange={e => setRosterForm(f => f ? { ...f, name: e.target.value } : f)}
                    placeholder="이름 입력"
                    style={{
                      width: "100%", padding: "10px 12px", borderRadius: "10px",
                      border: "1.5px solid #e0e0e0", fontSize: "0.95rem",
                      outline: "none", boxSizing: "border-box",
                    }}
                  />
                </div>

                {/* 조 선택 */}
                <div style={{ marginBottom: "14px" }}>
                  <label style={{ display: "block", fontSize: "0.8rem", fontWeight: 700, color: "#555", marginBottom: "8px" }}>조</label>
                  <div style={{ display: "flex", gap: "8px" }}>
                    {([1, 2, 3, 4] as const).map(jo => {
                      const JO_C = [
                        { bg: "#fce4ec", color: "#c62828" },
                        { bg: "#e8f5e9", color: "#2e7d32" },
                        { bg: "#e3f2fd", color: "#1565c0" },
                        { bg: "#fff8e1", color: "#f57f17" },
                      ][jo - 1];
                      const active = rosterForm.조 === jo;
                      return (
                        <button key={jo}
                          onClick={() => setRosterForm(f => f ? { ...f, 조: jo } : f)}
                          style={{
                            flex: 1, padding: "10px 0", borderRadius: "10px",
                            border: `2px solid ${active ? JO_C.color : "#e0e0e0"}`,
                            background: active ? JO_C.bg : "#fff",
                            color: active ? JO_C.color : "#aaa",
                            fontWeight: 700, fontSize: "0.9rem", cursor: "pointer",
                          }}>
                          {jo}조
                        </button>
                      );
                    })}
                  </div>
                </div>

                {/* 반 선택 */}
                <div style={{ marginBottom: "20px" }}>
                  <label style={{ display: "block", fontSize: "0.8rem", fontWeight: 700, color: "#555", marginBottom: "8px" }}>반</label>
                  <div style={{ display: "flex", gap: "8px" }}>
                    {(["하우스", "주중", "주말"] as const).map(grp => {
                      const gc = GROUP_STYLE[grp];
                      const active = rosterForm.group === grp;
                      return (
                        <button key={grp}
                          onClick={() => setRosterForm(f => f ? { ...f, group: grp } : f)}
                          style={{
                            flex: 1, padding: "10px 0", borderRadius: "10px",
                            border: `2px solid ${active ? gc.color : "#e0e0e0"}`,
                            background: active ? gc.bg : "#fff",
                            color: active ? gc.color : "#aaa",
                            fontWeight: 700, fontSize: "0.85rem", cursor: "pointer",
                          }}>
                          {gc.label}
                        </button>
                      );
                    })}
                  </div>
                  <div style={{ fontSize: "0.7rem", color: "#999", marginTop: "6px" }}>
                    주중반: 토·일 자동 휴무 / 주말반: 월~목 자동 휴무 / 하우스: 항상 근무
                  </div>
                </div>

                <div style={{ display: "flex", gap: "8px" }}>
                  <button onClick={() => setRosterForm(null)}
                    style={{
                      flex: 1, padding: "13px", borderRadius: "12px",
                      border: "1.5px solid #e0e0e0", background: "#fff",
                      color: "#555", fontWeight: 700, fontSize: "0.95rem", cursor: "pointer",
                    }}>취소</button>
                  <button onClick={savePerson}
                    style={{
                      flex: 2, padding: "13px", borderRadius: "12px", border: "none",
                      background: "#1a1a2e", color: "#fff",
                      fontWeight: 700, fontSize: "0.95rem", cursor: "pointer",
                    }}>
                    {rosterForm.mode === "add" ? "추가" : "수정 완료"}
                  </button>
                </div>
              </div>
            ) : (
              <>
                {/* 검색 */}
                <div style={{ padding: "10px 14px", borderBottom: "1px solid #f0f0f0" }}>
                  <input
                    value={rosterEditorSearch}
                    onChange={e => setRosterEditorSearch(e.target.value)}
                    placeholder="🔍 이름 검색..."
                    style={{
                      width: "100%", padding: "8px 12px", borderRadius: "10px",
                      border: "1.5px solid #e0e0e0", fontSize: "0.9rem",
                      outline: "none", boxSizing: "border-box",
                    }}
                  />
                </div>

                {/* 인원 목록 */}
                <div style={{ overflowY: "auto", flex: 1, padding: "4px 0" }}>
                  {(() => {
                    const q = rosterEditorSearch.trim().toLowerCase();
                    const filtered = sortedCustomRoster.filter(p =>
                      !q || p.name.includes(q) || p.조.toString().includes(q) || p.group.includes(q)
                    );
                    const JO_COLORS: Record<number, { bg: string; color: string }> = {
                      1: { bg: "#fce4ec", color: "#c62828" },
                      2: { bg: "#e8f5e9", color: "#2e7d32" },
                      3: { bg: "#e3f2fd", color: "#1565c0" },
                      4: { bg: "#fff8e1", color: "#f57f17" },
                    };
                    let lastJo: number | null = null;
                    const items: React.ReactNode[] = [];

                    filtered.forEach(p => {
                      if (!q && p.조 !== lastJo) {
                        lastJo = p.조;
                        const jc = JO_COLORS[p.조] ?? { bg: "#f5f5f5", color: "#555" };
                        items.push(
                          <div key={`jh-${p.조}`} style={{
                            display: "flex", alignItems: "center", gap: "8px",
                            padding: "6px 14px 3px",
                          }}>
                            <span style={{
                              background: jc.bg, color: jc.color, fontWeight: 700,
                              fontSize: "0.72rem", padding: "1px 10px", borderRadius: "20px",
                              border: `1px solid ${jc.color}44`,
                            }}>{p.조}조</span>
                            <div style={{ flex: 1, height: "1px", background: jc.color + "33" }} />
                          </div>
                        );
                      }

                      const gc = GROUP_STYLE[p.group];
                      items.push(
                        <div key={p.name} style={{
                          display: "flex", alignItems: "center", gap: "8px",
                          padding: "9px 14px", borderBottom: "1px solid #f9f9f9",
                        }}>
                          <div style={{ flex: 1 }}>
                            <span style={{ fontWeight: 600, fontSize: "0.92rem" }}>{p.name}</span>
                            <span style={{
                              marginLeft: "7px", fontSize: "0.65rem",
                              color: gc.color, fontWeight: 600,
                            }}>{gc.label}</span>
                          </div>
                          <button
                            onClick={() => openEditForm(p)}
                            style={{
                              padding: "4px 10px", borderRadius: "8px",
                              border: "1.5px solid #e0e0e0", background: "#fff",
                              color: "#555", fontSize: "0.78rem", fontWeight: 600, cursor: "pointer",
                            }}>✏ 수정</button>
                          <button
                            onClick={() => deletePerson(p)}
                            style={{
                              padding: "4px 10px", borderRadius: "8px",
                              border: "1.5px solid #ffcdd2", background: "#fff9f9",
                              color: "#e53935", fontSize: "0.78rem", fontWeight: 600, cursor: "pointer",
                            }}>🗑</button>
                        </div>
                      );
                    });

                    if (items.length === 0) {
                      return <div style={{ textAlign: "center", color: "#bbb", padding: "30px" }}>검색 결과 없음</div>;
                    }
                    return items;
                  })()}
                </div>

                {/* 추가 버튼 */}
                <div style={{ padding: "10px 14px", borderTop: "1px solid #f0f0f0" }}>
                  <button onClick={openAddForm}
                    style={{
                      width: "100%", padding: "13px", borderRadius: "12px", border: "none",
                      background: "#1565c0", color: "#fff",
                      fontWeight: 700, fontSize: "0.95rem", cursor: "pointer",
                    }}>
                    + 새 직원 추가
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      )}

      {/* ── 상태 선택 모달 (전역, 어느 화면에서나 열림) ── */}
      {modalStatus && (
        <div style={{
          position: "fixed", inset: 0, zIndex: 200,
          background: "rgba(0,0,0,0.45)",
          display: "flex", flexDirection: "column", justifyContent: "flex-end",
        }}
          onClick={(e) => { if (e.target === e.currentTarget) setModalStatus(null); }}
        >
          <div style={{
            background: "#fff", borderRadius: "18px 18px 0 0",
            maxHeight: "88vh", display: "flex", flexDirection: "column",
            overflow: "hidden",
          }}>
            {/* 헤더 */}
            <div style={{
              display: "flex", alignItems: "center", padding: "14px 16px",
              borderBottom: "1px solid #f0f0f0",
              background: (STATUS_COLOR[modalStatus] ?? { bg: "#f5f5f5" }).bg,
            }}>
              <span style={{ fontWeight: 800, fontSize: "1rem", color: (STATUS_COLOR[modalStatus] ?? { color: "#333" }).color }}>
                {modalStatus} 배정
              </span>
              <span style={{ marginLeft: "8px", fontSize: "0.78rem", color: (STATUS_COLOR[modalStatus] ?? { color: "#fff" }).color + "cc" }}>
                현재 {names.filter(n => effectiveStatus(n) === modalStatus).length}명 선택됨
              </span>
              <button onClick={() => setModalStatus(null)}
                style={{
                  marginLeft: "auto", background: "rgba(255,255,255,0.25)", border: "none",
                  borderRadius: "50%", width: "30px", height: "30px",
                  cursor: "pointer", fontSize: "1rem", fontWeight: 700,
                  color: (STATUS_COLOR[modalStatus] ?? { color: "#333" }).color,
                }}>✕</button>
            </div>

            {/* 검색창 */}
            <div style={{ padding: "10px 14px", borderBottom: "1px solid #f0f0f0" }}>
              <input
                value={modalSearch}
                onChange={(e) => setModalSearch(e.target.value)}
                placeholder="🔍 이름 검색..."
                autoFocus
                style={{
                  width: "100%", padding: "8px 12px", borderRadius: "10px",
                  border: "1.5px solid #e0e0e0", fontSize: "0.9rem",
                  outline: "none", boxSizing: "border-box",
                }}
              />
            </div>

            {/* 선택된 사람 칩 */}
            {names.filter(n => effectiveStatus(n) === modalStatus).length > 0 && (
              <div style={{ padding: "8px 14px 6px", display: "flex", flexWrap: "wrap", gap: "5px", borderBottom: "1px solid #f0f0f0" }}>
                {names.filter(n => effectiveStatus(n) === modalStatus).map(n => (
                  <span key={n}
                    onClick={() => toggleStatus(n, modalStatus)}
                    style={{
                      display: "inline-flex", alignItems: "center", gap: "4px",
                      padding: "4px 10px", borderRadius: "20px",
                      background: (STATUS_COLOR[modalStatus] ?? { bg: "#eee" }).bg,
                      color: (STATUS_COLOR[modalStatus] ?? { color: "#333" }).color,
                      fontWeight: 700, fontSize: "0.8rem", cursor: "pointer",
                      border: `1.5px solid ${(STATUS_COLOR[modalStatus] ?? { color: "#888" }).color}55`,
                    }}>
                    {n} <span style={{ fontSize: "0.7rem" }}>✕</span>
                  </span>
                ))}
              </div>
            )}

            {/* 순번표 목록 */}
            <div style={{ overflowY: "auto", flex: 1, padding: "6px 0 16px" }}>
              {(() => {
                const query = modalSearch.trim().toLowerCase();
                const filtered = names.filter(n =>
                  !query || n.toLowerCase().includes(query) ||
                  (customRosterMap[n]?.조?.toString() ?? "").includes(query)
                );
                const JO_COLORS: Record<number, { bg: string; color: string }> = {
                  1: { bg: "#fce4ec", color: "#c62828" },
                  2: { bg: "#e8f5e9", color: "#2e7d32" },
                  3: { bg: "#e3f2fd", color: "#1565c0" },
                  4: { bg: "#fff8e1", color: "#f57f17" },
                };
                let lastJo: number | null = null;
                const items: React.ReactNode[] = [];

                filtered.forEach((name) => {
                  const person = customRosterMap[name];
                  const joNum = person?.조;
                  const effS = effectiveStatus(name);
                  const isSelected = effS === modalStatus;
                  const isDifferent = effS !== null && effS !== modalStatus;

                  if (!query && rosterLoaded && joNum !== undefined && joNum !== lastJo) {
                    lastJo = joNum;
                    const jc = JO_COLORS[joNum] ?? { bg: "#f5f5f5", color: "#555" };
                    items.push(
                      <div key={`h-${joNum}`} style={{ padding: "5px 14px 3px", display: "flex", alignItems: "center", gap: "8px" }}>
                        <span style={{
                          background: jc.bg, color: jc.color, fontWeight: 700,
                          fontSize: "0.72rem", padding: "1px 10px", borderRadius: "20px",
                          border: `1px solid ${jc.color}44`,
                        }}>{joNum}조</span>
                        <div style={{ flex: 1, height: "1px", background: jc.color + "33" }} />
                      </div>
                    );
                  }

                  const isDisabled = (() => {
                    if (isSelected) return false;
                    if (modalStatus === "찾근") return !canChakgeun(name);
                    if (modalStatus === "조출") return !cho가능 || cho현재수 >= 4;
                    if (modalStatus === "후출") return hu현재수 >= 4;
                    return false;
                  })();

                  const sc = STATUS_COLOR[modalStatus] ?? { bg: "#eee", color: "#333" };
                  items.push(
                    <div key={name}
                      onClick={() => !isDisabled && toggleStatus(name, modalStatus)}
                      style={{
                        display: "flex", alignItems: "center", gap: "10px",
                        padding: "10px 14px",
                        background: isSelected ? sc.bg + "33" : "transparent",
                        borderLeft: isSelected ? `3px solid ${sc.color}` : "3px solid transparent",
                        cursor: isDisabled ? "not-allowed" : "pointer",
                        opacity: isDisabled ? 0.38 : 1,
                      }}
                    >
                      <div style={{
                        width: "22px", height: "22px", borderRadius: "6px", flexShrink: 0,
                        border: `2px solid ${isSelected ? sc.color : "#ddd"}`,
                        background: isSelected ? sc.bg : "#fff",
                        display: "flex", alignItems: "center", justifyContent: "center",
                      }}>
                        {isSelected && <span style={{ fontSize: "0.9rem", color: sc.color, fontWeight: 900 }}>✓</span>}
                      </div>
                      <div style={{ flex: 1 }}>
                        <span style={{ fontWeight: 600, fontSize: "0.92rem" }}>{name}</span>
                        {person && (
                          <span style={{ marginLeft: "6px", fontSize: "0.65rem", color: GROUP_STYLE[person.group].color, fontWeight: 600 }}>
                            {person.조}조 · {GROUP_STYLE[person.group].label}
                          </span>
                        )}
                      </div>
                      {isDifferent && (
                        <span style={{
                          padding: "2px 8px", borderRadius: "12px", fontSize: "0.7rem", fontWeight: 700,
                          background: (STATUS_COLOR[effS!] ?? { bg: "#eee" }).bg,
                          color: (STATUS_COLOR[effS!] ?? { color: "#555" }).color,
                        }}>{effS}</span>
                      )}
                      {isDisabled && !isSelected && (
                        <span style={{ fontSize: "0.65rem", color: "#bbb" }}>불가</span>
                      )}
                    </div>
                  );
                });

                if (items.length === 0) {
                  return <div style={{ textAlign: "center", color: "#bbb", padding: "30px" }}>검색 결과 없음</div>;
                }
                return items;
              })()}
            </div>

            {/* 완료 버튼 */}
            <div style={{ padding: "10px 14px", borderTop: "1px solid #f0f0f0" }}>
              <button onClick={() => setModalStatus(null)}
                style={{
                  width: "100%", padding: "13px", borderRadius: "12px", border: "none",
                  background: (STATUS_COLOR[modalStatus] ?? { bg: "#1a1a2e" }).bg,
                  color: (STATUS_COLOR[modalStatus] ?? { color: "#fff" }).color,
                  fontWeight: 700, fontSize: "0.95rem", cursor: "pointer",
                }}>
                완료 — {names.filter(n => effectiveStatus(n) === modalStatus).length}명 선택됨
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── 명단 보기 모달 (당번/휴무/병가 해당자만 표시) ── */}
      {viewStatusModal && (() => {
        const st = viewStatusModal;
        const sc = STATUS_COLOR[st] ?? { bg: "#eee", color: "#333" };
        const people = names.filter(n => effectiveStatus(n) === st);
        return (
          <div style={{
            position: "fixed", inset: 0, zIndex: 200,
            background: "rgba(0,0,0,0.45)",
            display: "flex", flexDirection: "column", justifyContent: "flex-end",
          }}
            onClick={(e) => { if (e.target === e.currentTarget) setViewStatusModal(null); }}
          >
            <div style={{
              background: "#fff", borderRadius: "18px 18px 0 0",
              maxHeight: "75vh", display: "flex", flexDirection: "column",
              overflow: "hidden",
            }}>
              {/* 헤더 */}
              <div style={{
                padding: "16px 16px 10px",
                background: sc.bg,
                display: "flex", alignItems: "center", justifyContent: "space-between",
              }}>
                <div>
                  <span style={{ fontWeight: 800, fontSize: "1rem", color: sc.color }}>{st} 명단</span>
                  <span style={{ marginLeft: "8px", fontSize: "0.8rem", color: sc.color + "cc" }}>
                    총 {people.length}명
                  </span>
                </div>
                <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
                  {/* 추가 배정 버튼 → 전체 모달로 이동 */}
                  <button
                    onClick={() => { setViewStatusModal(null); setModalStatus(st); setModalSearch(""); }}
                    style={{
                      padding: "5px 12px", borderRadius: "8px", border: "none",
                      background: sc.color + "33", color: sc.color,
                      fontWeight: 700, fontSize: "0.78rem", cursor: "pointer",
                    }}>
                    + 추가
                  </button>
                  <button onClick={() => setViewStatusModal(null)}
                    style={{
                      background: "transparent", border: "none",
                      fontSize: "1.2rem", cursor: "pointer", color: sc.color, lineHeight: 1,
                    }}>✕</button>
                </div>
              </div>

              {/* 명단 */}
              <div style={{ overflowY: "auto", flex: 1, padding: "10px 14px" }}>
                {people.length === 0 ? (
                  <div style={{ textAlign: "center", color: "#bbb", padding: "40px 0", fontSize: "0.95rem" }}>
                    현재 {st}인 사람이 없습니다
                  </div>
                ) : (
                  people.map((name) => {
                    const person = customRosterMap[name];
                    const isAuto = !manualStatuses[name];
                    return (
                      <div key={name} style={{
                        display: "flex", alignItems: "center", gap: "10px",
                        padding: "10px 12px", borderRadius: "10px", marginBottom: "6px",
                        background: sc.bg + "18",
                        border: `1px solid ${sc.bg}44`,
                      }}>
                        {/* 이름 + 조 정보 */}
                        <div style={{ flex: 1 }}>
                          <span style={{ fontWeight: 700, fontSize: "0.95rem" }}>{name}</span>
                          {person && (
                            <span style={{ marginLeft: "8px", fontSize: "0.7rem", color: GROUP_STYLE[person.group].color, fontWeight: 600 }}>
                              {person.조}조 · {GROUP_STYLE[person.group].label}
                            </span>
                          )}
                          {isAuto && (
                            <span style={{ marginLeft: "6px", fontSize: "0.62rem", color: "#9e9e9e" }}>자동</span>
                          )}
                        </div>
                        {/* 해제 버튼 (자동 휴무는 해제 불가) */}
                        {!isAuto && (
                          <button
                            onClick={() => toggleStatus(name, st)}
                            style={{
                              padding: "3px 10px", borderRadius: "8px", border: "none",
                              background: "#f5f5f5", color: "#777",
                              fontWeight: 700, fontSize: "0.75rem", cursor: "pointer",
                            }}>
                            해제
                          </button>
                        )}
                      </div>
                    );
                  })
                )}
              </div>

              {/* 하단 닫기 */}
              <div style={{ padding: "10px 14px", borderTop: "1px solid #f0f0f0" }}>
                <button onClick={() => setViewStatusModal(null)}
                  style={{
                    width: "100%", padding: "13px", borderRadius: "12px", border: "none",
                    background: sc.bg, color: sc.color,
                    fontWeight: 700, fontSize: "0.95rem", cursor: "pointer",
                  }}>
                  닫기
                </button>
              </div>
            </div>
          </div>
        );
      })()}

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

            {/* ── 빠른 상태 배정 버튼 ── */}
            {names.length > 0 && (
              <>
                {/* 조출·후출·찾근 (2부제만) */}
                {mode === "2부제" && (
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "7px", marginBottom: "8px" }}>
                    {[
                      {
                        status: "조출" as StatusType,
                        icon: "⬆", color: "#ff6b35", bg: "#fff3ee",
                        label: "조출", sub: `1부 앞배치`,
                        count: cho현재수, max: 4,
                        disabled: !cho가능,
                        hint: !cho가능 ? "1부 6팀+" : ""
                      },
                      {
                        status: "후출" as StatusType,
                        icon: "⬇", color: "#2196f3", bg: "#e8f4ff",
                        label: "후출", sub: `2부 뒤배치`,
                        count: hu현재수, max: 4, disabled: false, hint: ""
                      },
                      {
                        status: "찾근" as StatusType,
                        icon: "🔄", color: "#00838f", bg: "#e0faf9",
                        label: "찾근", sub: `1+2부 투라운드`,
                        count: checkedCounts.찾근 ?? 0, max: null,
                        disabled: !cho가능,
                        hint: !cho가능 ? "1부 6팀+" : ""
                      },
                    ].map(({ status, icon, color, bg, label, sub, count, max, disabled, hint }) => (
                      <button key={status}
                        onClick={() => { if (!disabled) { setModalStatus(status); setModalSearch(""); } }}
                        style={{
                          background: disabled ? "#f5f5f5" : bg,
                          border: `1.5px solid ${disabled ? "#e0e0e0" : color + "66"}`,
                          borderRadius: "10px", padding: "8px 6px",
                          cursor: disabled ? "not-allowed" : "pointer",
                          textAlign: "center", opacity: disabled ? 0.55 : 1,
                        }}
                        title={hint}
                      >
                        <div style={{ fontSize: "1.2rem", marginBottom: "2px" }}>{icon}</div>
                        <div style={{ fontWeight: 700, fontSize: "0.82rem", color: disabled ? "#aaa" : color }}>{label}</div>
                        <div style={{ fontSize: "0.68rem", color: "#888" }}>{sub}</div>
                        <div style={{
                          marginTop: "4px", fontWeight: 700, fontSize: "0.9rem",
                          color: count > 0 ? color : "#bbb",
                        }}>
                          {count}{max ? `/${max}` : ""}명
                        </div>
                        {hint && <div style={{ fontSize: "0.6rem", color: "#c62828", marginTop: "2px" }}>{hint}</div>}
                      </button>
                    ))}
                  </div>
                )}

                {/* 당번·휴무·병가 */}
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "7px", marginBottom: "10px" }}>
                  {[
                    {
                      status: "당번" as StatusType,
                      icon: "📋", color: "#e53935", bg: "#ffebee",
                      label: "당번", sub: "당일 당번 근무",
                      count: checkedCounts.당번 ?? 0
                    },
                    {
                      status: "휴무" as StatusType,
                      icon: "🌿", color: "#757575", bg: "#f5f5f5",
                      label: "휴무", sub: "오늘 휴무",
                      count: checkedCounts.휴무 ?? 0
                    },
                    {
                      status: "병가" as StatusType,
                      icon: "🏥", color: "#9e9e9e", bg: "#fafafa",
                      label: "병가", sub: "병가 처리",
                      count: checkedCounts.병가 ?? 0
                    },
                  ].map(({ status, icon, color, bg, label, sub, count }) => (
                    <button key={status}
                      onClick={() => { setModalStatus(status); setModalSearch(""); }}
                      style={{
                        background: bg, border: `1.5px solid ${color}44`,
                        borderRadius: "10px", padding: "8px 6px",
                        cursor: "pointer", textAlign: "center",
                      }}
                    >
                      <div style={{ fontSize: "1.2rem", marginBottom: "2px" }}>{icon}</div>
                      <div style={{ fontWeight: 700, fontSize: "0.82rem", color }}>{label}</div>
                      <div style={{ fontSize: "0.68rem", color: "#888" }}>{sub}</div>
                      <div style={{
                        marginTop: "4px", fontWeight: 700, fontSize: "0.9rem",
                        color: count > 0 ? color : "#bbb",
                      }}>
                        {count}명
                      </div>
                    </button>
                  ))}
                </div>
              </>
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
                const person = customRosterMap[name];
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

  const nextDay = result.nextDayQueue ?? [];

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

      {/* ── 다음날 예상 순번 ── */}
      {!compact && nextDay.length > 0 && (
        <NextDayQueueView
          queue={nextDay}
          spare1={result.spare1}
          spare2={result.spare2}
          twoRound={result.twoRound}
        />
      )}
    </div>
  );
}

// ── 다음날 예상 순번 뷰 ─────────────────────────────
function NextDayQueueView({
  queue, spare1, spare2, twoRound,
}: {
  queue: string[];
  spare1: string[];
  spare2: string[];
  twoRound: string[];
}) {
  const spare1Set  = new Set(spare1);
  const spare2Set  = new Set(spare2);
  const twoRndSet  = new Set(twoRound);

  // 종류별 색상
  function tagOf(name: string) {
    if (spare1Set.has(name))  return { label: "1부스페어", color: "#e65100", bg: "#fff3e0" };
    if (spare2Set.has(name))  return { label: "2부스페어", color: "#6a1b9a", bg: "#f3e5f5" };
    if (twoRndSet.has(name))  return { label: "찾근",     color: "#00838f", bg: "#e0f7fa" };
    return null;
  }

  // 스페어가 없으면 표시 생략
  if (spare1.length === 0 && spare2.length === 0) return null;

  return (
    <div style={{
      marginTop: "6px",
      background: "linear-gradient(135deg, #e8eaf6 0%, #f3e5f5 100%)",
      border: "1px solid #c5cae9",
      borderRadius: "10px",
      padding: "10px 12px",
    }}>
      <div style={{ fontWeight: 700, fontSize: "0.78rem", color: "#3f51b5", marginBottom: "8px" }}>
        📅 다음날 예상 첫 순번
        <span style={{ fontWeight: 400, color: "#888", marginLeft: "6px", fontSize: "0.72rem" }}>
          (스페어 앞번호 → 찾근 → 근무자 순)
        </span>
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: "5px" }}>
        {queue.slice(0, 20).map((name, i) => {
          const tag = tagOf(name);
          return (
            <div key={name} style={{
              display: "flex", alignItems: "center", gap: "3px",
              background: tag ? tag.bg : "#f5f5f5",
              border: `1px solid ${tag ? tag.color + "44" : "#ddd"}`,
              borderRadius: "20px",
              padding: "3px 9px",
              fontSize: "0.78rem",
            }}>
              <span style={{ color: "#888", fontSize: "0.68rem", fontWeight: 600 }}>{i + 1}.</span>
              <span style={{ fontWeight: tag ? 700 : 400, color: tag ? tag.color : "#333" }}>{name}</span>
              {tag && (
                <span style={{
                  fontSize: "0.62rem", color: tag.color, fontWeight: 700,
                  background: tag.bg, borderRadius: "8px", padding: "0 4px",
                }}>
                  {tag.label}
                </span>
              )}
            </div>
          );
        })}
        {queue.length > 20 && (
          <span style={{ fontSize: "0.75rem", color: "#888", alignSelf: "center" }}>
            +{queue.length - 20}명 더...
          </span>
        )}
      </div>
      {/* 규칙 설명 */}
      <div style={{ marginTop: "6px", fontSize: "0.7rem", color: "#666" }}>
        ▶ 뒷번호가 스페어를 서면 앞번호가 다음날 첫 대기
        {spare1.length > 0 && (
          <span style={{ marginLeft: "8px", color: "#e65100", fontWeight: 600 }}>
            오늘 1부스페어: {spare1.join(", ")}
          </span>
        )}
        {spare2.length > 0 && (
          <span style={{ marginLeft: "8px", color: "#6a1b9a", fontWeight: 600 }}>
            오늘 2부스페어: {spare2.join(", ")}
          </span>
        )}
      </div>
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
