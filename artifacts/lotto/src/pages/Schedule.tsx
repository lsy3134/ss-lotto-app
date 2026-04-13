import { useState, useEffect, useMemo, useRef } from "react";
import { useLocation } from "wouter";
import * as XLSX from "xlsx";
import { ROSTER, isAutoOff, type GroupType, type PersonData } from "../data/roster";
import { createWorker } from "tesseract.js";

const BASE = import.meta.env.BASE_URL.replace(/\/$/, "");

// ── 타입 ──────────────────────────────────────────
type StatusType =
  | "조출" | "후출" | "찾근" | "대기"
  | "당번" | "병가" | "휴무" | "하우스"
  | null;

type Mode = "2부제" | "단부제";
type DaegeunType = "1부" | "2부" | "투라운드";

const STATUS_BUTTONS: StatusType[] = [
  "대기", "조출", "후출", "찾근", "당번", "병가", "휴무", "하우스",
];

const EXCLUDED_SET = new Set(["당번", "병가", "휴무", "하우스"]);

const STATUS_COLOR: Record<string, { bg: string; color: string }> = {
  대기:   { bg: "#fce7f3", color: "#9d174d" },  // 분홍 (1부 출근대기 = spare1)
  조출:   { bg: "#fed7aa", color: "#9a3412" },  // 주황 (조출)
  후출:   { bg: "#ddd6fe", color: "#5b21b6" },  // 연보라 (속수대기)
  찾근:   { bg: "#cffafe", color: "#164e63" },  // 하늘 (투라운드)
  당번:   { bg: "#fecaca", color: "#991b1b" },  // 연빨 (주의)
  병가:   { bg: "#e5e7eb", color: "#4b5563" },  // 회색 (비활성)
  휴무:   { bg: "#f3f4f6", color: "#6b7280" },  // 연회색 (비활성)
  하우스: { bg: "#fef08a", color: "#713f12" },  // 금 (하우스)
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

// ── 월 달력 생성 헬퍼 (엑셀 없이도 달력 표시) ──────
const KR_DAY = ["월", "화", "수", "목", "금", "토", "일"];
function generateMonthDays(mm: string, year: number): ExcelDayData[] {
  const m = parseInt(mm, 10);
  const daysInMonth = new Date(year, m, 0).getDate();
  const result: ExcelDayData[] = [];
  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(year, m - 1, d);
    const dayIdx = (date.getDay() + 6) % 7; // Mon=0 … Sun=6
    const dayName = KR_DAY[dayIdx];
    const dd = String(d).padStart(2, "0");
    result.push({
      dateLabel: `${mm}.${dd} (${dayName})`,
      dayName,
      dayIdx,
      당번: 0, 휴무: 0, 병가: 0,
      가용인원: 0, 예약팀수: 0,
    });
  }
  return result;
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
  // 다음날 예상 순번: [스페어(앞번호순)] → [오늘 근무자 + 찾근자(일반 순번으로 포함)]
  nextDayQueue?: string[];
}

// ── 엑셀 ArrayBuffer → ExcelDayData[] 공통 파서 ──
function parseExcelBuffer(buf: ArrayBuffer): ExcelDayData[] {
  const wb = XLSX.read(buf, { type: "array" });
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
    days.push({
      dateLabel: String(dateLabel),
      dayName,
      dayIdx: DAY_MAP[dayName],
      당번:    Number(getVal(5)) || 0,
      휴무:    Number(getVal(6)) || 0,
      병가:    Number(getVal(7)) || 0,
      가용인원: Number(getVal(9)) || 0,
      예약팀수: Number(getVal(10)) || 0,
    });
  }
  return days;
}

// ── 휴무 엑셀 파서 (날짜 → 이름 리스트 구조) ─────────
// 포맷A: 첫 열이 날짜, 해당 행에 이름들 (행 기반)
// 포맷B: 첫 행이 날짜, 해당 열에 이름들 (열 기반)
// 유연하게 두 포맷 모두 시도 후 더 많은 날짜를 인식한 것 채택
// 날짜 값 → "MM.DD" 키 (공통 유틸)
function toHolidayDateKey(v: unknown, month?: number): string | null {
  if (v === undefined || v === null || v === "") return null;
  if (v instanceof Date) {
    return `${String(v.getMonth() + 1).padStart(2, "0")}.${String(v.getDate()).padStart(2, "0")}`;
  }
  if (typeof v === "number") {
    if (v >= 1 && v <= 31 && month) {
      // 일 번호 단독 (월 context 있을 때만)
      return `${String(month).padStart(2, "0")}.${String(v).padStart(2, "0")}`;
    }
    if (v > 35000 && v < 60000) {
      const d = new Date(Math.round((v - 25569) * 86400 * 1000));
      return `${String(d.getUTCMonth() + 1).padStart(2, "0")}.${String(d.getUTCDate()).padStart(2, "0")}`;
    }
  }
  const s = String(v).trim();
  // "N (요일)" 또는 "N(요일)" — 달력 일 번호 형식
  const dayNumMatch = s.match(/^(\d{1,2})\s*[\(（][일월화수목금토][\)）]?$/);
  if (dayNumMatch && month) {
    return `${String(month).padStart(2, "0")}.${dayNumMatch[1].padStart(2, "0")}`;
  }
  // MM.DD / MM/DD / MM-DD 등
  let m = s.match(/^(\d{1,2})[.\-\/](\d{1,2})/);
  if (m) return `${m[1].padStart(2, "0")}.${m[2].padStart(2, "0")}`;
  // YYYY-MM-DD / YYYY.MM.DD
  m = s.match(/^\d{4}[-.](\d{1,2})[-.](\d{1,2})/);
  if (m) return `${m[1].padStart(2, "0")}.${m[2].padStart(2, "0")}`;
  // "4월 13일", "4월13일"
  m = s.match(/(\d{1,2})월\s*(\d{1,2})일/);
  if (m) return `${m[1].padStart(2, "0")}.${m[2].padStart(2, "0")}`;
  // "13일" 단독 — 월 context 있을 때
  m = s.match(/^(\d{1,2})일$/);
  if (m && month) return `${String(month).padStart(2, "0")}.${m[1].padStart(2, "0")}`;
  return null;
}

// 이름 여부: 한글 2~6자 (공백 허용)
function isKoreanName(v: unknown): boolean {
  if (!v) return false;
  const s = String(v).trim();
  if (s.length < 2 || s.length > 7) return false;
  if (/^[\uAC00-\uD7A3]{2,6}$/.test(s)) return true;
  if (/^[\uAC00-\uD7A3 ]{2,7}$/.test(s) && /[\uAC00-\uD7A3]{2,}/.test(s)) return true;
  return false;
}

// 셀에서 이름 목록 추출 — 쉼표/공백 구분 포함
function extractNames(v: unknown): string[] {
  if (!v) return [];
  const s = String(v).trim();
  if (!s) return [];
  // 쉼표/줄바꿈/슬래시로 분리
  const parts = s.split(/[,，、\/\n\r]+/).map(p => p.trim()).filter(p => isKoreanName(p));
  if (parts.length > 0) return parts;
  // 단일 이름
  if (isKoreanName(s)) return [s];
  return [];
}

// 체크 마크 여부: "O", "o", "○", "●", "V", "v", "✓", "✔", "휴", "휴무", "1", true
function isCheckMark(v: unknown): boolean {
  if (v === true || v === 1) return true;
  if (!v) return false;
  return /^[oOvV○●✓✔1]$|^휴무?$/.test(String(v).trim());
}

function parseHolidayExcelBuffer(buf: ArrayBuffer, contextMonth?: number): { map: Record<string, string[]>; debug: string } {
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const debugLines: string[] = [`시트 수: ${wb.SheetNames.length} (${wb.SheetNames.join(", ")})`];
  const nowMonth = contextMonth ?? (new Date().getMonth() + 1);

  const allResults: { map: Record<string, string[]>; score: number; desc: string }[] = [];

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    const ref = ws["!ref"];
    if (!ref) continue;
    const range = XLSX.utils.decode_range(ref);
    const rows = range.e.r - range.s.r + 1;
    const cols = range.e.c - range.s.c + 1;
    debugLines.push(`[${sheetName}] ${rows}행 × ${cols}열`);

    const getCell = (r: number, c: number) => {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      return cell ? cell.v : undefined;
    };

    const dk = (v: unknown) => toHolidayDateKey(v, nowMonth);

    // ── 전체 셀 스캔: 날짜/이름 수집 ──
    let totalDates = 0, totalNames = 0;
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const v = getCell(r, c);
        if (dk(v)) totalDates++;
        if (isKoreanName(v) || extractNames(v).length > 0) totalNames++;
      }
    }
    debugLines.push(`  날짜셀: ${totalDates}, 이름셀: ${totalNames}`);

    // 샘플 셀 값 (A1~E3 범위)
    const samples: string[] = [];
    for (let r = range.s.r; r <= Math.min(range.s.r + 2, range.e.r); r++) {
      for (let c = range.s.c; c <= Math.min(range.s.c + 4, range.e.c); c++) {
        const v = getCell(r, c);
        if (v !== undefined && v !== null && v !== "") samples.push(String(v).slice(0, 12));
      }
    }
    debugLines.push(`  샘플: ${samples.join(" | ")}`);

    // ── 포맷A: 첫 열=날짜, 나머지=이름(들) ──
    const fmtA: Record<string, string[]> = {}; let aS = 0;
    for (let r = range.s.r; r <= range.e.r; r++) {
      const key = dk(getCell(r, range.s.c));
      if (!key) continue;
      const names: string[] = [];
      for (let c = range.s.c + 1; c <= range.e.c; c++) {
        names.push(...extractNames(getCell(r, c)));
      }
      if (names.length) { fmtA[key] = [...(fmtA[key] ?? []), ...names]; aS += names.length; }
    }

    // ── 포맷B: 첫 행=날짜, 나머지=이름(들) ──
    const fmtB: Record<string, string[]> = {}; let bS = 0;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const key = dk(getCell(range.s.r, c));
      if (!key) continue;
      const names: string[] = [];
      for (let r = range.s.r + 1; r <= range.e.r; r++) {
        names.push(...extractNames(getCell(r, c)));
      }
      if (names.length) { fmtB[key] = [...(fmtB[key] ?? []), ...names]; bS += names.length; }
    }

    // ── 포맷C: 첫 열=이름, 나머지=날짜 ──
    const fmtC: Record<string, string[]> = {}; let cS = 0;
    for (let r = range.s.r; r <= range.e.r; r++) {
      const nms = extractNames(getCell(r, range.s.c));
      if (!nms.length) continue;
      for (let c = range.s.c + 1; c <= range.e.c; c++) {
        const key = dk(getCell(r, c));
        if (!key) continue;
        if (!fmtC[key]) fmtC[key] = [];
        for (const name of nms) if (!fmtC[key].includes(name)) { fmtC[key].push(name); cS++; }
      }
    }

    // ── 포맷D: 첫 행=이름, 나머지=날짜 ──
    const fmtD: Record<string, string[]> = {}; let dS = 0;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const nms = extractNames(getCell(range.s.r, c));
      if (!nms.length) continue;
      for (let r = range.s.r + 1; r <= range.e.r; r++) {
        const key = dk(getCell(r, c));
        if (!key) continue;
        if (!fmtD[key]) fmtD[key] = [];
        for (const name of nms) if (!fmtD[key].includes(name)) { fmtD[key].push(name); dS++; }
      }
    }

    // ── 포맷E: 첫 열=이름, 첫 행=날짜, 본문=체크마크 ──
    const fmtE: Record<string, string[]> = {}; let eS = 0;
    for (let r = range.s.r + 1; r <= range.e.r; r++) {
      const nms = extractNames(getCell(r, range.s.c));
      if (!nms.length) continue;
      for (let c = range.s.c + 1; c <= range.e.c; c++) {
        const key = dk(getCell(range.s.r, c));
        if (!key || !isCheckMark(getCell(r, c))) continue;
        if (!fmtE[key]) fmtE[key] = [];
        for (const name of nms) if (!fmtE[key].includes(name)) { fmtE[key].push(name); eS++; }
      }
    }

    // ── 포맷F: 첫 행=이름, 첫 열=날짜, 본문=체크마크 ──
    const fmtF: Record<string, string[]> = {}; let fS = 0;
    for (let c = range.s.c + 1; c <= range.e.c; c++) {
      const nms = extractNames(getCell(range.s.r, c));
      if (!nms.length) continue;
      for (let r = range.s.r + 1; r <= range.e.r; r++) {
        const key = dk(getCell(r, range.s.c));
        if (!key || !isCheckMark(getCell(r, c))) continue;
        if (!fmtF[key]) fmtF[key] = [];
        for (const name of nms) if (!fmtF[key].includes(name)) { fmtF[key].push(name); fS++; }
      }
    }

    // ── 포맷G: 주간 달력 — 요일 헤더행 + (날짜행, 이름행) 교대 ──
    // 구조: Row0=요일(일월화수목금토), Row1=날짜번호, Row2=이름들, Row3=날짜번호, Row4=이름들 ...
    const fmtG: Record<string, string[]> = {}; let gS = 0;
    // 첫 행이 요일 헤더인지 확인 (일/월/화/수/목/금/토 포함)
    const weekdayHeader = /^[일월화수목금토]$/;
    let headerRow = -1;
    for (let r = range.s.r; r <= Math.min(range.s.r + 2, range.e.r); r++) {
      let wdCount = 0;
      for (let c = range.s.c; c <= range.e.c; c++) {
        const v = String(getCell(r, c) ?? "").trim();
        if (weekdayHeader.test(v)) wdCount++;
      }
      if (wdCount >= 3) { headerRow = r; break; }
    }
    if (headerRow >= 0) {
      // 요일→열 목록
      const colsWithDates: number[] = [];
      for (let c = range.s.c; c <= range.e.c; c++) {
        const v = String(getCell(headerRow, c) ?? "").trim();
        if (weekdayHeader.test(v)) colsWithDates.push(c);
      }
      // headerRow 이후 행들을 2행씩 처리 (날짜행, 이름행)
      for (let r = headerRow + 1; r < range.e.r; r += 2) {
        const dateRow = r;
        const nameRow = r + 1;
        for (const c of colsWithDates) {
          const key = dk(getCell(dateRow, c));
          if (!key) continue;
          // 이름행의 같은 열에서만 이름 수집 (인접 열 제외)
          const names = extractNames(getCell(nameRow, c));
          if (!fmtG[key]) fmtG[key] = [];
          for (const name of names) if (!fmtG[key].includes(name)) { fmtG[key].push(name); gS++; }
        }
      }
    }

    debugLines.push(`  A:${aS} B:${bS} C:${cS} D:${dS} E:${eS} F:${fS} G:${gS}`);

    const candidates = [
      { map: fmtA, score: aS, desc: "A(행날짜→이름)" },
      { map: fmtB, score: bS, desc: "B(열날짜→이름)" },
      { map: fmtC, score: cS, desc: "C(행이름→날짜)" },
      { map: fmtD, score: dS, desc: "D(열이름→날짜)" },
      { map: fmtE, score: eS, desc: "E(격자체크-열이름)" },
      { map: fmtF, score: fS, desc: "F(격자체크-행이름)" },
      { map: fmtG, score: gS, desc: "G(주간달력)" },
    ];
    const best = candidates.reduce((a, b) => b.score > a.score ? b : a);
    if (best.score > 0) allResults.push(best);
  }

  const debug = debugLines.join("\n");
  if (!allResults.length) return { map: {}, debug };
  const winner = allResults.reduce((a, b) => Object.keys(b.map).length > Object.keys(a.map).length ? b : a);
  return { map: winner.map, debug };
}

// ── 엑셀 파싱 훅 ──────────────────────────────────
function useExcelData() {
  const [excelDays, setExcelDays] = useState<ExcelDayData[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [uploadedName, setUploadedName] = useState<string | null>(null); // 업로드 파일명

  async function load() {
    setLoading(true);
    setError("");
    try {
      const res = await fetch(`${BASE}/schedule.xlsx`);
      if (!res.ok) throw new Error("파일을 찾을 수 없습니다");
      const buf = await res.arrayBuffer();
      setExcelDays(parseExcelBuffer(buf));
      setUploadedName(null);
    } catch (e: any) {
      setError(e.message ?? "오류");
    } finally {
      setLoading(false);
    }
  }

  // ★ 기능1: 사용자가 업로드한 파일로 대체
  async function loadFromFile(file: File) {
    setLoading(true);
    setError("");
    try {
      const buf = await file.arrayBuffer();
      const days = parseExcelBuffer(buf);
      if (days.length === 0) throw new Error("날짜 데이터를 찾을 수 없습니다. '투입계산' 시트를 확인하세요.");
      setExcelDays(days);
      setUploadedName(file.name);
    } catch (e: any) {
      setError(e.message ?? "파일 읽기 오류");
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => { load(); }, []);

  return { excelDays, loading, error, reload: load, loadFromFile, uploadedName };
}

// ── 배정 엔진: 2부제 (PDF 규정 기준) ──────────────
// 배치 순서:
//   1부: 찾근자(선발3번~) → 조출자 → 일반순번 → [1부스페어]
//   2부: 일반순번 → 후출자(뒤에서 3번째) → 일반순번(막팀) → [2부스페어]
function assignDouble(
  names: string[],
  statuses: Record<string, StatusType>,
  shift1Size: number,
  shift2Size: number,
  daegeunMap: Record<string, string> = {}  // 대근 유형 맵 (1부|2부|투라운드)
): DayResult {
  const twoRound: string[] = [];   // 찾근 (1부+2부 투라운드)
  const 조출List: string[] = [];   // 조출 (1부 앞 고정, 최대 4명)
  const 후출List: string[] = [];   // 후출 (2부 뒤에서 3번째, 최대 4명)
  const 대기List: string[] = [];   // 대기 (1부 출근대기 → spare1로 2부 첫번째 고정)
  const 대근1부List: string[] = []; // 대근-1부: 1부만 근무 후 귀가
  const 대근2부List: string[] = []; // 대근-2부: 2부만 근무
  const excluded: string[] = [];
  const autoQueue: string[] = [];  // 일반 순번 대기열

  for (const name of names) {
    const s = statuses[name] ?? null;
    if (s === "찾근")  { twoRound.push(name); }
    else if (s === "대기") {
      대기List.push(name);
    } else if (s === "조출") {
      if (조출List.length < 4) 조출List.push(name); else autoQueue.push(name);
    } else if (s === "후출") {
      if (후출List.length < 4) 후출List.push(name); else autoQueue.push(name);
    } else if (EXCLUDED_SET.has(s ?? "")) { excluded.push(name); }
    else {
      // status null(정상근무) — 대근 유형 확인
      const dg = daegeunMap[name];
      if (dg === "1부")       대근1부List.push(name); // 1부만 출근
      else if (dg === "2부")  대근2부List.push(name); // 2부만 출근
      else                    autoQueue.push(name);
    }
  }

  // ── 1부 배치: 찾근 → 조출 → 대근1부 → 일반순번 ── (대기자는 1부 미포함)
  const fixed1 = [...twoRound, ...조출List, ...대근1부List];
  const avail1 = Math.max(0, shift1Size - fixed1.length);
  const shift1 = [...fixed1, ...autoQueue.slice(0, avail1)];
  // spare1: 명시적 대기자 우선, 없으면 autoQueue에서 순번상 다음번호
  const spare1 = 대기List.length > 0
    ? 대기List.slice(0, 1)
    : autoQueue.slice(avail1, avail1 + 1);
  const remaining = 대기List.length > 0
    ? autoQueue.slice(avail1)
    : autoQueue.slice(avail1 + 1);

  // ── 2부 배치 ────────────────────────────────────────
  // 규정:
  //   - 1부스페어(spare1)가 2부 제일 앞(첫번째)
  //   - 대근-2부: spare1 바로 뒤에 고정 (1부 미참여, 2부만 출근)
  //   - 찾근자(twoRound): 2부의 약 1/4 지점에 삽입
  //   - 후출자: 2부 뒤에서 3번째 위치
  //   - 2부스페어: 2부에 들어가지 못한 나머지 → 다음날 첫번호
  const shift1Regular = autoQueue.slice(0, avail1);
  const avail2Normal = Math.max(0, shift2Size - spare1.length - 대근2부List.length - twoRound.length - 후출List.length);
  const normalFor2 = remaining.slice(0, avail2Normal);
  const spare2fromRemaining = remaining.slice(avail2Normal); // remaining에서 2부 못 들어간 사람

  // 1부 돌고 온 사람 중 2부에 투라운드로 들어갈 인원 계산
  const extra2부Count = Math.max(0, shift2Size - spare1.length - 대근2부List.length - twoRound.length - normalFor2.length - 후출List.length);
  const extra2부 = shift1Regular.slice(0, extra2부Count);

  // ★ 2부 스페어: remaining 잔여 + 1부 배정에서 투라운드 못 한 사람
  //   (extra2부Count > 0 일 때만: 실제로 투라운드가 발생한 경우에만 투라운드 탈락자가 spare)
  const spare2fromShift1 = extra2부Count > 0 ? shift1Regular.slice(extra2부Count) : [];
  const spare2 = [...spare2fromRemaining, ...spare2fromShift1];

  // twoRound 삽입 위치: 2부의 약 1/4 지점 (spare1 + 대근2부List 뒤 기준)
  const twoRoundInsertAt = Math.max(0, Math.floor(shift2Size / 4) - spare1.length - 대근2부List.length);
  const normalBefore = normalFor2.slice(0, twoRoundInsertAt);
  const normalAfter  = normalFor2.slice(twoRoundInsertAt);

  const afterTwoRound = [...normalAfter, ...extra2부];
  let shift2: string[];
  if (후출List.length > 0 && afterTwoRound.length >= 2) {
    const insertAt = Math.max(0, afterTwoRound.length - 2);
    shift2 = [
      ...spare1,
      ...대근2부List,
      ...normalBefore,
      ...twoRound,
      ...afterTwoRound.slice(0, insertAt),
      ...후출List,
      ...afterTwoRound.slice(insertAt),
    ];
  } else {
    shift2 = [...spare1, ...대근2부List, ...normalBefore, ...twoRound, ...afterTwoRound, ...후출List];
  }

  // ── 다음날 예상 순번 ──────────────────────────────────
  let nextDayQueue: string[];
  {
    const exclSet2  = new Set(excluded);
    const spare1Set = new Set(spare1);
    const spare2Set = new Set(spare2);

    if (spare2.length > 0) {
      const twoRoundSet = new Set(twoRound);
      const firstSpares = spare2.slice(0, 2);       // 찾근 앞 첫번호
      const restSpares  = spare2.slice(2);           // 나머지 spare2 → 찾근 뒤
      const rest     = names.filter(n => !spare2Set.has(n) && !exclSet2.has(n));
      const twoInRest = rest.filter(n => twoRoundSet.has(n));  // 찾근 (spare2 제외)
      const normalRest = rest.filter(n => !twoRoundSet.has(n));
      const excls    = names.filter(n => exclSet2.has(n));
      nextDayQueue = [...firstSpares, ...twoInRest, ...restSpares, ...normalRest, ...excls];
    } else {
      const todayLast = shift2.at(-1);
      const rest  = names.filter(n => !spare1Set.has(n) && !exclSet2.has(n));
      const excls = names.filter(n => exclSet2.has(n));
      if (todayLast) {
        const li = rest.indexOf(todayLast);
        if (li >= 0 && rest.length > 1) {
          const startAt = (li + 1) % rest.length;
          nextDayQueue = [...rest.slice(startAt), ...rest.slice(0, startAt), ...excls];
        } else {
          nextDayQueue = [...rest, ...excls];
        }
      } else {
        nextDayQueue = [...rest, ...excls];
      }
    }
  }

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

  // 다음날 순번: spare2=0 이면 마지막 근무자 다음부터 회전
  let nextDayQueue: string[];
  {
    const exclSet2  = new Set(excluded);
    const spare2Set = new Set(spare2);

    if (spare2.length > 0) {
      const rest  = names.filter(n => !spare2Set.has(n) && !exclSet2.has(n));
      const excls = names.filter(n => exclSet2.has(n));
      nextDayQueue = [...spare2, ...rest, ...excls];
    } else {
      const todayLast = shift1.at(-1);
      const rest  = names.filter(n => !exclSet2.has(n));
      const excls = names.filter(n => exclSet2.has(n));
      if (todayLast) {
        const li = rest.indexOf(todayLast);
        if (li >= 0 && rest.length > 1) {
          const startAt = (li + 1) % rest.length;
          nextDayQueue = [...rest.slice(startAt), ...rest.slice(0, startAt), ...excls];
        } else {
          nextDayQueue = [...rest, ...excls];
        }
      } else {
        nextDayQueue = [...rest, ...excls];
      }
    }
  }

  return { twoRound, shift1, spare1: [], shift2: [], spare2, excluded, nextDayQueue };
}

// ── 다음날 예상 순번 계산기 ───────────────────────────
// 규정:
//   - 찾근은 당일만 적용 → 다음날은 일반 순번 대기
//   - 오늘 2부스페어(앞번호) → 나머지 전체 큐 순서 (찾근 여부 무시)
//   - 제외(병가·당번 등)는 맨 뒤
function buildNextDayQueue(
  allNames: string[],     // 전체 이름 (오늘 currentNames 큐 순서)
  spare1: string[],
  spare2: string[],
  twoRound: string[],     // 사용하지 않음 — 찾근은 당일만 적용
  excluded: string[]
): string[] {
  const spareSet = new Set([...spare1, ...spare2]);
  const exclSet  = new Set(excluded);

  // ① 오늘 2부스페어: 앞번호부터 (1부스페어는 오늘 2부도 나갔으므로 포함X)
  const spares = [...spare2];

  // ② 나머지: 스페어·제외 아닌 전원 → 큐 순서 그대로 (찾근 구분 없음)
  const rest = allNames.filter(n => !spareSet.has(n) && !exclSet.has(n));

  // ③ 제외자(병가·당번 등) → 맨 뒤
  const excls = allNames.filter(n => exclSet.has(n));

  return [...spares, ...rest, ...excls];
}



// ── OCR 달력 파싱 ────────────────────────────────
interface OcrWord {
  text: string;
  bbox: { x0: number; x1: number; y0: number; y1: number };
}

function parseCalendarOCR(words: OcrWord[]): Record<string, string[]> {
  const dateNums: { date: number; cx: number; y: number }[] = [];
  const korNames: { name: string; cx: number; y: number }[] = [];

  for (const w of words) {
    const raw = w.text.trim();
    const cx = (w.bbox.x0 + w.bbox.x1) / 2;
    const cy = (w.bbox.y0 + w.bbox.y1) / 2;

    // 날짜 숫자 (1~31)
    if (/^\d{1,2}$/.test(raw)) {
      const n = parseInt(raw, 10);
      if (n >= 1 && n <= 31) {
        dateNums.push({ date: n, cx, y: cy });
        continue;
      }
    }

    // 한국어 이름 (2~4자)
    const clean = raw.replace(/[^가-힣]/g, "");
    if (/^[가-힣]{2,4}$/.test(clean)) {
      korNames.push({ name: clean, cx, y: cy });
    }
  }

  const result: Record<string, string[]> = {};

  for (const nm of korNames) {
    let bestDate: number | null = null;
    let bestScore = Infinity;

    for (const dn of dateNums) {
      const dx = Math.abs(nm.cx - dn.cx);
      const dy = nm.y - dn.y; // 양수 = 이름이 날짜 아래
      if (dx > 120) continue;  // 다른 열
      if (dy < -30) continue;  // 이름이 날짜 위에 너무 많이 올라가면 제외

      const score = dx * 2 + Math.abs(dy) * 0.1;
      if (score < bestScore) { bestScore = score; bestDate = dn.date; }
    }

    if (bestDate !== null) {
      const key = String(bestDate);
      if (!result[key]) result[key] = [];
      if (!result[key].includes(nm.name)) result[key].push(nm.name);
    }
  }

  return result;
}

async function runCalendarOCR(file: File): Promise<Record<string, string[]>> {
  const worker = await createWorker("kor");
  try {
    const { data } = await worker.recognize(file);
    return parseCalendarOCR(data.words as OcrWord[]);
  } finally {
    await worker.terminate();
  }
}

// ── 메인 컴포넌트 ─────────────────────────────────
export default function SchedulePage() {
  const [, setLocation] = useLocation();
  const { excelDays, loading: xlLoading, error: xlError, loadFromFile, uploadedName } = useExcelData();

  // ── 날짜별 팀수 설정 (lotto_teamSettingsV2: { [dateLabel]: { mode, totalSize, shift1Size, singleSize, locked } }) ──
  function _readTeamMap(): Record<string, { mode: Mode; totalSize: number; shift1Size: number; singleSize: number; locked: boolean }> {
    try { return JSON.parse(localStorage.getItem("lotto_teamSettingsV2") ?? "{}"); } catch { return {}; }
  }
  function _writeTeamForDate(dateLabel: string, settings: { mode: Mode; totalSize: number; shift1Size: number; singleSize: number; locked: boolean }) {
    const map = _readTeamMap();
    map[dateLabel] = settings;
    localStorage.setItem("lotto_teamSettingsV2", JSON.stringify(map));
  }

  // React state 업데이트 비동기 문제 방지: 현재 선택 날짜를 ref로 동기 추적
  const activeDateLabelRef = useRef<string>("");

  const [mode, setMode] = useState<Mode>("단부제");
  // 2부제: totalSize = 총팀수, shift1Size = 1부팀수, shift2Size = 총팀수 - 1부팀수
  const [totalSize, setTotalSize] = useState<number>(70);
  const [shift1Size, setShift1Size] = useState<number>(35);
  const shift2Size = Math.max(0, totalSize - shift1Size);
  const [singleSize, setSingleSize] = useState<number>(60);
  // 팀수 설정 잠금 (저장 완료 상태)
  const [teamsLocked, setTeamsLocked] = useState<boolean>(false);

  // 팀수 설정 저장 — ref로 동기 추적한 날짜 기준 저장 (React 상태 비동기 문제 방지)
  function saveTeamSettings() {
    const dateLabel = activeDateLabelRef.current || selectedDate?.dateLabel;
    if (!dateLabel) return;
    _writeTeamForDate(dateLabel, { mode, totalSize, shift1Size, singleSize, locked: true });
    setTeamsLocked(true);
  }
  function unlockTeamSettings() {
    const dateLabel = activeDateLabelRef.current || selectedDate?.dateLabel;
    if (!dateLabel) return;
    _writeTeamForDate(dateLabel, { mode, totalSize, shift1Size, singleSize, locked: false });
    setTeamsLocked(false);
  }
  const [nameText, setNameText] = useState("");
  const [dayOfWeek, setDayOfWeek] = useState(0);

  // 선택된 날짜 (엑셀 날짜)
  const [selectedDate, setSelectedDate] = useState<ExcelDayData | null>(null);

  // 월 네비게이션 상태 ("04", "05", "06" …)
  const [viewMonth, setViewMonth] = useState<string>(() => {
    const now = new Date();
    return String(now.getMonth() + 1).padStart(2, "0");
  });

  // 기준 연도: 엑셀이 있으면 거기서 추출, 없으면 현재 연도
  const viewYear = useMemo(() => new Date().getFullYear(), []);

  // 엑셀이 있는 달 목록 ── 자동 선택 시 참고용
  // (월 네비게이션 자체는 viewMonth 기반으로 엑셀과 무관하게 동작)

  // ── viewDays: 현재 월 전체 날짜 생성 + 엑셀 데이터 오버레이 ──
  const viewDays = useMemo(() => {
    const generated = generateMonthDays(viewMonth, viewYear);
    const excelMap = new Map(excelDays.map(d => [d.dateLabel, d]));
    return generated.map(d => excelMap.get(d.dateLabel) ?? d);
  }, [viewMonth, viewYear, excelDays]);

  // ── 월 네비게이션 (01-12 범위, 엑셀 무관) ──
  const prevMonthStr = useMemo(() => {
    const n = parseInt(viewMonth, 10);
    return n > 1 ? String(n - 1).padStart(2, "0") : null;
  }, [viewMonth]);
  const nextMonthStr = useMemo(() => {
    const n = parseInt(viewMonth, 10);
    return n < 12 ? String(n + 1).padStart(2, "0") : null;
  }, [viewMonth]);

  // 엑셀 로드 완료 시 오늘 날짜 자동 선택
  useEffect(() => {
    if (excelDays.length > 0 && !selectedDate) {
      const now = new Date();
      const mm = String(now.getMonth() + 1).padStart(2, "0");
      const dd = String(now.getDate()).padStart(2, "0");
      const todayPrefix = `${mm}.${dd}`;
      const today = excelDays.find((d) => d.dateLabel.startsWith(todayPrefix));
      if (today) {
        selectExcelDate(today);
        setViewMonth(mm);
      }
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [excelDays]);

  // 선택 날짜가 바뀌면 viewMonth도 해당 달로 이동
  useEffect(() => {
    if (selectedDate) {
      const m = selectedDate.dateLabel.substring(0, 2);
      setViewMonth(m);
    }
  }, [selectedDate]);

  // 인원
  const [names, setNames] = useState<string[]>([]);
  const [rosterLoaded, setRosterLoaded] = useState(false);

  // 날짜별 수동 상태 (localStorage 영구 저장) — 연도별 키로 분리
  const DS_KEY = `lotto_dateStatuses_${new Date().getFullYear()}`;
  const [dateStatuses, setDateStatuses] = useState<Record<string, Record<string, StatusType>>>(() => {
    try {
      // ★ Bug1 수정: 구버전 키("lotto_dateStatuses") → 연도 키로 마이그레이션
      const oldKey = "lotto_dateStatuses";
      const oldData = localStorage.getItem(oldKey);
      if (oldData && !localStorage.getItem(DS_KEY)) {
        localStorage.setItem(DS_KEY, oldData);
        localStorage.removeItem(oldKey);
      }
      const saved = localStorage.getItem(DS_KEY);
      return saved ? JSON.parse(saved) : {};
    } catch { return {}; }
  });
  useEffect(() => {
    localStorage.setItem(DS_KEY, JSON.stringify(dateStatuses));
  }, [dateStatuses, DS_KEY]);

  // 병가 지속 상태 (해제 전까지 모든 날짜에 자동 적용)
  const SL_KEY = `lotto_sickLeave_${new Date().getFullYear()}`;
  const [sickLeave, setSickLeave] = useState<Record<string, boolean>>(() => {
    try {
      const saved = localStorage.getItem(SL_KEY);
      return saved ? JSON.parse(saved) : {};
    } catch { return {}; }
  });
  useEffect(() => {
    localStorage.setItem(SL_KEY, JSON.stringify(sickLeave));
  }, [sickLeave, SL_KEY]);

  // 휴무 엑셀 데이터 (날짜 "MM.DD" → 이름 리스트)
  const HM_KEY = "lotto_holidayMap";
  const HA_KEY = "lotto_holidayApplied";
  const [holidayMap, setHolidayMap] = useState<Record<string, string[]>>(() => {
    try { return JSON.parse(localStorage.getItem(HM_KEY) ?? "{}"); } catch { return {}; }
  });
  const [holidayFileName, setHolidayFileName] = useState<string | null>(() =>
    localStorage.getItem("lotto_holidayFileName")
  );
  // 날짜별 "이미 엑셀 휴무 자동 적용 완료" 여부 추적 → 재선택 시 덮어쓰기 방지
  const [holidayAppliedDates, setHolidayAppliedDates] = useState<Record<string, true>>(() => {
    try { return JSON.parse(localStorage.getItem(HA_KEY) ?? "{}"); } catch { return {}; }
  });
  useEffect(() => {
    localStorage.setItem(HM_KEY, JSON.stringify(holidayMap));
  }, [holidayMap]);
  useEffect(() => {
    localStorage.setItem(HA_KEY, JSON.stringify(holidayAppliedDates));
  }, [holidayAppliedDates]);

  function loadHolidayFile(file: File) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const buf = e.target?.result as ArrayBuffer;
        const { map, debug } = parseHolidayExcelBuffer(buf, parseInt(viewMonth, 10));
        const dateCount = Object.keys(map).length;
        if (dateCount === 0) {
          alert(
            "휴무 데이터를 인식하지 못했습니다.\n\n" +
            "── 진단 정보 (개발자에게 전달) ──\n" +
            debug +
            "\n\n지원 포맷: 행/열 기반 날짜+이름, 체크마크 표"
          );
          return;
        }
        // ── 기존 holidayMap 완전 교체 (이전 월 데이터 제거) ──
        setHolidayMap(map);
        setHolidayFileName(file.name);
        localStorage.setItem("lotto_holidayFileName", file.name);
        localStorage.setItem("lotto_holidayMap", JSON.stringify(map));

        // ── "자동 적용 완료" 기록 초기화 → 모든 날짜에 새 엑셀로 다시 적용 가능 ──
        setHolidayAppliedDates({});
        localStorage.setItem("lotto_holidayApplied", "{}");

        // ── 기존 날짜별 휴무 상태 자동 초기화 (휴무 항목만 제거, 다른 상태 유지) ──
        setDateStatuses(prev => {
          const next: typeof prev = {};
          for (const [dl, statuses] of Object.entries(prev)) {
            const cleaned: Record<string, StatusType> = {};
            for (const [name, st] of Object.entries(statuses)) {
              if (st !== "휴무") cleaned[name] = st;
            }
            if (Object.keys(cleaned).length > 0) next[dl] = cleaned;
          }
          return next;
        });

        const totalPeople = Object.values(map).reduce((s, a) => s + a.length, 0);
        alert(`✅ 휴무 엑셀 업로드 완료!\n${dateCount}개 날짜 · 총 ${totalPeople}건\n기존 휴무 데이터는 자동 초기화됐습니다.`);
      } catch (err) {
        alert("엑셀 파일 읽기 실패: " + String(err));
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // 대근 날짜별 저장 (localStorage)
  const DG_KEY = `lotto_daegeun_${new Date().getFullYear()}`;
  const [dateDaegeun, setDateDaegeun] = useState<Record<string, Record<string, DaegeunType>>>(() => {
    try {
      const saved = localStorage.getItem(`lotto_daegeun_${new Date().getFullYear()}`);
      return saved ? JSON.parse(saved) : {};
    } catch { return {}; }
  });
  useEffect(() => {
    localStorage.setItem(DG_KEY, JSON.stringify(dateDaegeun));
  }, [dateDaegeun, DG_KEY]);

  // VIP 날짜별 저장 (localStorage)
  type VipRound = "1부" | "2부" | null;
  interface VipEntry { count: number; round: VipRound; members: string[] }
  const [vipData, setVipData] = useState<Record<string, VipEntry>>(() => {
    try {
      const saved = localStorage.getItem("lotto_vipData");
      return saved ? JSON.parse(saved) : {};
    } catch { return {}; }
  });
  useEffect(() => {
    localStorage.setItem("lotto_vipData", JSON.stringify(vipData));
  }, [vipData]);

  const [vipPickerOpen, setVipPickerOpen] = useState(false);
  const [vipMemberPickerOpen, setVipMemberPickerOpen] = useState(false);
  const [vipMemberSearch, setVipMemberSearch] = useState("");

  // ★ 기능2: 날짜별 2부스페어 저장 (다음날 첫번호 힌트용)
  const SPARE2_KEY = `lotto_spare2_${new Date().getFullYear()}`;
  const [savedSpare2, setSavedSpare2] = useState<Record<string, string[]>>(() => {
    try {
      const saved = localStorage.getItem(SPARE2_KEY);
      return saved ? JSON.parse(saved) : {};
    } catch { return {}; }
  });
  useEffect(() => {
    localStorage.setItem(SPARE2_KEY, JSON.stringify(savedSpare2));
  }, [savedSpare2, SPARE2_KEY]);

  // 이전날 날짜 레이블 찾기
  const prevDateLabel = useMemo(() => {
    if (!selectedDate) return null;
    const idx = excelDays.findIndex(d => d.dateLabel === selectedDate.dateLabel);
    return idx > 0 ? excelDays[idx - 1].dateLabel : null;
  }, [selectedDate, excelDays]);

  // 오늘 첫번호 힌트 = 전날 2부스페어[0]
  const todayFirstHint = prevDateLabel ? (savedSpare2[prevDateLabel]?.[0] ?? null) : null;

  // 현재 선택 날짜 키 (e.g. "04.01 (수)")
  const currentDateKey = selectedDate?.dateLabel ?? "";

  const currentVip: VipEntry = useMemo(
    () => vipData[currentDateKey] ?? { count: 0, round: null, members: [] },
    [vipData, currentDateKey]
  );
  function setCurrentVip(v: Partial<VipEntry>) {
    if (!currentDateKey) return;
    setVipData(prev => {
      const cur = prev[currentDateKey] ?? { count: 0, round: null, members: [] };
      const next = { ...cur, ...v };
      if (next.count === 0 && !next.round && (!next.members || next.members.length === 0)) {
        const upd = { ...prev };
        delete upd[currentDateKey];
        return upd;
      }
      return { ...prev, [currentDateKey]: next };
    });
  }

  // 현재 날짜의 수동 상태 (derived)
  const manualStatuses = useMemo(
    () => dateStatuses[currentDateKey] ?? {},
    [dateStatuses, currentDateKey]
  );

  // 현재 날짜의 대근 맵 (derived)
  const currentDaegeun = useMemo(
    () => dateDaegeun[currentDateKey] ?? {},
    [dateDaegeun, currentDateKey]
  );

  // 대근 모달 (어느 사람의 대근 선택 중인지)
  const [daegeunModal, setDaegeunModal] = useState<string | null>(null);

  // 대근 일괄 모달 (비근무 그룹 전체 목록)
  const [batchDaegeunOpen, setBatchDaegeunOpen] = useState(false);
  const [batchDaegeunSearch, setBatchDaegeunSearch] = useState("");

  function setDaegeunForDate(name: string, type: DaegeunType) {
    if (!currentDateKey) return;
    setDateDaegeun(prev => ({
      ...prev,
      [currentDateKey]: { ...(prev[currentDateKey] ?? {}), [name]: type },
    }));
    setDaegeunModal(null);
  }

  function cancelDaegeun(name: string) {
    if (!currentDateKey) return;
    setDateDaegeun(prev => {
      const cur = { ...(prev[currentDateKey] ?? {}) };
      delete cur[name];
      if (Object.keys(cur).length === 0) {
        const updated = { ...prev };
        delete updated[currentDateKey];
        return updated;
      }
      return { ...prev, [currentDateKey]: cur };
    });
    setDaegeunModal(null);
  }

  // 현재 날짜 수동 상태 세터
  function setManualStatuses(
    updater: ((prev: Record<string, StatusType>) => Record<string, StatusType>) | Record<string, StatusType>
  ) {
    setDateStatuses((prev) => {
      const cur = prev[currentDateKey] ?? {};
      const next = typeof updater === "function" ? updater(cur) : updater;
      if (Object.keys(next).length === 0) {
        const updated = { ...prev };
        delete updated[currentDateKey];
        return updated;
      }
      return { ...prev, [currentDateKey]: next };
    });
  }

  // 결과
  const [dayResult, setDayResult] = useState<DayResult | null>(null);
  const [weekly, setWeekly] = useState<{ day: string; result: DayResult }[]>([]);
  // 주간 근무표 날짜별 개별 토글 (기본: 요약 보기)
  const [expandedDays, setExpandedDays] = useState<Set<string>>(new Set());
  function toggleDayExpand(day: string) {
    setExpandedDays(prev => {
      const next = new Set(prev);
      if (next.has(day)) next.delete(day); else next.add(day);
      return next;
    });
  }
  const [view, setView] = useState<"input" | "assign">("input");

  // 상태 선택 모달 (전체 순번표 표시)
  const [modalStatus, setModalStatus] = useState<StatusType | null>(null);
  const [modalSearch, setModalSearch] = useState("");

  // 명단 보기 모달 (해당 상태인 사람만 표시)
  const [viewStatusModal, setViewStatusModal] = useState<"당번" | "휴무" | "병가" | null>(null);

  // ── OCR 상태 ────────────────────────────────────
  const ocrFileRef = useRef<HTMLInputElement>(null);
  const [ocrState, setOcrState] = useState<"idle" | "running" | "done" | "error">("idle");
  const [ocrProgress, setOcrProgress] = useState(0);
  const [ocrAllDates, setOcrAllDates] = useState<Record<string, string[]>>({});
  const [ocrPreview, setOcrPreview] = useState<string[]>([]);  // 현재 날짜 추출 이름들

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
  const rosterImportRef = useRef<HTMLInputElement>(null);
  const [rosterForm, setRosterForm] = useState<{ mode: "add"|"edit"; orig?: PersonData; name: string; 조: 1|2|3|4; group: GroupType } | null>(null);

  // 현재 요일의 유효 상태 반환
  function effectiveStatus(name: string, dayIdx: number = dayOfWeek): StatusType {
    // 명시적 오버라이드(해당 날짜에 직접 설정된 값)가 있으면 우선
    if (name in manualStatuses) return manualStatuses[name];
    // 병가 지속: sickLeave에 등록된 사람은 해제 전까지 자동 병가
    if (sickLeave[name]) return "병가";
    const person = customRosterMap[name];
    if (person && isAutoOff(person.group, dayIdx)) {
      const dg = currentDaegeun[name];
      if (dg === "투라운드") return "찾근";
      if (dg === "1부" || dg === "2부") return null;
      return "휴무";
    }
    return null;
  }

  // 상태 토글
  function toggleStatus(name: string, btn: StatusType) {
    if (btn === "병가") {
      const cur = effectiveStatus(name);
      if (cur === "병가") {
        // 병가 해제: sickLeave에서 제거 + 해당 날짜 manualStatuses도 정리
        setSickLeave((prev) => {
          const next = { ...prev };
          delete next[name];
          return next;
        });
        setManualStatuses((prev) => {
          if (!(name in prev)) return prev;
          const next = { ...prev };
          delete next[name];
          return next;
        });
      } else {
        // 병가 지정: sickLeave에 추가, 다른 상태 있으면 제거
        setSickLeave((prev) => ({ ...prev, [name]: true }));
        setManualStatuses((prev) => {
          if (!(name in prev)) return prev;
          const next = { ...prev };
          delete next[name];
          return next;
        });
      }
      return;
    }
    // 병가 외 상태 토글 (기존 로직)
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

  // 날짜 선택 → 팀수/요일 설정 (저장된 값 우선, 없으면 예약팀수 자동 반영)
  function selectExcelDate(day: ExcelDayData) {
    // ref를 먼저 동기 업데이트 → saveTeamSettings가 올바른 날짜에 저장
    activeDateLabelRef.current = day.dateLabel;
    setSelectedDate(day);
    setDayOfWeek(day.dayIdx);
    const saved = _readTeamMap()[day.dateLabel];
    if (saved) {
      // 해당 날짜에 저장된 팀수 설정 복원
      setMode(saved.mode ?? "단부제");
      setTotalSize(saved.totalSize ?? 70);
      setShift1Size(saved.shift1Size ?? 35);
      setSingleSize(saved.singleSize ?? 60);
      setTeamsLocked(saved.locked ?? false);
    } else if (day.예약팀수 > 0) {
      // 저장 없음 → 예약팀수로 자동 설정, 모드는 현재 그대로 유지
      const total = day.예약팀수;
      const half = Math.round(total / 2);
      setTotalSize(total);
      setShift1Size(half);
      setSingleSize(total);
      setTeamsLocked(false);
    } else {
      // 기본값으로 팀수만 리셋, 모드는 현재 그대로 유지
      setTotalSize(70);
      setShift1Size(35);
      setSingleSize(60);
      setTeamsLocked(false);
    }

    // 휴무 엑셀 자동 입력 — 최초 1회만 (이미 적용된 날짜는 사용자 수정값 유지)
    const dk = day.dateLabel.slice(0, 5);
    const hdNames = holidayMap[dk];
    if (hdNames && hdNames.length > 0 && !holidayAppliedDates[day.dateLabel]) {
      setDateStatuses(prev => {
        const cur = prev[day.dateLabel] ?? {};
        const next = { ...cur };
        for (const hName of hdNames) {
          const matched = sortedCustomRoster.find(p =>
            p.name === hName || (hName.length >= 2 && p.name.startsWith(hName.slice(0, 2)))
          )?.name ?? hName;
          // 이미 사용자가 다른 상태로 지정한 경우 덮어쓰지 않음
          if (!next[matched]) next[matched] = "휴무";
        }
        return { ...prev, [day.dateLabel]: next };
      });
      // 이 날짜를 "자동 적용 완료"로 표시 → 다음에 다시 와도 덮어쓰기 안 함
      setHolidayAppliedDates(prev => ({ ...prev, [day.dateLabel]: true }));
    }
  }

  // ── 첫번호 지정 (세션 전용 — localStorage 저장 X, 다음날 이어지지 않음) ───
  const TODAY_KEY = `lotto_queueStart_${new Date().toISOString().slice(0, 10)}`;
  const [queueStartName, setQueueStartName] = useState<string | null>(() =>
    localStorage.getItem(TODAY_KEY)
  );
  const [queueModal, setQueueModal] = useState<"ask" | "pick" | null>(null);
  const [queuePickSearch, setQueuePickSearch] = useState("");
  // 모달 드래그 위치
  const [queueModalPos, setQueueModalPos] = useState({ x: 0, y: 0 });
  const dragRef = useRef<{ startX: number; startY: number; origX: number; origY: number } | null>(null);
  function onModalDragStart(e: React.MouseEvent | React.TouchEvent) {
    const clientX = "touches" in e ? e.touches[0].clientX : (e as React.MouseEvent).clientX;
    const clientY = "touches" in e ? e.touches[0].clientY : (e as React.MouseEvent).clientY;
    dragRef.current = { startX: clientX, startY: clientY, origX: queueModalPos.x, origY: queueModalPos.y };
    const onMove = (ev: MouseEvent | TouchEvent) => {
      if (!dragRef.current) return;
      const cx = "touches" in ev ? (ev as TouchEvent).touches[0].clientX : (ev as MouseEvent).clientX;
      const cy = "touches" in ev ? (ev as TouchEvent).touches[0].clientY : (ev as MouseEvent).clientY;
      setQueueModalPos({ x: dragRef.current.origX + cx - dragRef.current.startX, y: dragRef.current.origY + cy - dragRef.current.startY });
    };
    const onUp = () => { dragRef.current = null; document.removeEventListener("mousemove", onMove); document.removeEventListener("mouseup", onUp); document.removeEventListener("touchmove", onMove); document.removeEventListener("touchend", onUp); };
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
    document.addEventListener("touchmove", onMove);
    document.addEventListener("touchend", onUp);
  }

  function saveQueueStart(name: string | null) {
    setQueueStartName(name);
    // 오늘 날짜 key로 저장 — 자정 지나면 다른 key가 되어 자동 만료
    if (name) localStorage.setItem(TODAY_KEY, name);
    else localStorage.removeItem(TODAY_KEY);
  }

  // 이름 배열을 startName 위치부터 회전
  function rotateNames(base: string[], startName: string | null): string[] {
    if (!startName) return base;
    const idx = base.indexOf(startName);
    if (idx <= 0) return base;
    return [...base.slice(idx), ...base.slice(0, idx)];
  }

  // 실제로 names 적용 (회전 포함)
  function applyRoster(startName: string | null) {
    const base = sortedCustomRoster.map((p) => p.name);
    const rotated = rotateNames(base, startName);
    setNames(rotated);
    setRosterLoaded(true);
    setDayResult(null);
    setWeekly([]);
    setView("assign");
  }

  // 홈 이동 후 복귀 시 자동 재적용
  // — customRoster가 있고 queueStartName이 저장돼 있으면 바로 배정 화면으로
  useEffect(() => {
    if (sortedCustomRoster.length > 0 && queueStartName && names.length === 0) {
      const base = sortedCustomRoster.map((p) => p.name);
      setNames(rotateNames(base, queueStartName));
      setRosterLoaded(true);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // 순번표 불러오기 → 첫번호 지정돼 있으면 바로 배정, 없으면 ask 모달
  function loadRoster() {
    if (queueStartName) {
      applyRoster(queueStartName);
    } else {
      setQueueModalPos({ x: 0, y: 0 });
      setQueueModal("ask");
    }
  }

  // 직접 입력으로 다음 단계
  function confirmNames() {
    const parsed = nameText.split("\n").map((n) => n.trim()).filter(Boolean);
    if (!parsed.length) return;
    setNames(parsed);
    setRosterLoaded(false);
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

  // ── 순번표 백업 / 복원 ────────────────────────────
  function exportRoster() {
    const now = new Date();
    const ymd = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,"0")}${String(now.getDate()).padStart(2,"0")}`;
    const blob = new Blob([JSON.stringify(customRoster, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `순번표_백업_${ymd}.json`;
    a.click();
    URL.revokeObjectURL(url);
  }

  function handleImportRoster(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    e.target.value = "";
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = JSON.parse(ev.target?.result as string) as PersonData[];
        if (!Array.isArray(data) || data.some(d => typeof d.name !== "string")) {
          alert("올바른 순번표 백업 파일이 아닙니다.");
          return;
        }
        if (!confirm(`백업 파일에서 ${data.length}명을 불러옵니다. 현재 순번표를 덮어쓸까요?`)) return;
        setCustomRoster(data);
        alert(`✅ ${data.length}명 복원 완료`);
      } catch {
        alert("파일 파싱 오류 — JSON 형식을 확인해 주세요.");
      }
    };
    reader.readAsText(file);
  }

  function getEffective(dayIdx: number = dayOfWeek) {
    const base: Record<string, StatusType> = {};
    names.forEach((n) => { base[n] = effectiveStatus(n, dayIdx); });
    return base;
  }

  // ── 실시간 배정 미리보기 ──────────────────────────
  // 현재 상태(휴무/조출/...)를 반영한 배정 경계 계산 (배정 버튼 누르지 않아도 표시)
  const livePreview = useMemo(() => {
    if (names.length === 0) return null;
    const statuses = getEffective(dayOfWeek);
    return mode === "2부제"
      ? assignDouble(names, statuses, shift1Size, shift2Size, currentDaegeun)
      : assignSingle(names, statuses, singleSize);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [names, manualStatuses, currentDaegeun, mode, shift1Size, shift2Size, singleSize, dayOfWeek, customRosterMap]);

  // 이름 → 배정 카테고리 맵 (live)
  const liveCategoryMap = useMemo<Record<string, "1부" | "1부스페어" | "2부" | "2부스페어" | "스페어" | "단부" | "찾근" | "제외">>(() => {
    if (!livePreview) return {};
    const map: Record<string, "1부" | "1부스페어" | "2부" | "2부스페어" | "스페어" | "단부" | "찾근" | "제외"> = {};
    livePreview.twoRound?.forEach((n: string) => { map[n] = "찾근"; });
    livePreview.shift1?.forEach((n: string) => { map[n] = mode === "2부제" ? "1부" : "단부"; });
    livePreview.shift2?.forEach((n: string) => { map[n] = "2부"; });
    livePreview.spare2?.forEach((n: string) => { map[n] = mode === "2부제" ? "2부스페어" : "스페어"; });
    livePreview.excluded?.forEach((n: string) => { map[n] = "제외"; });
    // spare1(1부스페어)은 shift2에도 포함되므로 마지막에 덮어써야 "1부스페어" 표시 유지
    livePreview.spare1?.forEach((n: string) => { map[n] = "1부스페어"; });
    return map;
  }, [livePreview, mode]);

  // 상태 버튼 목록
  const STATUS_BTNS: { st: StatusType; label: string; color: string; bg: string }[] = [
    { st: "휴무",  label: "휴무",  color: "#757575", bg: "#f5f5f5" },
    { st: "당번",  label: "당번",  color: "#e65100", bg: "#fff3e0" },
    { st: "대기",  label: "1부대기", color: "#9d174d", bg: "#fce7f3" },
    { st: "조출",  label: "조출",  color: "#1565c0", bg: "#e3f2fd" },
    { st: "후출",  label: "후출",  color: "#6a1b9a", bg: "#f3e5f5" },
    { st: "찾근",  label: "찾근",  color: "#2e7d32", bg: "#e8f5e9" },
  ];

  function openStatusPicker(st: StatusType) {
    if (names.length === 0) {
      const base = sortedCustomRoster.map((p) => p.name);
      setNames(rotateNames(base, queueStartName));
      setRosterLoaded(true);
    }
    setModalStatus(st);
    setModalSearch("");
  }

  // OCR 핸들러: 이미지 업로드 → Tesseract OCR → 날짜별 이름 추출
  async function handleOcrFile(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    e.target.value = "";

    setOcrState("running");
    setOcrProgress(0);
    setOcrPreview([]);

    try {
      // Tesseract.js는 파일 크기에 따라 30초~2분 소요
      // progress 가상 타이머 (UI 피드백용)
      const timer = setInterval(() => {
        setOcrProgress((p) => Math.min(p + 4, 90));
      }, 600);

      const result = await runCalendarOCR(file);

      clearInterval(timer);
      setOcrProgress(100);
      setOcrAllDates(result);

      // 현재 선택된 날짜의 결과를 미리보기로 표시
      if (currentDateKey) {
        const dayNum = currentDateKey.split(".")[1]?.split(" ")[0]?.replace(/^0/, "");
        const dayNames = dayNum ? (result[dayNum] ?? []) : [];
        setOcrPreview(dayNames);
      }

      setOcrState("done");
    } catch (err) {
      console.error("OCR error:", err);
      setOcrState("error");
    }
  }

  // OCR 결과를 현재 날짜에 휴무로 자동 적용
  function applyOcrToCurrentDate() {
    if (!currentDateKey || ocrPreview.length === 0) return;

    // 순번표가 없으면 자동 로드 (첫번호 회전 적용)
    let currentNames = names;
    if (currentNames.length === 0) {
      const base = sortedCustomRoster.map((p) => p.name);
      currentNames = rotateNames(base, queueStartName);
      setNames(currentNames);
      setRosterLoaded(true);
    }

    setManualStatuses((prev) => {
      const next = { ...prev };
      for (const ocrName of ocrPreview) {
        // 순번표에 있는 이름과 매칭 (이름이 포함되어 있거나 앞 2자가 같을 때)
        const matched = currentNames.find(
          (n) => n === ocrName || n.startsWith(ocrName.slice(0, 2))
        );
        const key = matched ?? ocrName;
        next[key] = "휴무";
      }
      return next;
    });

    setOcrPreview([]);
    setOcrState("idle");
  }

  // OCR 전체 결과에서 현재 날짜 미리보기 갱신
  useEffect(() => {
    if (ocrState === "done" && currentDateKey && Object.keys(ocrAllDates).length > 0) {
      const dayNum = currentDateKey.split(".")[1]?.split(" ")[0]?.replace(/^0/, "");
      const preview = dayNum ? (ocrAllDates[dayNum] ?? []) : [];
      setOcrPreview(preview);
    }
  }, [currentDateKey, ocrAllDates, ocrState]);

  function assign() {
    const statuses = getEffective(dayOfWeek);
    const result = mode === "2부제"
      ? assignDouble(names, statuses, shift1Size, shift2Size, currentDaegeun)
      : assignSingle(names, statuses, singleSize);
    setDayResult(result);
    setWeekly([]);
    // ★ 기능2: 이 날짜의 2부스페어 자동 저장 (다음날 첫번호 힌트)
    if (currentDateKey && result.spare2.length > 0) {
      setSavedSpare2(prev => ({ ...prev, [currentDateKey]: result.spare2 }));
    }
  }

  function generateWeek() {
    // 현재 선택된 날짜 기준으로 해당 주(월~일)의 excelDay 찾기
    const selIdx = selectedDate
      ? excelDays.findIndex(d => d.dateLabel === selectedDate.dateLabel)
      : -1;
    const mondayIdx = selIdx >= 0 ? selIdx - (selectedDate?.dayIdx ?? 0) : -1;

    // 첫날은 현재 names 배열 사용 (queueStartName 회전 이미 적용됨)
    // 다음날부터는 전날 2부스페어 첫번째 인원을 첫번호로 회전
    let currentNames = [...names];

    const results = DAY_LABELS.reduce<{ day: string; result: DayResult }[]>((acc, day, di) => {
      // ★ Bug3 수정: mondayIdx + di >= 0 체크 (mondayIdx 자체가 음수여도 di로 보정)
      const absIdx = mondayIdx + di;
      const weekDay = absIdx >= 0 && absIdx < excelDays.length
        ? excelDays[absIdx]
        : null;
      const dateLabel = weekDay?.dateLabel ?? "";
      const dayIdx    = weekDay?.dayIdx ?? di;

      // ★ 날짜별 저장된 상태 사용 → 모든 상태는 지정한 날에만!
      const savedDay = dateStatuses[dateLabel] ?? {};
      const statuses: Record<string, StatusType> = {};
      currentNames.forEach((n) => {
        if (n in savedDay) { statuses[n] = savedDay[n]; return; }
        const person = customRosterMap[n];
        if (person && isAutoOff(person.group, dayIdx)) { statuses[n] = "휴무"; return; }
        statuses[n] = null;
      });

      // 날짜별 예약팀수 자동 반영
      let s1 = shift1Size, s2 = shift2Size, ss = singleSize;
      if (weekDay && weekDay.예약팀수 > 0) {
        const tot = weekDay.예약팀수;
        s1 = Math.round(tot / 2); s2 = tot - s1; ss = tot;
      }

      const result = mode === "2부제"
        ? assignDouble(currentNames, statuses, s1, s2, dateDaegeun[dateLabel] ?? {})
        : assignSingle(currentNames, statuses, ss);

      // ── 다음날 currentNames 재정렬 (규정: 2부스페어 앞번호 → 나머지 순번 대기) ──
      {
        const spare2Set = new Set(result.spare2);

        // 찾근은 당일만 적용 — 내일 큐 재정렬 시 찾근 구분 없음
        // ① 오늘 2부스페어 → 앞번호 순서 (내일 첫번호)
        const nextSpares = [...result.spare2];

        // ② 나머지 전원 (스페어 제외, 찾근 여부 무관) → 큐 순서 유지
        let nextRest = currentNames.filter(n => !spare2Set.has(n));

        if (result.spare2.length > 0) {
          // spare2 있음: rest 순서는 그대로 유지
        } else {
          // spare2 없음: 오늘 마지막 근무자 다음 번호부터 시작
          const todayLast = (mode === "2부제" ? result.shift2 : result.shift1).at(-1);
          if (todayLast) {
            const li = nextRest.indexOf(todayLast);
            if (li >= 0 && nextRest.length > 1) {
              const startAt = (li + 1) % nextRest.length;
              nextRest = [...nextRest.slice(startAt), ...nextRest.slice(0, startAt)];
            }
          }
        }

        currentNames = [...nextSpares, ...nextRest];
      }

      return [...acc, { day: dateLabel || day, result }];
    }, []);
    setWeekly(results);
    setDayResult(null);
    // ★ 기능2: 일주일 각 날짜의 2부스페어 자동 저장
    const newSpare2: Record<string, string[]> = {};
    results.forEach(({ day, result: r }) => {
      if (day && r.spare2.length > 0) newSpare2[day] = r.spare2;
    });
    if (Object.keys(newSpare2).length > 0) {
      setSavedSpare2(prev => ({ ...prev, ...newSpare2 }));
    }
  }

  // 활성 인원 대기열(제외·찾근 제외)에서의 순번 인덱스
  function activeQueueIndex(name: string): number {
    const active = names.filter((n) => {
      const s = effectiveStatus(n);
      return s !== "찾근" && !EXCLUDED_SET.has(s ?? "");
    });
    return active.indexOf(name);
  }

  // 찾근 가능 여부
  // 2부제: 1부 6팀 이상 + 1부 또는 2부 배정(qi < totalSize) + 이미 찾근(투라운드)이 아닐 것
  //        → 이미 찾근 상태(투라운드 중)인 사람은 본인 순번으로 일해야 하므로 불가
  // 단부제: 근무 안 될 때(singleSize 초과)만 가능
  function canChakgeun(name: string): boolean {
    const s = effectiveStatus(name);
    if (s === "찾근") return false; // 이미 투라운드 → 불가 (본인 순번으로 일해야 함)
    if (mode === "단부제") {
      // 단부제: 근무 안 될 때(singleSize 초과)만 가능
      const qi = activeQueueIndex(name);
      return qi > singleSize;
    }
    // 2부제: 1부 6팀 이상 필수
    if (shift1Size < 6) return false;
    // 1부 또는 2부 배정인 사람(qi < totalSize)만 가능
    const qi = activeQueueIndex(name);
    return qi < totalSize;
  }

  // 조출 가능 여부 (1부 6팀 이상 필수)
  const cho가능 = shift1Size >= 6;
  const cho현재수 = names.filter((n) => effectiveStatus(n) === "조출").length;
  const hu현재수 = names.filter((n) => effectiveStatus(n) === "후출").length;

  // 순번 위치에 따른 배정 구간 레이블 (livePreview 기반 정확한 계산)
  function getSlotLabel(name: string): { label: string; color: string } | null {
    if (!livePreview) return null;
    const cat = liveCategoryMap[name];
    if (!cat) return null;
    if (cat === "찾근")  return { label: "투라운드", color: "#00bcd4" };
    if (cat === "제외")  return null;
    if (cat === "1부") {
      const idx = livePreview.shift1.indexOf(name);
      const isLast = idx === livePreview.shift1.length - 1;
      return { label: `1부 #${idx + 1}${isLast ? " ★" : ""}`, color: "#1565c0" };
    }
    if (cat === "1부스페어") return { label: "1부스페어", color: "#e65100" };
    if (cat === "2부") {
      const idx = livePreview.shift2.indexOf(name);
      const isLast = idx === livePreview.shift2.length - 1;
      return { label: `2부 #${idx + 1}${isLast ? " ★" : ""}`, color: "#2e7d32" };
    }
    if (cat === "2부스페어") return { label: "2부스페어", color: "#6a1b9a" };
    if (cat === "단부") {
      const idx = livePreview.shift1.indexOf(name);
      const isLast = idx === livePreview.shift1.length - 1;
      return { label: `단부 #${idx + 1}${isLast ? " ★" : ""}`, color: "#1565c0" };
    }
    if (cat === "스페어") return { label: "스페어", color: "#6a1b9a" };
    return null;
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
        <img src={`${BASE}/char_dino.png`} alt="" style={{ width: 30, height: 30, objectFit: "contain" }} />
        <span style={S.headerTitle}>캐디 근무표</span>
        {view === "assign" && (
          <button onClick={() => { setView("input"); setDayResult(null); setWeekly([]); }} style={S.smallBtn}>
            ↩ 다시
          </button>
        )}
        {view === "input" && customRoster.length > 0 && (
          <button
            onClick={() => applyRoster(queueStartName)}
            style={{
              ...S.smallBtn,
              background: "#1a1a2e",
              color: "#fff",
              fontWeight: 800,
              fontSize: "1rem",
              padding: "6px 14px",
              letterSpacing: 1,
            }}
            title="배정 화면 바로 가기"
          >
            »
          </button>
        )}
      </div>

      {/* ─── 입력 단계 ─── */}
      {view === "input" && (
        <div style={S.card}>
          {/* 운영 모드 */}
          <label style={S.label}>운영 방식</label>
          {/* iOS 세그먼트 컨트롤 */}
          <div style={S.segmentTrack}>
            {(["2부제", "단부제"] as Mode[]).map((m) => (
              <button key={m}
                onClick={() => { setMode(m); if (teamsLocked) setTeamsLocked(false); }}
                style={{
                  ...S.segmentBtn,
                  background: mode === m
                    ? "linear-gradient(135deg, #1a1a2e 0%, #16213e 100%)"
                    : "transparent",
                  color: mode === m ? "#fff" : "#6b7280",
                  boxShadow: mode === m ? "0 2px 8px rgba(26,26,46,0.3)" : "none",
                }}>
                {m === "2부제" ? "☀️ 2부제" : "🌙 단부제"}
              </button>
            ))}
          </div>

          {/* ── 엑셀 날짜 선택 ── */}
          {/* ★ 기능1: 엑셀 파일 업로드 */}
          <div style={{ display: "flex", alignItems: "center", gap: "8px", marginBottom: "6px" }}>
            <label style={{ ...S.label, margin: 0, flex: 1 }}>
              {selectedDate ? selectedDate.dateLabel : "날짜 선택"}
              {xlLoading && <span style={{ color: "#aaa", fontWeight: 400, marginLeft: "6px" }}>불러오는 중…</span>}
              {xlError && <span style={{ color: "#e53935", fontWeight: 400, marginLeft: "6px" }}>{xlError}</span>}
            </label>
            <label style={{
              padding: "4px 10px", borderRadius: "8px", fontSize: "0.72rem", fontWeight: 700,
              background: uploadedName ? "#e8f5e9" : "#f0f0f0",
              color: uploadedName ? "#2e7d32" : "#555",
              border: uploadedName ? "1px solid #81c784" : "1px solid #ddd",
              cursor: "pointer", whiteSpace: "nowrap",
            }}>
              {uploadedName ? "📄 " + uploadedName.replace(/\.xlsx?$/i, "") : "📂 엑셀 교체"}
              <input
                type="file"
                accept=".xlsx,.xls"
                style={{ display: "none" }}
                onChange={e => {
                  const f = e.target.files?.[0];
                  if (f) loadFromFile(f);
                  e.target.value = "";
                }}
              />
            </label>
          </div>

          {/* 휴무 엑셀 자동입력 업로드 */}
          <div style={{ display: "flex", alignItems: "center", gap: "8px", marginBottom: "8px" }}>
            <div style={{ flex: 1, fontSize: "0.72rem", color: "#555" }}>
              {holidayFileName ? (
                <span>
                  📂 <strong style={{ color: "#2e7d32" }}>{holidayFileName.replace(/\.xlsx?$/i, "")}</strong>
                  {selectedDate && holidayMap[selectedDate.dateLabel.slice(0, 5)] ? (
                    <span style={{ color: "#1565c0", marginLeft: 6 }}>
                      ({holidayMap[selectedDate.dateLabel.slice(0, 5)].length}명 자동입력)
                    </span>
                  ) : selectedDate ? (
                    <span style={{ color: "#bbb", marginLeft: 6 }}>(해당 날짜 데이터 없음)</span>
                  ) : null}
                </span>
              ) : (
                <span style={{ color: "#aaa" }}>휴무 엑셀 미업로드</span>
              )}
            </div>
            <label style={{
              padding: "4px 10px", borderRadius: "8px", fontSize: "0.72rem", fontWeight: 700,
              background: holidayFileName ? "#e8f5e9" : "#fff3e0",
              color: holidayFileName ? "#2e7d32" : "#e65100",
              border: holidayFileName ? "1px solid #81c784" : "1px solid #ffb74d",
              cursor: "pointer", whiteSpace: "nowrap",
            }}>
              📂 {holidayFileName ? "휴무 교체" : "휴무 업로드"}
              <input
                type="file"
                accept=".xlsx,.xls"
                style={{ display: "none" }}
                onChange={e => {
                  const f = e.target.files?.[0];
                  if (f) loadHolidayFile(f);
                  e.target.value = "";
                }}
              />
            </label>
          </div>

          {/* ── 캐릭터 + 월 네비게이션 (엑셀 유무와 무관하게 항상 표시) ── */}
          <>
            <div style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "8px" }}>
              <img
                src={`${BASE}/char_cloud.png`}
                alt="cloud"
                style={{ width: 44, height: 44, objectFit: "contain", animation: "floatBob 3s ease-in-out infinite" }}
              />
              <div style={{ flex: 1 }}>
                {todayFirstHint && (
                  <div style={{ fontSize: "0.72rem", color: "#92400e", background: "#fef3c7", borderRadius: "6px", padding: "3px 8px", marginBottom: "4px" }}>
                    🔢 오늘 첫번호: <strong>{todayFirstHint}</strong>
                  </div>
                )}
                {excelDays.length === 0 && (
                  <div style={{ fontSize: "0.7rem", color: "#6b7280" }}>
                    엑셀 없이 달력 이동 가능 · 날짜를 선택해 배정하세요
                  </div>
                )}
              </div>
            </div>
            {/* 월 네비게이션 바 */}
            <div style={{
              display: "flex", alignItems: "center", justifyContent: "space-between",
              marginBottom: "8px",
              background: "linear-gradient(135deg, #1a1a2e 0%, #16213e 100%)",
              borderRadius: "12px", padding: "8px 12px",
              boxShadow: "0 2px 8px rgba(26,26,46,0.25)",
            }}>
              <button
                onClick={() => prevMonthStr && setViewMonth(prevMonthStr)}
                disabled={!prevMonthStr}
                style={{
                  background: "none", border: "none",
                  color: prevMonthStr ? "#fff" : "rgba(255,255,255,0.25)",
                  fontSize: "1.1rem", cursor: prevMonthStr ? "pointer" : "default",
                  padding: "2px 8px", minHeight: "36px",
                }}
              >‹</button>
              <span style={{ color: "#fff", fontWeight: 700, fontSize: "0.95rem" }}>
                {parseInt(viewMonth, 10)}월
                <span style={{ fontWeight: 400, fontSize: "0.75rem", marginLeft: "6px", opacity: 0.6 }}>
                  ({viewDays.length}일)
                </span>
              </span>
              <button
                onClick={() => nextMonthStr && setViewMonth(nextMonthStr)}
                disabled={!nextMonthStr}
                style={{
                  background: "none", border: "none",
                  color: nextMonthStr ? "#fff" : "rgba(255,255,255,0.25)",
                  fontSize: "1.1rem", cursor: nextMonthStr ? "pointer" : "default",
                  padding: "2px 8px", minHeight: "36px",
                }}
              >›</button>
            </div>

            {/* ── 날짜 그리드 ── */}
            <div style={S.dateGrid}>
              {viewDays.map((d, idx) => {
                const isSelected = selectedDate?.dateLabel === d.dateLabel;
                const isWeekend = d.dayIdx === 5 || d.dayIdx === 6;
                const hasTeams = d.예약팀수 > 0;
                const savedStatuses = dateStatuses[d.dateLabel] ?? {};
                const hasManual = Object.keys(savedStatuses).length > 0;
                const hasExcel = d.가용인원 > 0 || hasTeams;

                // 이 날의 "첫대기" = 전날 배정의 spare2[0]
                // viewDays에 없으면 excelDays에서 전날 탐색
                const prevDayLabel: string | null = (() => {
                  if (idx > 0) return viewDays[idx - 1].dateLabel;
                  // 첫날: excelDays에서 이 날 바로 앞 날짜 탐색
                  const ei = excelDays.findIndex(e => e.dateLabel === d.dateLabel);
                  return ei > 0 ? excelDays[ei - 1].dateLabel : null;
                })();
                const nextFirstHint = prevDayLabel ? (savedSpare2[prevDayLabel]?.[0] ?? null) : null;

                return (
                  <button
                    key={d.dateLabel}
                    onClick={() => selectExcelDate(d)}
                    style={{
                      ...S.dateBtn,
                      background: isSelected
                        ? "linear-gradient(135deg, #1a1a2e 0%, #4e89ae 100%)"
                        : hasManual ? "#eff6ff" : "#f8fafc",
                      color: isSelected ? "#fff" : isWeekend ? "#c62828" : "#1a1a2e",
                      border: isSelected
                        ? "2px solid #4e89ae"
                        : hasManual ? "2px solid #93c5fd"
                        : hasTeams ? "2px solid #60a5fa" : "1.5px solid #e5e7eb",
                      animation: isSelected ? "glowPulse 2s ease-in-out infinite" : "none",
                      transform: isSelected ? "scale(1.05)" : "scale(1)",
                      opacity: !hasExcel && !hasManual && !isSelected ? 0.75 : 1,
                      minHeight: nextFirstHint ? "64px" : "52px",
                    }}
                  >
                    <span style={{ fontSize: "0.72rem", fontWeight: 700 }}>{d.dateLabel.split(" ")[0]}</span>
                    <span style={{ fontSize: "0.65rem", opacity: 0.7 }}>{d.dayName}</span>
                    {d.가용인원 > 0 && (
                      <span style={{
                        fontSize: "0.6rem", fontWeight: 700,
                        color: isSelected ? "#fff" : "#2e7d32",
                        lineHeight: 1,
                      }}>
                        {d.가용인원}명
                      </span>
                    )}
                    {hasTeams && (
                      <span style={{
                        fontSize: "0.55rem", fontWeight: 700,
                        color: isSelected ? "rgba(255,255,255,0.75)" : "#1565c0",
                        lineHeight: 1,
                      }}>
                        {d.예약팀수}팀
                      </span>
                    )}
                    {/* 전날 spare2[0] → 이 날의 첫대기 힌트 */}
                    {nextFirstHint && (
                      <span style={{
                        fontSize: "0.52rem", fontWeight: 800, lineHeight: 1,
                        color: isSelected ? "rgba(255,255,255,0.9)" : "#b91c1c",
                        background: isSelected ? "rgba(255,255,255,0.15)" : "#fef2f2",
                        borderRadius: 4, padding: "1px 4px",
                        marginTop: 1,
                        maxWidth: "100%", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
                      }}>
                        ↑{nextFirstHint}
                      </span>
                    )}
                  </button>
                );
              })}
            </div>
          </>

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
              {/* 가용인원 · 예약팀수 · 대근 요약 */}
              <div style={{ ...S.excelStatRow, alignItems: "center" }}>
                <StatBadge label="가용인원" value={selectedDate.가용인원} color="#1565c0" />
                <StatBadge label="예약팀수" value={selectedDate.예약팀수 || "미입력"} color={selectedDate.예약팀수 > 0 ? "#2e7d32" : "#9e9e9e"} />
                {(() => {
                  const daegeunBaseNames = names.length > 0 ? names : sortedCustomRoster.map(p => p.name);
                  const daegeunCandidates = daegeunBaseNames.filter(n => {
                    const p = customRosterMap[n];
                    return p != null && isAutoOff(p.group, dayOfWeek) && !(n in manualStatuses);
                  });
                  const activeDaegeun = daegeunCandidates.filter(n => currentDaegeun[n]);
                  return (
                    <button
                      onClick={() => { setBatchDaegeunOpen(true); setBatchDaegeunSearch(""); }}
                      style={{
                        display: "flex", flexDirection: "column",
                        alignItems: "center", justifyContent: "center",
                        padding: "6px 10px",
                        borderRadius: "8px",
                        border: `1px solid ${activeDaegeun.length > 0 ? "#f59e0b" : "#f59e0b33"}`,
                        background: activeDaegeun.length > 0 ? "#fef3c7" : "#f59e0b15",
                        cursor: "pointer",
                        minWidth: "52px",
                        fontFamily: "inherit",
                        lineHeight: 1,
                        outline: "none",
                        WebkitTapHighlightColor: "transparent",
                      }}>
                      <span style={{ fontSize: "0.65rem", fontWeight: 600, color: "#92400e" }}>대근</span>
                      <span style={{ fontSize: "1rem", fontWeight: 700, color: activeDaegeun.length > 0 ? "#f59e0b" : "#d1d5db" }}>
                        {activeDaegeun.length > 0 ? activeDaegeun.length : daegeunCandidates.length}
                      </span>
                    </button>
                  );
                })()}
                {/* 병가 버튼 (대근 버튼 옆) */}
                {(() => {
                  const baseNames = names.length > 0 ? names : sortedCustomRoster.map(p => p.name);
                  const sickCnt = baseNames.filter(n => effectiveStatus(n, dayOfWeek) === "병가").length;
                  const active = sickCnt > 0;
                  return (
                    <button
                      onClick={() => openStatusPicker("병가")}
                      style={{
                        display: "flex", flexDirection: "column",
                        alignItems: "center", justifyContent: "center",
                        padding: "6px 10px",
                        borderRadius: "8px",
                        border: `1px solid ${active ? "#c62828" : "#c6282833"}`,
                        background: active ? "#ffebee" : "#c6282815",
                        cursor: "pointer",
                        minWidth: "52px",
                        fontFamily: "inherit",
                        lineHeight: 1,
                        outline: "none",
                        WebkitTapHighlightColor: "transparent",
                      }}>
                      <span style={{ fontSize: "0.65rem", fontWeight: 600, color: "#b71c1c" }}>병가</span>
                      <span style={{ fontSize: "1rem", fontWeight: 700, color: active ? "#c62828" : "#d1d5db" }}>
                        {active ? sickCnt : "–"}
                      </span>
                    </button>
                  );
                })()}
              </div>

              {/* ── 6개 상태 선택 버튼 ── */}
              <div style={{
                display: "grid", gridTemplateColumns: "repeat(3, 1fr)",
                gap: "7px", marginTop: "12px",
              }}>
                {STATUS_BTNS.map(({ st, label, color, bg }) => {
                  const countNames = names.length > 0 ? names : sortedCustomRoster.map(p => p.name);
                  const cnt = countNames.filter(n => effectiveStatus(n, dayOfWeek) === st).length;
                  const active = cnt > 0;
                  return (
                    <button key={st}
                      onClick={() => openStatusPicker(st)}
                      style={{
                        display: "flex", flexDirection: "column",
                        alignItems: "center", justifyContent: "center",
                        gap: "2px", padding: "10px 4px",
                        borderRadius: "12px", border: `2px solid ${active ? color : color + "44"}`,
                        background: active ? bg : "#fafafa",
                        cursor: "pointer", position: "relative",
                        WebkitTapHighlightColor: "transparent",
                      }}>
                      <span style={{ fontSize: "0.85rem", fontWeight: 800, color }}>
                        {label}
                      </span>
                      <span style={{
                        fontSize: "1.1rem", fontWeight: 900,
                        color: active ? color : "#ccc", lineHeight: 1,
                      }}>
                        {active ? cnt : "–"}
                      </span>
                      {active && (
                        <span style={{
                          position: "absolute", top: "3px", right: "5px",
                          fontSize: "0.48rem", color: color + "99",
                        }}>●</span>
                      )}
                    </button>
                  );
                })}
              </div>
              {selectedDate.예약팀수 === 0 && (
                <div style={{ fontSize: "0.72rem", color: "#ff8f00", marginTop: "4px" }}>
                  ⚠ 예약팀수 미입력 — 아래에서 직접 팀수를 입력해 주세요
                </div>
              )}

              {/* ── VIP 섹션 ── */}
              <div style={{ marginTop: "12px", borderTop: "1px solid #e8e0f0", paddingTop: "10px" }}>
                {/* VIP 헤더 행 */}
                <div style={{ display: "flex", alignItems: "center", gap: "8px", marginBottom: "8px" }}>
                  <span style={{ fontWeight: 800, fontSize: "0.85rem", color: "#7b1fa2" }}>
                    👑 VIP
                  </span>
                  {currentVip.count > 0 && (
                    <span style={{
                      padding: "2px 8px", borderRadius: "8px",
                      background: "#f3e5f5", color: "#7b1fa2",
                      fontSize: "0.78rem", fontWeight: 700,
                    }}>
                      {currentVip.count}팀 {currentVip.round ? `(${currentVip.round})` : ""}
                    </span>
                  )}
                  {currentVip.members && currentVip.members.length > 0 && (
                    <span style={{ fontSize: "0.75rem", color: "#9c27b0", fontWeight: 600 }}>
                      · {currentVip.members.join(", ")}
                    </span>
                  )}
                  <button
                    onClick={() => setVipPickerOpen(v => !v)}
                    style={{
                      marginLeft: "auto", padding: "4px 12px", borderRadius: "10px",
                      border: `1.5px solid ${vipPickerOpen ? "#7b1fa2" : "#ce93d8"}`,
                      background: vipPickerOpen ? "#7b1fa2" : "#fce4ec",
                      color: vipPickerOpen ? "#fff" : "#7b1fa2",
                      fontWeight: 700, fontSize: "0.78rem", cursor: "pointer",
                      flexShrink: 0,
                    }}>
                    {vipPickerOpen ? "닫기" : (currentVip.count > 0 ? "수정" : "+ 추가")}
                  </button>
                </div>

                {/* VIP 입력 패널 */}
                {vipPickerOpen && (
                  <div style={{
                    background: "#fdf5ff", borderRadius: "12px",
                    padding: "12px", border: "1.5px solid #ce93d8",
                  }}>
                    {/* 팀수 입력 */}
                    <div style={{ marginBottom: "10px" }}>
                      <label style={{ fontSize: "0.78rem", fontWeight: 700, color: "#555", display: "block", marginBottom: "5px" }}>
                        VIP 팀수
                      </label>
                      <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                        <button
                          onClick={() => setCurrentVip({ count: Math.max(0, currentVip.count - 1) })}
                          style={{
                            width: "36px", height: "36px", borderRadius: "10px",
                            border: "1.5px solid #ce93d8", background: "#fff",
                            color: "#7b1fa2", fontSize: "1.2rem", fontWeight: 700, cursor: "pointer",
                          }}>−</button>
                        <span style={{ fontSize: "1.4rem", fontWeight: 900, color: "#7b1fa2", minWidth: "32px", textAlign: "center" }}>
                          {currentVip.count}
                        </span>
                        <button
                          onClick={() => setCurrentVip({ count: currentVip.count + 1 })}
                          style={{
                            width: "36px", height: "36px", borderRadius: "10px",
                            border: "1.5px solid #ce93d8", background: "#fff",
                            color: "#7b1fa2", fontSize: "1.2rem", fontWeight: 700, cursor: "pointer",
                          }}>＋</button>
                        {(currentVip.count > 0 || (currentVip.members && currentVip.members.length > 0)) && (
                          <button
                            onClick={() => { setCurrentVip({ count: 0, round: null, members: [] }); setVipPickerOpen(false); }}
                            style={{
                              marginLeft: "auto", padding: "4px 10px", borderRadius: "8px",
                              border: "1px solid #ef9a9a", background: "#ffebee",
                              color: "#c62828", fontSize: "0.72rem", fontWeight: 700, cursor: "pointer",
                            }}>삭제</button>
                        )}
                      </div>
                    </div>

                    {/* 1부 / 2부 선택 */}
                    {currentVip.count > 0 && (
                      <div style={{ marginBottom: "10px" }}>
                        <label style={{ fontSize: "0.78rem", fontWeight: 700, color: "#555", display: "block", marginBottom: "5px" }}>
                          나가는 부
                        </label>
                        <div style={{ display: "flex", gap: "8px" }}>
                          {(["1부", "2부"] as VipRound[]).filter(Boolean).map(r => (
                            <button
                              key={r!}
                              onClick={() => setCurrentVip({ round: currentVip.round === r ? null : r })}
                              style={{
                                flex: 1, padding: "10px", borderRadius: "12px",
                                border: `2px solid ${currentVip.round === r ? "#7b1fa2" : "#ce93d8"}`,
                                background: currentVip.round === r ? "#7b1fa2" : "#fff",
                                color: currentVip.round === r ? "#fff" : "#7b1fa2",
                                fontWeight: 800, fontSize: "0.9rem", cursor: "pointer",
                              }}>
                              {r}
                            </button>
                          ))}
                        </div>
                        {currentVip.round && (
                          <div style={{ marginTop: "8px", fontSize: "0.75rem", color: "#7b1fa2", fontWeight: 600 }}>
                            ✅ VIP {currentVip.count}팀 → {currentVip.round} 배정
                          </div>
                        )}
                      </div>
                    )}

                    {/* VIP 인원 선택 */}
                    <div>
                      <div style={{ display: "flex", alignItems: "center", marginBottom: "6px" }}>
                        <label style={{ fontSize: "0.78rem", fontWeight: 700, color: "#555", flex: 1 }}>
                          VIP 담당 인원
                        </label>
                        <button
                          onClick={() => { setVipMemberSearch(""); setVipMemberPickerOpen(true); }}
                          style={{
                            padding: "4px 10px", borderRadius: "8px", border: "none",
                            background: "#ce93d8", color: "#fff",
                            fontWeight: 700, fontSize: "0.72rem", cursor: "pointer",
                          }}>
                          + 인원 선택
                        </button>
                      </div>

                      {/* 선택된 인원 칩 */}
                      {currentVip.members && currentVip.members.length > 0 ? (
                        <div style={{ display: "flex", flexWrap: "wrap", gap: "5px" }}>
                          {currentVip.members.map(name => (
                            <span
                              key={name}
                              style={{
                                display: "inline-flex", alignItems: "center", gap: "4px",
                                padding: "4px 8px", borderRadius: "20px",
                                background: "#7b1fa2", color: "#fff",
                                fontSize: "0.76rem", fontWeight: 700,
                              }}
                            >
                              {name}
                              <button
                                onClick={() => setCurrentVip({ members: currentVip.members.filter(n => n !== name) })}
                                style={{
                                  background: "none", border: "none", color: "#e1bee7",
                                  cursor: "pointer", fontSize: "0.85rem", lineHeight: 1, padding: 0,
                                }}
                              >×</button>
                            </span>
                          ))}
                        </div>
                      ) : (
                        <div style={{ fontSize: "0.75rem", color: "#bbb", fontStyle: "italic" }}>
                          선택된 인원 없음
                        </div>
                      )}
                    </div>
                  </div>
                )}
              </div>

            </div>
          )}

          {/* 팀수 입력 / 저장 */}
          {teamsLocked ? (
            /* ── 저장된 팀수 요약 카드 ── */
            <div style={{
              marginTop: "14px", marginBottom: "14px",
              background: "#f0f7ff", borderRadius: "12px",
              padding: "12px 14px", border: "1.5px solid #90caf9",
            }}>
              <div style={{ display: "flex", alignItems: "center", marginBottom: "8px" }}>
                <span style={{ fontWeight: 800, fontSize: "0.85rem", color: "#1565c0" }}>
                  ✅ 팀수 저장됨 ({mode})
                </span>
                <button
                  onClick={unlockTeamSettings}
                  style={{
                    marginLeft: "auto", padding: "4px 12px", borderRadius: "8px",
                    border: "1.5px solid #1565c0", background: "#fff",
                    color: "#1565c0", fontWeight: 700, fontSize: "0.78rem", cursor: "pointer",
                  }}>
                  ✏ 수정
                </button>
              </div>
              {mode === "2부제" ? (
                <div style={S.calcBox}>
                  <span style={{ color: "#1565c0" }}>1부 {shift1Size}팀</span>
                  <span style={{ color: "#aaa" }}>+</span>
                  <span style={{ color: "#2e7d32" }}>2부 {shift2Size}팀</span>
                  <span style={{ color: "#aaa" }}>=</span>
                  <span style={{ fontWeight: 700 }}>총 {totalSize}팀</span>
                  <span style={{ color: "#aaa", margin: "0 4px" }}>│</span>
                  <span style={{ color: "#e65100", fontSize: "0.75rem" }}>1부스페어: {shift1Size + 1}번째</span>
                </div>
              ) : (
                <div style={{ fontSize: "0.88rem", fontWeight: 700, color: "#333" }}>
                  팀수 {singleSize}팀
                </div>
              )}
            </div>
          ) : (
            /* ── 팀수 입력 폼 ── */
            <div style={{ marginBottom: "4px", marginTop: "14px" }}>
              {mode === "2부제" ? (
                <>
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
                </>
              ) : (
                <div style={{ marginBottom: "8px" }}>
                  <label style={S.label}>팀수</label>
                  <input type="number" value={singleSize} min={1}
                    onChange={(e) => setSingleSize(Number(e.target.value))} style={S.numInput} />
                </div>
              )}
              <button
                onClick={saveTeamSettings}
                style={{
                  width: "100%", marginTop: "10px", padding: "11px",
                  borderRadius: "12px", border: "none",
                  background: "#2e7d32", color: "#fff",
                  fontWeight: 800, fontSize: "0.9rem", cursor: "pointer",
                }}>
                💾 저장하기
              </button>
            </div>
          )}

          {/* 순번표 불러오기 + 편집 */}
          <div style={{ display: "flex", gap: "8px", marginBottom: "8px" }}>
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

          {/* 첫번호 / 다음날 첫 순번 표시 */}
          {(todayFirstHint || queueStartName) && (() => {
            const isAuto = !!todayFirstHint;
            const displayName = todayFirstHint ?? queueStartName!;
            return (
              <div style={{
                display: "flex", alignItems: "center", gap: 8, marginBottom: "12px",
                padding: "9px 12px", borderRadius: 10,
                background: isAuto ? "#f3e5f5" : "#e3f2fd",
                border: isAuto ? "1px solid #ce93d8" : "1px solid #90caf9",
              }}>
                <span style={{ fontSize: 14, flex: 1, color: isAuto ? "#6a1b9a" : "#1565c0" }}>
                  <span style={{ fontWeight: 700 }}>
                    {isAuto ? "🔢 다음날 첫 순번 " : "📌 첫번호 고정 "}
                  </span>
                  <span style={{ fontWeight: 800 }}>"{displayName}"</span>
                </span>
                <button
                  onClick={() => { setQueueModalPos({ x: 0, y: 0 }); setQueueModal("ask"); }}
                  style={{
                    padding: "6px 14px", borderRadius: 8, border: "none",
                    background: isAuto ? "#6a1b9a" : "#1565c0",
                    color: "#fff", fontWeight: 700, fontSize: 12, cursor: "pointer", whiteSpace: "nowrap",
                  }}
                >변경</button>
              </div>
            );
          })()}

        </div>
      )}

      {/* ── 첫번호: 이 순번대로 가시겠습니까? ── */}
      {queueModal === "ask" && (() => {
        // 가장 최근 저장된 2부스페어 첫번째 찾기
        const spare2First: string | null = (() => {
          // 선택된 날짜 전날 스페어 우선
          if (todayFirstHint) return todayFirstHint;
          // 없으면 savedSpare2에서 가장 최근 날짜의 첫번째
          const entries = Object.entries(savedSpare2);
          if (!entries.length) return null;
          entries.sort((a, b) => a[0].localeCompare(b[0]));
          return entries[entries.length - 1][1][0] ?? null;
        })();

        return (
          <div style={{
            position: "fixed", inset: 0, zIndex: 400,
            background: "rgba(0,0,0,0.65)",
            display: "flex", alignItems: "center", justifyContent: "center",
            userSelect: "none",
          }}>
            <div style={{
              background: "#fff", borderRadius: 20, padding: "0 0 22px",
              maxWidth: 330, width: "90%", textAlign: "center",
              boxShadow: "0 8px 40px rgba(0,0,0,0.28)",
              transform: `translate(${queueModalPos.x}px, ${queueModalPos.y}px)`,
              position: "relative",
              cursor: "default",
            }}>
              {/* 드래그 핸들 + X 버튼 영역 */}
              <div
                onMouseDown={onModalDragStart}
                onTouchStart={onModalDragStart}
                style={{
                  display: "flex", alignItems: "center", justifyContent: "space-between",
                  padding: "12px 16px 0",
                  cursor: "grab",
                  borderRadius: "20px 20px 0 0",
                }}
              >
                <div style={{ fontSize: "0.7rem", color: "#bbb", letterSpacing: 1 }}>☰ 드래그</div>
                <button
                  onMouseDown={e => e.stopPropagation()}
                  onClick={() => { setQueueModal(null); setQueueModalPos({ x: 0, y: 0 }); }}
                  style={{
                    background: "#f3f4f6", border: "none", borderRadius: "50%",
                    width: 28, height: 28, display: "flex", alignItems: "center", justifyContent: "center",
                    cursor: "pointer", fontSize: 14, color: "#6b7280", fontWeight: 700,
                  }}
                >✕</button>
              </div>
              <div style={{ padding: "8px 24px 0" }}>
              <div style={{ fontSize: 38, marginBottom: 10 }}>📋</div>
              <div style={{ fontWeight: 800, fontSize: 18, marginBottom: 16, color: "#1a1a2e" }}>
                첫번호부터 시작하겠습니까?
              </div>

              {/* 하겠습니다 → 첫번호 직접 선택 */}
              <button
                onClick={() => {
                  setQueuePickSearch("");
                  setQueueModal("pick");
                }}
                style={{
                  width: "100%", padding: "14px 0", borderRadius: 12, border: "none",
                  background: "#1565c0", color: "#fff",
                  fontWeight: 700, fontSize: 15, cursor: "pointer",
                  marginBottom: 10,
                  display: "flex", alignItems: "center", justifyContent: "center", gap: 8,
                }}
              >
                <span>✅ 하겠습니다</span>
                <span style={{
                  background: "rgba(255,255,255,0.2)", borderRadius: 8,
                  padding: "2px 10px", fontSize: 14, fontWeight: 800,
                }}>
                  첫번호 선택 →
                </span>
              </button>

              {/* 안하겠습니다 */}
              <button
                onClick={() => {
                  if (spare2First) {
                    applyRoster(spare2First);
                  } else {
                    // 저장된 스페어 없으면 첫번호 그대로
                    applyRoster(queueStartName);
                  }
                  setQueueModal(null);
                }}
                style={{
                  width: "100%", padding: "14px 0", borderRadius: 12,
                  border: "2px solid #f8b400",
                  background: spare2First ? "#fffbeb" : "#f5f5f5",
                  color: spare2First ? "#92400e" : "#999",
                  fontWeight: 700, fontSize: 15, cursor: "pointer",
                  marginBottom: 14,
                  display: "flex", alignItems: "center", justifyContent: "center", gap: 8,
                }}
              >
                <span>❌ 안하겠습니다</span>
                {spare2First ? (
                  <span style={{
                    background: "#fef3c7", border: "1px solid #fcd34d",
                    borderRadius: 8, padding: "2px 10px", fontSize: 14, fontWeight: 800, color: "#b45309",
                  }}>
                    🏁 {spare2First}부터
                  </span>
                ) : (
                  <span style={{ fontSize: 12, color: "#bbb" }}>(저장된 스페어 없음)</span>
                )}
              </button>

              {/* 구분선 */}
              <div style={{ borderTop: "1px solid #eee", paddingTop: 12 }}>
                <button
                  onClick={() => { applyRoster(null); saveQueueStart(null); setQueueModal(null); }}
                  style={{
                    width: "100%", padding: "8px 0", borderRadius: 8, border: "1px solid #ddd",
                    background: "transparent", color: "#999", fontSize: 13, cursor: "pointer",
                  }}
                >
                  첫번호 없이
                </button>
              </div>
              </div>{/* /padding wrapper */}
            </div>
          </div>
        );
      })()}

      {/* ── 첫번호 지정 picker ── */}
      {queueModal === "pick" && (
        <div style={{
          position: "fixed", inset: 0, zIndex: 400,
          background: "rgba(0,0,0,0.55)",
          display: "flex", flexDirection: "column", justifyContent: "flex-end",
        }}>
          <div style={{
            background: "#fff", borderRadius: "18px 18px 0 0",
            maxHeight: "75vh", display: "flex", flexDirection: "column",
            boxShadow: "0 -4px 24px rgba(0,0,0,0.18)",
          }}>
            {/* 헤더 */}
            <div style={{
              padding: "18px 18px 10px", borderBottom: "1px solid #eee",
              display: "flex", alignItems: "center", justifyContent: "space-between",
            }}>
              <div>
                <div style={{ fontWeight: 700, fontSize: 16 }}>🔢 첫번호 지정</div>
                <div style={{ fontSize: 12, color: "#777", marginTop: 2 }}>
                  이 사람부터 순번이 시작됩니다
                </div>
              </div>
              <button
                onClick={() => setQueueModal(null)}
                style={{ background: "none", border: "none", fontSize: 22, cursor: "pointer", color: "#888" }}
              >×</button>
            </div>
            {/* ── 빠른 선택: 현재 저장된 첫번호 ── */}
            {queueStartName && sortedCustomRoster.some(p => p.name === queueStartName) && (
              <div style={{ padding: "8px 14px", borderBottom: "1px solid #eee" }}>
                <div style={{ fontSize: 11, color: "#888", marginBottom: 6 }}>⚡ 빠른 선택 (이전 첫번호)</div>
                <button
                  onTouchEnd={(e) => {
                    e.preventDefault();
                    saveQueueStart(queueStartName);
                    applyRoster(queueStartName);
                    setQueueModal(null);
                  }}
                  onClick={() => {
                    saveQueueStart(queueStartName);
                    applyRoster(queueStartName);
                    setQueueModal(null);
                  }}
                  style={{
                    width: "100%", padding: "12px 16px", borderRadius: 12,
                    border: "2px solid #1565c0", background: "#e3f2fd",
                    display: "flex", alignItems: "center", gap: 10, cursor: "pointer",
                    touchAction: "manipulation",
                  }}
                >
                  {(() => {
                    const person = sortedCustomRoster.find(p => p.name === queueStartName);
                    return <>
                      <span style={{
                        background: "#1565c0", borderRadius: 8, padding: "2px 8px",
                        fontSize: 11, color: "#fff", fontWeight: 700, minWidth: 44,
                      }}>
                        {person?.조}조 {person?.no}번
                      </span>
                      <span style={{ fontWeight: 800, fontSize: 16, color: "#1565c0" }}>{queueStartName}</span>
                      <span style={{ marginLeft: "auto", fontSize: 12, color: "#1565c0", fontWeight: 700 }}>✓ 선택 →</span>
                    </>;
                  })()}
                </button>
              </div>
            )}

            {/* 검색 */}
            <div style={{ padding: "10px 14px" }}>
              <input
                value={queuePickSearch}
                onChange={e => setQueuePickSearch(e.target.value)}
                placeholder="이름 검색..."
                style={{
                  width: "100%", padding: "9px 12px", borderRadius: 10,
                  border: "1px solid #ddd", fontSize: 14, boxSizing: "border-box",
                }}
              />
            </div>
            {/* 목록 */}
            <div style={{ overflowY: "auto", flex: 1, paddingBottom: 16 }}>
              {sortedCustomRoster
                .filter(p => !queuePickSearch || p.name.includes(queuePickSearch))
                .map(p => (
                  <button
                    key={p.name}
                    onTouchEnd={(e) => {
                      e.preventDefault();
                      saveQueueStart(p.name);
                      applyRoster(p.name);
                      setQueueModal(null);
                    }}
                    onClick={() => {
                      saveQueueStart(p.name);
                      applyRoster(p.name);
                      setQueueModal(null);
                    }}
                    style={{
                      display: "flex", width: "100%", padding: "12px 18px",
                      alignItems: "center", gap: 10, border: "none",
                      background: p.name === queueStartName ? "#e3f2fd" : "transparent",
                      cursor: "pointer", textAlign: "left",
                      touchAction: "manipulation",
                    }}
                  >
                    <span style={{
                      background: "#e8eaf6", borderRadius: 8, padding: "2px 8px",
                      fontSize: 11, color: "#5c6bc0", fontWeight: 700, minWidth: 44,
                    }}>
                      {p.조}조 {p.no}번
                    </span>
                    <span style={{ fontWeight: 600, fontSize: 15 }}>{p.name}</span>
                    <span style={{
                      fontSize: 11, color: "#888", marginLeft: "auto",
                      background: "#f5f5f5", borderRadius: 6, padding: "2px 8px",
                    }}>
                      {p.group}
                    </span>
                    {p.name === queueStartName && (
                      <span style={{ color: "#1565c0", fontSize: 13, fontWeight: 700 }}>✓ 현재</span>
                    )}
                  </button>
                ))}
            </div>
          </div>
        </div>
      )}

      {/* ── VIP 인원 선택 모달 ── */}
      {vipMemberPickerOpen && (
        <div style={{
          position: "fixed", inset: 0, zIndex: 400,
          background: "rgba(0,0,0,0.55)",
          display: "flex", flexDirection: "column", justifyContent: "flex-end",
        }}>
          <div style={{
            background: "#fff", borderRadius: "18px 18px 0 0",
            maxHeight: "72vh", display: "flex", flexDirection: "column",
            boxShadow: "0 -4px 24px rgba(0,0,0,0.18)",
          }}>
            {/* 헤더 */}
            <div style={{
              padding: "16px 18px 10px", borderBottom: "1px solid #f3e5f5",
              display: "flex", alignItems: "center", justifyContent: "space-between",
            }}>
              <div>
                <div style={{ fontWeight: 700, fontSize: 16, color: "#7b1fa2" }}>👑 VIP 담당 인원 선택</div>
                <div style={{ fontSize: 12, color: "#9c27b0", marginTop: 2 }}>
                  탭하면 선택/해제됩니다 ({(currentVip.members ?? []).length}명 선택됨)
                </div>
              </div>
              <button
                onClick={() => setVipMemberPickerOpen(false)}
                style={{ background: "none", border: "none", fontSize: 22, cursor: "pointer", color: "#888" }}
              >×</button>
            </div>

            {/* 선택된 인원 미리보기 */}
            {(currentVip.members ?? []).length > 0 && (
              <div style={{ padding: "8px 14px", display: "flex", flexWrap: "wrap", gap: 5, borderBottom: "1px solid #f3e5f5" }}>
                {(currentVip.members ?? []).map(name => (
                  <span key={name} style={{
                    display: "inline-flex", alignItems: "center", gap: 4,
                    padding: "4px 8px", borderRadius: 20,
                    background: "#7b1fa2", color: "#fff",
                    fontSize: "0.76rem", fontWeight: 700,
                  }}>
                    {name}
                    <button
                      onClick={() => setCurrentVip({ members: (currentVip.members ?? []).filter(n => n !== name) })}
                      style={{ background: "none", border: "none", color: "#e1bee7", cursor: "pointer", fontSize: "0.85rem", padding: 0 }}
                    >×</button>
                  </span>
                ))}
                <button
                  onClick={() => setCurrentVip({ members: [] })}
                  style={{
                    padding: "4px 10px", borderRadius: 20, border: "none",
                    background: "#ffebee", color: "#c62828",
                    fontSize: "0.72rem", fontWeight: 700, cursor: "pointer",
                  }}
                >전체 해제</button>
              </div>
            )}

            {/* 검색 */}
            <div style={{ padding: "8px 14px" }}>
              <input
                value={vipMemberSearch}
                onChange={e => setVipMemberSearch(e.target.value)}
                placeholder="이름 검색..."
                style={{
                  width: "100%", padding: "9px 12px", borderRadius: 10,
                  border: "1.5px solid #ce93d8", fontSize: 14, boxSizing: "border-box",
                  outline: "none",
                }}
              />
            </div>

            {/* 인원 목록 */}
            <div style={{ overflowY: "auto", flex: 1, paddingBottom: 16 }}>
              {(() => {
                const roster = (names.length > 0 ? names : sortedCustomRoster.map(p => p.name));
                const filtered = vipMemberSearch
                  ? roster.filter(name => name.includes(vipMemberSearch))
                  : roster;
                const vipMembers = currentVip.members ?? [];
                return filtered.map(name => {
                  const person = customRosterMap[name];
                  const selected = vipMembers.includes(name);
                  return (
                    <button
                      key={name}
                      onClick={() => {
                        const cur = currentVip.members ?? [];
                        setCurrentVip({
                          members: selected ? cur.filter(n => n !== name) : [...cur, name],
                        });
                      }}
                      style={{
                        display: "flex", width: "100%", padding: "11px 18px",
                        alignItems: "center", gap: 10, border: "none",
                        background: selected ? "#f3e5f5" : "transparent",
                        cursor: "pointer", textAlign: "left",
                        borderLeft: selected ? "3px solid #7b1fa2" : "3px solid transparent",
                      }}
                    >
                      {person && (
                        <span style={{
                          background: selected ? "#7b1fa2" : "#e8eaf6",
                          borderRadius: 8, padding: "2px 8px",
                          fontSize: 11, color: selected ? "#fff" : "#5c6bc0",
                          fontWeight: 700, minWidth: 44,
                        }}>
                          {person.조}조 {person.no}번
                        </span>
                      )}
                      <span style={{ fontWeight: 600, fontSize: 15, flex: 1 }}>{name}</span>
                      {person && (
                        <span style={{
                          fontSize: 11, color: "#888",
                          background: "#f5f5f5", borderRadius: 6, padding: "2px 8px",
                        }}>
                          {person.group}
                        </span>
                      )}
                      {selected && (
                        <span style={{ color: "#7b1fa2", fontWeight: 800, fontSize: 16 }}>✓</span>
                      )}
                    </button>
                  );
                });
              })()}
            </div>

            {/* 확인 버튼 */}
            <div style={{ padding: "12px 16px", borderTop: "1px solid #f3e5f5" }}>
              <button
                onClick={() => setVipMemberPickerOpen(false)}
                style={{
                  width: "100%", padding: "13px", borderRadius: 12, border: "none",
                  background: "#7b1fa2", color: "#fff", fontWeight: 800, fontSize: 15, cursor: "pointer",
                }}
              >
                완료 ({(currentVip.members ?? []).length}명 선택)
              </button>
            </div>
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
              borderBottom: "1px solid #f0f0f0", background: "#1a1a2e", gap: "8px",
            }}>
              <span style={{ fontWeight: 800, fontSize: "1rem", color: "#fff" }}>
                📋 순번표 편집
              </span>
              <span style={{ fontSize: "0.78rem", color: "#aaa" }}>
                총 {customRoster.length}명
              </span>

              {/* 백업 / 복원 버튼 */}
              <input
                ref={rosterImportRef} type="file" accept=".json"
                style={{ display: "none" }} onChange={handleImportRoster}
              />
              <button onClick={exportRoster} style={{
                marginLeft: "auto", padding: "5px 10px", borderRadius: "8px",
                border: "none", background: "#2e7d32", color: "#fff",
                fontSize: "0.75rem", fontWeight: 700, cursor: "pointer",
              }}>💾 내보내기</button>
              <button onClick={() => rosterImportRef.current?.click()} style={{
                padding: "5px 10px", borderRadius: "8px",
                border: "none", background: "#1565c0", color: "#fff",
                fontSize: "0.75rem", fontWeight: 700, cursor: "pointer",
              }}>📂 복원</button>

              <button onClick={() => { setRosterEditorOpen(false); setRosterForm(null); }}
                style={{
                  background: "rgba(255,255,255,0.15)", border: "none",
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
                      <div key={`h-${joNum}-${name}`} style={{ padding: "5px 14px 3px", display: "flex", alignItems: "center", gap: "8px" }}>
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

      {/* ─── 대근 일괄 모달 ─── */}
      {batchDaegeunOpen && (() => {
        const batchBaseNames = names.length > 0 ? names : sortedCustomRoster.map(p => p.name);
        const candidates = batchBaseNames.filter(n => {
          const p = customRosterMap[n];
          return p != null && isAutoOff(p.group, dayOfWeek) && !(n in manualStatuses);
        });
        const filtered = batchDaegeunSearch.trim()
          ? candidates.filter(n => n.includes(batchDaegeunSearch.trim()))
          : candidates;
        const activeCnt = candidates.filter(n => currentDaegeun[n]).length;
        const groupLabel = dayOfWeek <= 3 ? "주말반" : dayOfWeek >= 5 ? "주중반" : "비근무";
        return (
          <div style={{
            position: "fixed", inset: 0, zIndex: 490,
            background: "rgba(0,0,0,0.55)",
            display: "flex", flexDirection: "column", justifyContent: "flex-end",
          }}
            onClick={(e) => { if (e.target === e.currentTarget) setBatchDaegeunOpen(false); }}
          >
            <div style={{
              background: "#fff", borderRadius: "20px 20px 0 0",
              maxHeight: "82vh", display: "flex", flexDirection: "column",
              overflow: "hidden",
            }}>
              {/* 헤더 */}
              <div style={{
                padding: "16px 16px 10px",
                background: "linear-gradient(135deg, #fef3c7 0%, #fde68a 100%)",
                display: "flex", alignItems: "center", justifyContent: "space-between",
                flexShrink: 0,
              }}>
                <div>
                  <span style={{ fontWeight: 800, fontSize: "1rem", color: "#78350f" }}>
                    대근 관리
                  </span>
                  <span style={{
                    marginLeft: "8px", fontSize: "0.75rem", color: "#92400e",
                    background: "#fef3c7", padding: "2px 8px", borderRadius: "8px",
                    fontWeight: 700,
                  }}>
                    {groupLabel} {candidates.length}명
                  </span>
                  {activeCnt > 0 && (
                    <span style={{
                      marginLeft: "6px", fontSize: "0.75rem", color: "#fff",
                      background: "#f59e0b", padding: "2px 8px", borderRadius: "8px",
                      fontWeight: 700,
                    }}>
                      투입 {activeCnt}명
                    </span>
                  )}
                </div>
                <button onClick={() => setBatchDaegeunOpen(false)}
                  style={{
                    background: "transparent", border: "none",
                    fontSize: "1.3rem", cursor: "pointer", color: "#92400e", lineHeight: 1,
                  }}>✕</button>
              </div>

              {/* 검색 */}
              <div style={{ padding: "10px 14px 6px", flexShrink: 0, borderBottom: "1px solid #f3f4f6" }}>
                <input
                  type="text"
                  placeholder="이름 검색..."
                  value={batchDaegeunSearch}
                  onChange={(e) => setBatchDaegeunSearch(e.target.value)}
                  style={{
                    width: "100%", padding: "9px 12px", borderRadius: "10px",
                    border: "1.5px solid #e5e7eb", fontSize: "0.9rem",
                    outline: "none", boxSizing: "border-box",
                  }}
                />
              </div>

              {/* 인원 목록 */}
              <div style={{ overflowY: "auto", flex: 1, padding: "8px 14px 20px" }}>
                {filtered.length === 0 ? (
                  <div style={{ textAlign: "center", color: "#bbb", padding: "40px 0", fontSize: "0.9rem" }}>
                    대상 인원이 없습니다
                  </div>
                ) : (
                  filtered.map((name) => {
                    const p = customRosterMap[name];
                    const dg: DaegeunType | undefined = currentDaegeun[name];
                    return (
                      <div key={name} style={{
                        display: "flex", alignItems: "center", gap: "10px",
                        padding: "10px 12px", borderRadius: "12px", marginBottom: "7px",
                        background: dg ? "#fef9ee" : "#fafafa",
                        border: dg ? "1.5px solid #f59e0b66" : "1px solid #f3f4f6",
                      }}>
                        {/* 아바타 */}
                        <div style={{
                          width: 32, height: 32, borderRadius: "50%", flexShrink: 0,
                          display: "flex", alignItems: "center", justifyContent: "center",
                          fontWeight: 800, fontSize: "0.82rem", color: "#fff",
                          background: dg ? "#f59e0b" : (p?.group === "주중" ? "#4e89ae" : "#f8b400"),
                        }}>
                          {name.charAt(0)}
                        </div>

                        {/* 이름 + 그룹 */}
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontWeight: 700, fontSize: "0.9rem", color: "#1a1a2e" }}>
                            {name}
                            {dg && (
                              <span style={{
                                marginLeft: "6px", fontSize: "0.68rem",
                                background: "#f59e0b", color: "#fff",
                                padding: "1px 6px", borderRadius: "6px", fontWeight: 700,
                              }}>
                                대근-{dg}
                              </span>
                            )}
                          </div>
                          {p && (
                            <div style={{ fontSize: "0.65rem", color: GROUP_STYLE[p.group].color }}>
                              {p.조}조 · {GROUP_STYLE[p.group].label}
                            </div>
                          )}
                        </div>

                        {/* 대근 유형 버튼 */}
                        <div style={{ display: "flex", gap: "5px", flexShrink: 0 }}>
                          {(["1부", "2부", "투라운드"] as DaegeunType[]).map((type) => (
                            <button key={type}
                              onClick={() => dg === type ? cancelDaegeun(name) : setDaegeunForDate(name, type)}
                              style={{
                                padding: "5px 8px", borderRadius: "8px", border: "none",
                                background: dg === type ? "#f59e0b" : "#f3f4f6",
                                color: dg === type ? "#fff" : "#6b7280",
                                fontWeight: 700, fontSize: "0.68rem", cursor: "pointer",
                                whiteSpace: "nowrap",
                              }}>
                              {type === "투라운드" ? "투R" : type}
                            </button>
                          ))}
                        </div>
                      </div>
                    );
                  })
                )}
              </div>

              {/* 하단 닫기 */}
              <div style={{ padding: "10px 14px", borderTop: "1px solid #f0f0f0", flexShrink: 0 }}>
                <button onClick={() => setBatchDaegeunOpen(false)}
                  style={{
                    width: "100%", padding: "13px", borderRadius: "12px", border: "none",
                    background: "#fef3c7", color: "#92400e",
                    fontWeight: 700, fontSize: "0.95rem", cursor: "pointer",
                  }}>
                  닫기
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* ─── 대근 선택 모달 ─── */}
      {daegeunModal && (
        <div style={{
          position: "fixed", inset: 0, zIndex: 500,
          background: "rgba(0,0,0,0.6)",
          display: "flex", alignItems: "center", justifyContent: "center",
        }}
          onClick={(e) => { if (e.target === e.currentTarget) setDaegeunModal(null); }}
        >
          <div style={{
            background: "#fff", borderRadius: 20, padding: "24px 20px 20px",
            maxWidth: 300, width: "90%", textAlign: "center",
            boxShadow: "0 8px 40px rgba(0,0,0,0.28)",
          }}>
            <div style={{ fontWeight: 800, fontSize: "1rem", color: "#1a1a2e", marginBottom: "4px" }}>
              대근 유형 선택
            </div>
            <div style={{ fontSize: "0.82rem", color: "#888", marginBottom: "18px" }}>
              {daegeunModal}
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: "10px", marginBottom: "14px" }}>
              {(["1부", "2부", "투라운드"] as DaegeunType[]).map((type) => (
                <button key={type}
                  onClick={() => setDaegeunForDate(daegeunModal, type)}
                  style={{
                    padding: "13px", borderRadius: "12px", border: "none",
                    background: currentDaegeun[daegeunModal] === type ? "#f59e0b" : "#fef3c7",
                    color: currentDaegeun[daegeunModal] === type ? "#fff" : "#78350f",
                    fontWeight: 700, fontSize: "0.95rem", cursor: "pointer",
                    boxShadow: currentDaegeun[daegeunModal] === type ? "0 2px 8px #f59e0b66" : "none",
                  }}>
                  {type === "투라운드" ? "투라운드 (1부+2부)" : type}
                </button>
              ))}
            </div>
            {currentDaegeun[daegeunModal] && (
              <button
                onClick={() => cancelDaegeun(daegeunModal)}
                style={{
                  width: "100%", padding: "10px", borderRadius: "10px",
                  border: "1px solid #e5e7eb", background: "#f9fafb",
                  color: "#6b7280", fontWeight: 700, fontSize: "0.85rem", cursor: "pointer",
                  marginBottom: "8px",
                }}>
                대근 해제
              </button>
            )}
            <button
              onClick={() => setDaegeunModal(null)}
              style={{
                width: "100%", padding: "10px", borderRadius: "10px",
                border: "none", background: "#f3f4f6",
                color: "#6b7280", fontWeight: 600, fontSize: "0.85rem", cursor: "pointer",
              }}>
              취소
            </button>
          </div>
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

            <div style={{ display: "flex", gap: "8px", marginTop: "8px" }}>
              <button onClick={assign} style={{ ...S.primaryBtn, flex: 1 }}>배정하기</button>
              <button onClick={generateWeek} style={{ ...S.primaryBtn, flex: 1, background: "#374151" }}>
                일주일 생성
              </button>
            </div>
          </div>

          {/* ── 컷 기준 요약 ── */}
          {livePreview && names.length > 0 && (
            <div style={{
              background: "#f8f9ff", border: "1.5px solid #c5cae9", borderRadius: 12,
              padding: "10px 14px", display: "flex", flexDirection: "column", gap: 5,
            }}>
              <div style={{ fontWeight: 800, fontSize: "0.78rem", color: "#3949ab", marginBottom: 3 }}>
                ✂️ 컷 기준 요약
              </div>

              {mode === "2부제" ? (() => {
                const s1 = livePreview.shift1 ?? [];
                const s2 = livePreview.shift2 ?? [];
                const sp1 = livePreview.spare1 ?? [];
                const sp2 = livePreview.spare2 ?? [];
                const cutRows: React.ReactNode[] = [];

                if (s1.length > 0) {
                  cutRows.push(
                    <div key="s1-cut" style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <span style={{
                        fontSize: "0.7rem", fontWeight: 800, color: "#1565c0",
                        background: "#e3f2fd", borderRadius: 6, padding: "2px 8px",
                        minWidth: 76, textAlign: "center", flexShrink: 0,
                      }}>1부 컷</span>
                      <span style={{ fontWeight: 800, fontSize: "0.95rem", color: "#1565c0" }}>
                        {s1[s1.length - 1]}
                      </span>
                      <span style={{ fontSize: "0.72rem", color: "#90a4ae" }}>까지</span>
                    </div>
                  );
                }

                if (s2.length > 0) {
                  cutRows.push(
                    <div key="s2-cut" style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <span style={{
                        fontSize: "0.7rem", fontWeight: 800, color: "#2e7d32",
                        background: "#e8f5e9", borderRadius: 6, padding: "2px 8px",
                        minWidth: 76, textAlign: "center", flexShrink: 0,
                      }}>2부 컷</span>
                      <span style={{ fontWeight: 800, fontSize: "0.95rem", color: "#2e7d32" }}>
                        {s2[s2.length - 1]}
                      </span>
                      <span style={{ fontSize: "0.72rem", color: "#90a4ae" }}>까지</span>
                    </div>
                  );
                }

                if (sp1.length > 0) {
                  cutRows.push(
                    <div key="sp1" style={{
                      display: "flex", alignItems: "center", gap: 8,
                      borderTop: "1px dashed #e0e7ff", paddingTop: 5, marginTop: 2,
                    }}>
                      <span style={{
                        fontSize: "0.7rem", fontWeight: 800, color: "#e65100",
                        background: "#fff3e0", borderRadius: 6, padding: "2px 8px",
                        minWidth: 76, textAlign: "center", flexShrink: 0,
                      }}>1부 스페어</span>
                      <span style={{ fontWeight: 800, fontSize: "0.95rem", color: "#e65100" }}>
                        {sp1[0]}
                      </span>
                    </div>
                  );
                }

                cutRows.push(
                  <div key="sp2" style={{ display: "flex", alignItems: "flex-start", gap: 8 }}>
                    <span style={{
                      fontSize: "0.7rem", fontWeight: 800, color: "#92400e",
                      background: "#fef3c7", borderRadius: 6, padding: "2px 8px",
                      minWidth: 76, textAlign: "center", flexShrink: 0, marginTop: sp2.length > 0 ? 2 : 0,
                    }}>2부 스페어</span>
                    {sp2.length === 0 ? (
                      <span style={{ fontWeight: 700, fontSize: "0.9rem", color: "#d1d5db" }}>없음</span>
                    ) : (
                      <div style={{ display: "flex", flexDirection: "column", gap: 2, flex: 1 }}>
                        {sp2.slice(0, 2).map((n, i) => (
                          <div key={n} style={{ display: "flex", alignItems: "center", gap: 5 }}>
                            <span style={{
                              fontSize: "0.62rem", fontWeight: 700, color: "#92400e",
                              background: "#fcd34d", borderRadius: 4, padding: "1px 5px",
                              flexShrink: 0,
                            }}>{i + 1}번</span>
                            <span style={{ fontWeight: 800, fontSize: "0.95rem", color: "#b45309" }}>{n}</span>
                            {i === 0 && (
                              <span style={{
                                marginLeft: "auto", fontSize: "0.68rem", fontWeight: 700,
                                color: "#92400e", background: "#fcd34d", borderRadius: 5,
                                padding: "1px 6px", flexShrink: 0,
                              }}>→ 내일 첫번호</span>
                            )}
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                );

                return cutRows;
              })() : (() => {
                const s1 = livePreview.shift1 ?? [];
                const sp2 = livePreview.spare2 ?? [];
                const cutRows: React.ReactNode[] = [];

                if (s1.length > 0) {
                  cutRows.push(
                    <div key="dan-cut" style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <span style={{
                        fontSize: "0.7rem", fontWeight: 800, color: "#1565c0",
                        background: "#e3f2fd", borderRadius: 6, padding: "2px 8px",
                        minWidth: 76, textAlign: "center", flexShrink: 0,
                      }}>컷</span>
                      <span style={{ fontWeight: 800, fontSize: "0.95rem", color: "#1565c0" }}>
                        {s1[s1.length - 1]}
                      </span>
                      <span style={{ fontSize: "0.72rem", color: "#90a4ae" }}>까지</span>
                    </div>
                  );
                }

                if (sp2.length > 0) {
                  cutRows.push(
                    <div key="dan-sp" style={{
                      display: "flex", alignItems: "center", gap: 8,
                      borderTop: "1px dashed #e0e7ff", paddingTop: 5, marginTop: 2,
                    }}>
                      <span style={{
                        fontSize: "0.7rem", fontWeight: 800, color: "#6a1b9a",
                        background: "#f3e5f5", borderRadius: 6, padding: "2px 8px",
                        minWidth: 76, textAlign: "center", flexShrink: 0,
                      }}>스페어</span>
                      <span style={{ fontWeight: 800, fontSize: "0.95rem", color: "#6a1b9a" }}>
                        {sp2.slice(0, 2).join("  ·  ")}
                      </span>
                    </div>
                  );
                }

                return cutRows;
              })()}
            </div>
          )}

          {/* 1일 결과 */}
          {dayResult && weekly.length === 0 && (
            <div style={S.card} id="print-area">
              <div style={{ ...S.sectionTitle, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                <span>📋 {selectedDate ? selectedDate.dateLabel : DAY_LABELS[dayOfWeek] + "요일"} 배정 결과</span>
                <button
                  onClick={() => window.print()}
                  style={{ ...S.smallBtn, fontSize: "0.75rem", padding: "4px 10px" }}
                >
                  🖨️ 출력
                </button>
              </div>
              <DayResultView result={dayResult} mode={mode} />
            </div>
          )}

          {/* 주간 결과 */}
          {weekly.length > 0 && (
            <div style={S.card} id="print-area">
              <div style={{ ...S.sectionTitle, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                <span>📅 주간 근무표 (월~일)</span>
                <button onClick={() => window.print()} style={{ ...S.smallBtn, fontSize: "0.75rem", padding: "4px 10px" }}>
                  🖨️ 출력
                </button>
              </div>

              {weekly.map(({ day, result: r }, di) => {
                const isExpanded = expandedDays.has(day);
                const isWeekend  = di === 5 || di === 6;
                const chipBg = isWeekend
                  ? "linear-gradient(135deg, #c62828, #ef5350)"
                  : "linear-gradient(135deg, #1a1a2e, #4e89ae)";

                // ── 요약 행: 컷 기준 + 스페어만 표시 ──
                const SummaryRow = () => {
                  const cut1  = r.shift1?.at(-1) ?? "-";
                  const spare1nm = r.spare1?.[0] ?? "-";
                  const sp2_1 = r.spare2?.[0];
                  const sp2_2 = r.spare2?.[1];

                  const InfoChip = ({ label, value, labelColor, labelBg }: { label: string; value: string; labelColor: string; labelBg: string }) => (
                    <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                      <span style={{
                        fontSize: "0.65rem", fontWeight: 700, color: labelColor,
                        background: labelBg, borderRadius: 6, padding: "1px 6px", flexShrink: 0,
                      }}>{label}</span>
                      <span style={{ fontSize: "0.8rem", fontWeight: 600, color: "#1a1a2e" }}>{value}</span>
                    </div>
                  );

                  if (mode === "2부제") {
                    const shift2NoSpare1 = (r.shift2 ?? []).filter(n => !r.spare1?.includes(n));
                    const cut2 = shift2NoSpare1.at(-1) ?? "-";
                    return (
                      <div style={{ display: "flex", flexDirection: "column", gap: 3, flex: 1 }}>
                        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                          <InfoChip label="1부컷" value={cut1}   labelColor="#1565c0" labelBg="#e3f2fd" />
                          <InfoChip label="2부컷" value={cut2}   labelColor="#2e7d32" labelBg="#e8f5e9" />
                        </div>
                        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                          <InfoChip label="1부스페어" value={spare1nm} labelColor="#9a3412" labelBg="#fed7aa" />
                          {sp2_1 && <InfoChip label="2부스페어①" value={sp2_1} labelColor="#6a1b9a" labelBg="#f3e5f5" />}
                          {sp2_2 && <InfoChip label="2부스페어②" value={sp2_2} labelColor="#6a1b9a" labelBg="#f3e5f5" />}
                        </div>
                      </div>
                    );
                  } else {
                    const sp = r.spare2?.[0] ?? r.spare1?.[0] ?? "-";
                    return (
                      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", flex: 1 }}>
                        <InfoChip label="컷" value={cut1} labelColor="#1565c0" labelBg="#e3f2fd" />
                        <InfoChip label="스페어" value={sp} labelColor="#6a1b9a" labelBg="#f3e5f5" />
                      </div>
                    );
                  }
                };

                return (
                  <div key={day} style={{ ...S.weekDay, flexDirection: "column", gap: 6 }}>
                    {/* 날짜 헤더 행: 칩 + 요약/자세히 버튼 */}
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <div style={{ ...S.dayChip, background: chipBg }}>
                        <span style={{ fontSize: "0.85rem", fontWeight: 800 }}>{day}</span>
                        <span style={{ fontSize: "0.55rem", opacity: 0.8 }}>요일</span>
                      </div>
                      <SummaryRow />
                      <button
                        onClick={() => toggleDayExpand(day)}
                        style={{
                          flexShrink: 0, padding: "3px 10px",
                          border: "1px solid #d1d5db", borderRadius: 8,
                          background: isExpanded ? "#1a1a2e" : "#f9fafb",
                          color: isExpanded ? "#fff" : "#374151",
                          fontSize: "0.72rem", fontWeight: 700, cursor: "pointer",
                        }}
                      >
                        {isExpanded ? "요약" : "자세히"}
                      </button>
                    </div>

                    {/* 상세 내용 (토글) */}
                    {isExpanded && (
                      <div style={{ paddingLeft: 4 }}>
                        <DayResultView result={r} mode={mode} compact />
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}

          {/* 인원 리스트 (조별 구분) — 배정 결과 아래로 이동 */}
          <div style={S.card}>
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
                    <div key={`jo-${joNum}-${idx}`} style={{
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
                const isWeekendOnWeekday = person != null && isAutoOff(person.group, dayOfWeek) && !(name in manualStatuses);
                const daegeunType: DaegeunType | undefined = currentDaegeun[name];

                rows.push(
                  <div key={name} style={{
                    ...S.personRow,
                    opacity: effS === "휴무" ? 0.45 : 1,
                  }}>
                    <span style={S.personNum}>{person?.no ?? idx + 1}</span>
                    {/* 아바타 이니셜 */}
                    <div style={{
                      width: 34, height: 34, borderRadius: "50%", flexShrink: 0,
                      display: "flex", alignItems: "center", justifyContent: "center",
                      fontWeight: 800, fontSize: "0.85rem", color: "white",
                      background: effS && effS !== "없음"
                        ? (STATUS_COLOR[effS]?.bg ?? "#e5e7eb")
                        : person?.group === "하우스" ? "#52de97"
                        : person?.group === "주중" ? "#4e89ae" : "#f8b400",
                      boxShadow: "0 1px 4px rgba(0,0,0,0.12)",
                    }}>
                      <span style={{ color: effS && effS !== "없음" ? (STATUS_COLOR[effS]?.color ?? "#333") : "#fff" }}>
                        {name.charAt(0)}
                      </span>
                    </div>

                    <div style={{ minWidth: "68px" }}>
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
                        {isWeekendOnWeekday && daegeunType && (
                          <span style={{
                            fontSize: "0.6rem", padding: "1px 4px", borderRadius: "4px",
                            background: "#fef3c7", color: "#92400e", fontWeight: 700,
                          }}>
                            대근-{daegeunType}
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
                      {isWeekendOnWeekday ? (
                        <button
                          onClick={() => setDaegeunModal(name)}
                          style={{
                            ...S.statusBtn,
                            background: daegeunType ? "#fbbf24" : "#f0f0f0",
                            color: daegeunType ? "#78350f" : "#555",
                            border: daegeunType ? "2px solid #f59e0b" : "1px solid #e0e0e0",
                            fontWeight: 700,
                            minWidth: "70px",
                          }}>
                          {daegeunType ? `대근-${daegeunType}` : "대근"}
                        </button>
                      ) : (
                        STATUS_BUTTONS.filter((btn) =>
                          mode !== "단부제" || (btn !== "조출" && btn !== "후출")
                        ).map((btn) => {
                          const active = effS === btn;
                          const isAutoActive = active && isAutoHumu;
                          const disabled = (btn === "조출" && !cho가능 && effS !== "조출");
                          const col = active ? STATUS_COLOR[btn!] : null;
                          const maxReached =
                            (btn === "조출" && cho현재수 >= 4 && effS !== "조출") ||
                            (btn === "후출" && hu현재수 >= 4 && effS !== "후출");
                          return (
                            <button key={btn} disabled={disabled || maxReached}
                              onClick={() => toggleStatus(name, btn)}
                              title={
                                btn === "조출" && !cho가능 ? "1부 6팀 이상일 때만 사용 가능" :
                                btn === "조출" && maxReached ? "조출 최대 4명" :
                                btn === "후출" && maxReached ? "후출 최대 4명" : ""
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
                        })
                      )}
                    </div>
                  </div>
                );
              });
              return rows;
            })()}

          </div>

          {false && livePreview && names.length > 0 && (
            <div style={{
              background: "#f8f9ff", border: "1.5px solid #c5cae9", borderRadius: 12,
              padding: "10px 14px", display: "flex", flexDirection: "column", gap: 6,
            }}>
              <div style={{ fontWeight: 700, fontSize: "0.78rem", color: "#3949ab", marginBottom: 2 }}>
                🔍예상 스페어
              </div>

              {mode === "2부제" ? (<>
                {/* 1부 마지막 */}
                {livePreview.shift1.length > 0 && (
                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <span style={{
                      fontSize: "0.7rem", fontWeight: 800, color: "#1565c0",
                      background: "#e3f2fd", borderRadius: 6, padding: "2px 7px", minWidth: 70, textAlign: "center",
                    }}>1부 마지막</span>
                    <span style={{ fontSize: 15, fontWeight: 800, color: "#1565c0" }}>
                      {livePreview.shift1[livePreview.shift1.length - 1]}
                    </span>
                    <span style={{ fontSize: "0.7rem", color: "#90a4ae" }}>
                      (총 {livePreview.shift1.length}명)
                    </span>
                  </div>
                )}

                {/* 1부 스페어 */}
                {livePreview.spare1.length > 0 && (
                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <span style={{
                      fontSize: "0.7rem", fontWeight: 800, color: "#e65100",
                      background: "#fff3e0", borderRadius: 6, padding: "2px 7px", minWidth: 70, textAlign: "center",
                    }}>1부 스페어</span>
                    <span style={{ fontSize: 15, fontWeight: 800, color: "#e65100" }}>
                      {livePreview.spare1.join(", ")}
                    </span>
                  </div>
                )}

                {/* 투라운드(찾근) — 2부 몇팀째 */}
                {livePreview.twoRound?.length > 0 && livePreview.shift2?.length > 0 && (() => {
                  const positions = livePreview.twoRound
                    .map((n: string) => livePreview.shift2.indexOf(n))
                    .filter((i: number) => i >= 0)
                    .map((i: number) => i + 1);
                  const minPos = positions.length > 0 ? Math.min(...positions) : null;
                  const maxPos = positions.length > 0 ? Math.max(...positions) : null;
                  return (
                    <div style={{
                      display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap",
                      background: "#ecfeff", border: "1.5px solid #a5f3fc", borderRadius: 8, padding: "5px 10px",
                    }}>
                      <span style={{
                        fontSize: "0.7rem", fontWeight: 800, color: "#164e63",
                        background: "#cffafe", borderRadius: 6, padding: "2px 7px", minWidth: 70, textAlign: "center", flexShrink: 0,
                      }}>🔄 투라운드</span>
                      <span style={{ fontSize: 14, fontWeight: 800, color: "#0e7490" }}>
                        {livePreview.twoRound.join("  ·  ")}
                      </span>
                      <span style={{ fontSize: "0.7rem", color: "#90a4ae" }}>
                        ({livePreview.twoRound.length}명)
                      </span>
                      {minPos !== null && (
                        <span style={{
                          marginLeft: "auto", fontSize: "0.72rem", fontWeight: 700,
                          color: "#0e7490", background: "#a5f3fc", borderRadius: 5, padding: "1px 7px", flexShrink: 0,
                        }}>
                          2부 {minPos === maxPos ? `${minPos}팀` : `${minPos}~${maxPos}팀`}
                        </span>
                      )}
                    </div>
                  );
                })()}

                {/* 2부 마지막 */}
                {livePreview.shift2.length > 0 && (
                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <span style={{
                      fontSize: "0.7rem", fontWeight: 800, color: "#2e7d32",
                      background: "#e8f5e9", borderRadius: 6, padding: "2px 7px", minWidth: 70, textAlign: "center",
                    }}>2부 마지막</span>
                    <span style={{ fontSize: 15, fontWeight: 800, color: "#2e7d32" }}>
                      {livePreview.shift2[livePreview.shift2.length - 1]}
                    </span>
                    <span style={{ fontSize: "0.7rem", color: "#90a4ae" }}>
                      (총 {livePreview.shift2.length}명)
                    </span>
                  </div>
                )}

                {/* 2부 스페어 — 항상 표시 */}
                <div style={{
                  display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap",
                  background: livePreview.spare2.length > 0 ? "#fffbeb" : "#f9f9f9",
                  border: `1.5px solid ${livePreview.spare2.length > 0 ? "#fcd34d" : "#e0e0e0"}`,
                  borderRadius: 8, padding: "5px 10px",
                }}>
                  <span style={{
                    fontSize: "0.7rem", fontWeight: 800,
                    color: livePreview.spare2.length > 0 ? "#92400e" : "#999",
                    background: livePreview.spare2.length > 0 ? "#fef3c7" : "#f0f0f0",
                    borderRadius: 6, padding: "2px 7px", minWidth: 70, textAlign: "center", flexShrink: 0,
                  }}>🏁 2부스페어</span>
                  {livePreview.spare2.length > 0 ? (<>
                    <span style={{ fontSize: 15, fontWeight: 800, color: "#b45309" }}>
                      {livePreview.spare2.join("  ·  ")}
                    </span>
                    <span style={{ fontSize: "0.7rem", color: "#90a4ae" }}>
                      ({livePreview.spare2.length}명)
                    </span>
                    <span style={{
                      marginLeft: "auto", fontSize: "0.72rem", fontWeight: 700,
                      color: "#92400e", background: "#fcd34d", borderRadius: 5, padding: "1px 7px", flexShrink: 0,
                    }}>→ 내일 첫번호</span>
                  </>) : (
                    <span style={{ fontSize: "0.8rem", color: "#bbb" }}>없음 (전원 2부 배정)</span>
                  )}
                </div>
              </>) : (<>
                {/* 단부제: 단부 마지막 & 스페어 */}
                {livePreview.shift1.length > 0 && (
                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <span style={{
                      fontSize: "0.7rem", fontWeight: 800, color: "#1565c0",
                      background: "#e3f2fd", borderRadius: 6, padding: "2px 7px", minWidth: 70, textAlign: "center",
                    }}>단부 마지막</span>
                    <span style={{ fontSize: 15, fontWeight: 800, color: "#1565c0" }}>
                      {livePreview.shift1[livePreview.shift1.length - 1]}
                    </span>
                    <span style={{ fontSize: "0.7rem", color: "#90a4ae" }}>
                      (총 {livePreview.shift1.length}명)
                    </span>
                  </div>
                )}
                {livePreview.spare2.length > 0 && (
                  <div style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
                    <span style={{
                      fontSize: "0.7rem", fontWeight: 800, color: "#6a1b9a",
                      background: "#f3e5f5", borderRadius: 6, padding: "2px 7px", minWidth: 70, textAlign: "center", flexShrink: 0,
                    }}>스페어</span>
                    <span style={{ fontSize: 14, fontWeight: 700, color: "#6a1b9a" }}>
                      {livePreview.spare2.join("  ·  ")}
                    </span>
                    <span style={{ fontSize: "0.7rem", color: "#90a4ae" }}>
                      ({livePreview.spare2.length}명)
                    </span>
                  </div>
                )}
              </>)}
            </div>
          )}


        </>
      )}

      {/* ── 플로팅 바: 다음날 첫번호 ── */}
      {dayResult && dayResult.spare2?.[0] && weekly.length === 0 && (
        <div style={S.floatingBar}>
          <img src={`${BASE}/char_smile.png`} alt="" style={{ width: 36, height: 36, objectFit: "contain" }} />
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: "0.68rem", color: "rgba(255,255,255,0.65)", marginBottom: 2 }}>🏁 내일 2부 첫번호</div>
            <div style={{ fontWeight: 800, fontSize: "1.1rem", color: "#f8b400" }}>{dayResult.spare2[0]}</div>
          </div>
          {dayResult.spare2[1] && (
            <div style={{ textAlign: "right" }}>
              <div style={{ fontSize: "0.65rem", color: "rgba(255,255,255,0.55)" }}>대기</div>
              <div style={{ fontWeight: 600, fontSize: "0.85rem", color: "rgba(255,255,255,0.8)" }}>{dayResult.spare2[1]}</div>
            </div>
          )}
        </div>
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
  { key: "twoRound" as const, label: "🔄 투라운드",          badge: { bg: "#cffafe", color: "#164e63" } },
  { key: "shift1"   as const, label: "☀️ 1부",               badge: { bg: "#dbeafe", color: "#1e40af" } },
  { key: "spare1"   as const, label: "⚡ 1부스페어→2부1번째", badge: { bg: "#fed7aa", color: "#9a3412" } },
  { key: "shift2"   as const, label: "🌙 2부",               badge: { bg: "#ede9fe", color: "#5b21b6" } },
  { key: "spare2"   as const, label: "🏁 2부스페어",          badge: { bg: "#fef3c7", color: "#92400e" } },
  { key: "excluded" as const, label: "💤 휴무/제외",          badge: { bg: "#f3f4f6", color: "#6b7280" } },
];
const CATS_SINGLE = [
  { key: "twoRound" as const, label: "🔄 투라운드", badge: { bg: "#cffafe", color: "#164e63" } },
  { key: "shift1"   as const, label: "⛳ 단부",     badge: { bg: "#dbeafe", color: "#1e40af" } },
  { key: "spare2"   as const, label: "🏁 스페어",   badge: { bg: "#fef3c7", color: "#92400e" } },
  { key: "excluded" as const, label: "💤 휴무/제외", badge: { bg: "#f3f4f6", color: "#6b7280" } },
];

function DayResultView({ result, mode, compact = false }: {
  result: DayResult; mode: Mode; compact?: boolean;
}) {
  const cats = mode === "2부제" ? CATS_DOUBLE : CATS_SINGLE;
  const 조출Set = new Set(result.조출List ?? []);
  const 후출Set = new Set(result.후출List ?? []);
  const spare1Set = new Set(result.spare1 ?? []); // 1부스페어는 shift2 앞에 이미 배정 → 중복 제거용

  function renderPeople(people: string[], key: string) {
    if (compact) return people.join("  ·  ");
    return (
      <span style={{ fontSize: "0.88rem", color: "#333", lineHeight: 1.7 }}>
        {people.map((n, i) => {
          const isCho   = (key === "shift1") && 조출Set.has(n);
          const isHu    = (key === "shift2") && 후출Set.has(n);
          const isSpare1= (key === "shift2") && spare1Set.has(n);
          const isTwoR  = key === "twoRound";
          const suffix  = isCho ? " [조출]" : isHu ? " [후출]" : isSpare1 ? " [1부스페어]" : "";
          return (
            <span key={n}>
              {i > 0 && <span style={{ color: "#d1d5db" }}> · </span>}
              <span style={{
                fontWeight: (isCho || isHu || isTwoR || isSpare1) ? 800 : 500,
                color: isCho ? "#9a3412" : isHu ? "#5b21b6" : isTwoR ? "#164e63" : isSpare1 ? "#9a3412" : "#374151",
                background: isCho ? "#fed7aa" : isHu ? "#ddd6fe" : "transparent",
                borderRadius: 4, padding: (isCho || isHu) ? "1px 4px" : 0,
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
        // shift2에서 spare1 중복 제거 (1부스페어는 별도 행에 표시되므로)
        const rawPeople = result[key];
        const people = key === "shift2" ? rawPeople.filter(n => !spare1Set.has(n)) : rawPeople;
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
  // 찾근은 당일만 적용 — 다음날 예상 순번에서는 태그 없음
  // 오직 2부스페어만 강조 표시 (내일 첫번호이므로)
  const spare2Set = new Set(spare2);

  function tagOf(name: string) {
    if (spare2Set.has(name)) return { label: "2부스페어", color: "#6a1b9a", bg: "#f3e5f5" };
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
          (2부스페어 앞번호 → 나머지 순번 대기 순)
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
        ▶ 2부스페어가 다음날 앞번호(첫 대기)
        {spare1.length > 0 && (
          <span style={{ marginLeft: "8px", color: "#e65100", fontWeight: 600 }}>
            오늘 1부스페어: {spare1.join(", ")} (2부 1번째로 나감)
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
  /* ─── Layout ─── */
  page: { fontFamily: "'Inter', sans-serif", background: "#eef2f7", minHeight: "100dvh", paddingBottom: "80px" },
  header: {
    display: "flex", alignItems: "center", gap: "10px",
    padding: "14px 16px",
    background: "linear-gradient(135deg, #1a1a2e 0%, #16213e 100%)",
    color: "white", position: "sticky", top: 0, zIndex: 20,
    boxShadow: "0 2px 12px rgba(26,26,46,0.25)",
  },
  backBtn: {
    background: "rgba(255,255,255,0.12)", border: "1px solid rgba(255,255,255,0.15)",
    color: "white", borderRadius: "10px", padding: "8px 14px",
    cursor: "pointer", fontSize: "1rem", minHeight: "44px",
  },
  smallBtn: {
    marginLeft: "auto",
    background: "rgba(255,255,255,0.12)", border: "1px solid rgba(255,255,255,0.15)",
    color: "white", borderRadius: "10px", padding: "8px 14px",
    cursor: "pointer", fontSize: "0.82rem", minHeight: "44px",
  },
  headerTitle: { fontWeight: 700, fontSize: "1.05rem", letterSpacing: "-0.02em" },

  /* ─── Cards ─── */
  card: {
    background: "#ffffff", borderRadius: "18px", padding: "18px 16px",
    margin: "10px 12px", boxShadow: "0 2px 12px rgba(26,26,46,0.08)",
  },
  card1부: {
    background: "linear-gradient(135deg, #e8f4fd 0%, #dbeafe 100%)",
    borderRadius: "14px", padding: "14px",
    border: "1.5px solid #93c5fd", marginBottom: "10px",
  },
  card2부: {
    background: "linear-gradient(135deg, #fdf4ff 0%, #ede9fe 100%)",
    borderRadius: "14px", padding: "14px",
    border: "1.5px solid #c4b5fd", marginBottom: "10px",
  },
  cardSpare: {
    background: "linear-gradient(135deg, #fffbeb 0%, #fef3c7 100%)",
    borderRadius: "14px", padding: "14px",
    border: "1.5px solid #fcd34d",
  },

  /* ─── Typography & Labels ─── */
  label: {
    display: "block", fontSize: "0.7rem", color: "#9ca3af",
    marginBottom: "8px", fontWeight: 700,
    textTransform: "uppercase", letterSpacing: "0.06em",
  },
  sectionTitle: {
    fontWeight: 800, fontSize: "0.95rem", marginBottom: "14px",
    color: "#1a1a2e", letterSpacing: "-0.01em",
  },

  /* ─── Mode Segment Control ─── */
  segmentTrack: {
    display: "flex", background: "#e5e7eb", borderRadius: "12px",
    padding: "3px", marginBottom: "18px", gap: "2px",
  },
  segmentBtn: {
    flex: 1, padding: "10px", border: "none", borderRadius: "10px",
    fontSize: "0.9rem", fontWeight: 700, cursor: "pointer", transition: "all 0.2s",
    minHeight: "44px",
  },

  /* ─── Day-of-week pills ─── */
  dayBtn: {
    padding: "8px 14px", border: "none", borderRadius: "10px",
    fontSize: "0.85rem", fontWeight: 700, cursor: "pointer", minWidth: "44px", minHeight: "44px",
  },

  /* ─── Date calendar grid ─── */
  dateGrid: {
    display: "grid", gridTemplateColumns: "repeat(5, 1fr)",
    gap: "5px", marginBottom: "14px",
  },
  dateBtn: {
    display: "flex", flexDirection: "column", alignItems: "center", gap: "2px",
    padding: "8px 4px", borderRadius: "10px", cursor: "pointer",
    fontSize: "0.73rem", fontWeight: 600, border: "1.5px solid #e5e7eb",
    minHeight: "52px", transition: "all 0.15s",
  },

  /* ─── Info rows ─── */
  infoRow: { display: "flex", gap: "6px", marginBottom: "12px", flexWrap: "wrap" },
  chip: {
    padding: "4px 12px", borderRadius: "20px", background: "#f3f4f6",
    color: "#374151", fontSize: "0.78rem", fontWeight: 600,
  },

  /* ─── Inputs ─── */
  textarea: {
    width: "100%", padding: "12px", borderRadius: "10px",
    border: "1.5px solid #e5e7eb", fontSize: "0.95rem", resize: "vertical",
    fontFamily: "'Inter', sans-serif", marginBottom: "14px", boxSizing: "border-box",
    background: "#fafafa",
  },
  numInput: {
    width: "100%", padding: "12px", borderRadius: "10px",
    border: "1.5px solid #e5e7eb", fontSize: "1rem",
    boxSizing: "border-box", background: "#fafafa",
  },
  primaryBtn: {
    width: "100%", padding: "15px",
    background: "linear-gradient(135deg, #1a1a2e 0%, #16213e 100%)",
    color: "white", border: "none", borderRadius: "12px",
    fontSize: "1rem", fontWeight: 800, cursor: "pointer",
    boxShadow: "0 4px 14px rgba(26,26,46,0.3)", minHeight: "50px",
  },

  /* ─── Person list ─── */
  personRow: {
    display: "flex", alignItems: "center", gap: "10px",
    padding: "10px 0", borderBottom: "1px solid #f3f4f6", flexWrap: "wrap",
    transition: "opacity 0.2s",
  },
  personNum: { minWidth: "20px", fontSize: "0.72rem", color: "#d1d5db", textAlign: "right" },
  personName: { fontWeight: 700, fontSize: "0.88rem", marginBottom: "2px" },
  btnGroup: { display: "flex", flexWrap: "wrap", gap: "4px", flex: 1 },
  statusBtn: {
    padding: "5px 10px", borderRadius: "8px", fontSize: "0.76rem",
    cursor: "pointer", fontWeight: 700, transition: "all 0.15s", minHeight: "32px",
  },

  /* ─── Weekly view ─── */
  weekDay: {
    display: "flex", gap: "10px", padding: "12px 0",
    borderBottom: "1px solid #f3f4f6", alignItems: "flex-start",
  },
  dayChip: {
    minWidth: "56px", borderRadius: "10px", color: "white",
    display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center",
    fontWeight: 800, fontSize: "0.7rem", flexShrink: 0, padding: "6px 4px", lineHeight: 1.3,
  },

  /* ─── Info boxes ─── */
  calcBox: {
    display: "flex", gap: "8px", flexWrap: "wrap", alignItems: "center",
    background: "#f8fafc", borderRadius: "10px", padding: "10px 14px",
    fontSize: "0.82rem", fontWeight: 600, border: "1px solid #e2e8f0",
  },
  cutoffBox: {
    display: "flex", gap: "8px", flexWrap: "wrap", alignItems: "center",
    background: "#fafafa", borderRadius: "10px", padding: "10px 14px",
    marginBottom: "8px", border: "1px solid #e5e7eb",
  },
  excelInfo: {
    background: "linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%)",
    borderRadius: "12px", padding: "12px 14px",
    marginBottom: "14px", border: "1px solid #bfdbfe",
  },
  excelInfoTitle: { fontSize: "0.78rem", fontWeight: 700, color: "#1d4ed8", marginBottom: "8px" },
  excelStatRow: { display: "flex", gap: "6px", flexWrap: "wrap" },
  dateBar: {
    display: "flex", alignItems: "center", gap: "6px",
    background: "#f8fafc", borderRadius: "10px", padding: "10px 14px",
    marginBottom: "10px", flexWrap: "wrap", border: "1px solid #e2e8f0",
  },

  /* ─── Floating next-first bar (fixed bottom) ─── */
  floatingBar: {
    position: "fixed", bottom: 0, left: 0, right: 0,
    background: "linear-gradient(135deg, #1a1a2e 0%, #4e89ae 100%)",
    color: "white", padding: "12px 20px", zIndex: 30,
    display: "flex", alignItems: "center", gap: "10px",
    boxShadow: "0 -4px 20px rgba(26,26,46,0.3)",
  },
};
