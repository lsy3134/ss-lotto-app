import { useState, useEffect, useMemo, useRef } from "react";
import { useLocation } from "wouter";
import * as XLSX from "xlsx";
import { ROSTER, isAutoOff, type GroupType, type PersonData } from "../data/roster";
import { createWorker } from "tesseract.js";
import { useAuth } from "../App";

const BASE = import.meta.env.BASE_URL.replace(/\/$/, "");

// ── 타입 ──────────────────────────────────────────
type StatusType =
  | "조출" | "후출" | "찾근" | "대기"
  | "당번" | "병가" | "휴무" | "하우스" | "휴무해제"
  | "VIP1부" | "VIP2부" | "VIP투근무"
  | null;

const VIP_STATUSES = new Set<StatusType>(["VIP1부", "VIP2부", "VIP투근무"]);

type Mode = "2부제" | "단부제";
type DaegeunType = "1부" | "2부" | "투라운드";

const STATUS_BUTTONS: StatusType[] = [
  "대기", "조출", "후출", "찾근", "당번", "병가", "휴무", "하우스",
];

const EXCLUDED_SET = new Set(["당번", "병가", "휴무", "하우스"]);

const STATUS_COLOR: Record<string, { bg: string; color: string }> = {
  대기:     { bg: "#fce7f3", color: "#9d174d" },  // 분홍 (1부 출근대기 = spare1)
  조출:     { bg: "#fed7aa", color: "#9a3412" },  // 주황 (조출)
  후출:     { bg: "#ddd6fe", color: "#5b21b6" },  // 연보라 (속수대기)
  찾근:     { bg: "#cffafe", color: "#164e63" },  // 하늘 (투라운드)
  당번:     { bg: "#fecaca", color: "#991b1b" },  // 연빨 (주의)
  병가:     { bg: "#e5e7eb", color: "#4b5563" },  // 회색 (비활성)
  휴무:     { bg: "#f3f4f6", color: "#6b7280" },  // 연회색 (비활성)
  하우스:   { bg: "#fef08a", color: "#713f12" },  // 금 (하우스)
  휴무해제: { bg: "#dcfce7", color: "#166534" },  // 연초록 (엑셀 휴무 해제)
  VIP1부:   { bg: "#f3e5f5", color: "#7b1fa2" },  // 보라 (VIP 1부)
  VIP2부:   { bg: "#ede7f6", color: "#4527a0" },  // 진보라 (VIP 2부)
  VIP투근무:{ bg: "#e8eaf6", color: "#283593" },  // 남보라 (VIP 투근무)
};

const GROUP_STYLE: Record<GroupType, { bg: string; color: string; label: string }> = {
  하우스: { bg: "#eeebff", color: "#7c6ef7", label: "하우스" },
  주중:   { bg: "#eef3ff", color: "#5b8dee", label: "주중" },
  주말:   { bg: "#fff7ed", color: "#f59e0b", label: "주말" },
};

// 휴무 chip 그룹 dot 색상 (사용자 지정 팔레트)
const GROUP_DOT: Record<GroupType, string> = {
  하우스: "#7c6ef7",
  주중:   "#5b8dee",
  주말:   "#f59e0b",
};

// 이름 정규화 (공백 완전 제거) — 저장/조회/비교 모든 곳에 일관 적용
const normalize = (name: string) => name.replace(/\s+/g, "").trim();
// roster 배열 전체 이름 정규화
const normalizeRoster = <T extends { name: string }>(roster: T[]): T[] =>
  roster.map(p => ({ ...p, name: normalize(p.name) }));

// 이름 → 그룹 조회 맵
const NAME_GROUP: Record<string, GroupType> = Object.fromEntries(
  ROSTER.map(p => [p.name, p.group])
);
// normalize된 key로 그룹 조회 (공백 차이로 인한 key 불일치 방지)
const NAME_GROUP_NORMALIZED: Record<string, GroupType> = Object.fromEntries(
  Object.entries(NAME_GROUP).map(([k, v]) => [normalize(k), v])
);

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

/** Date → "MM.DD (요일)" 형식 key 생성 (저장·조회·계산 전체에서 이 함수만 사용) */
function makeDateKey(date: Date): string {
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  const dayIdx = (date.getDay() + 6) % 7; // Mon=0 … Sun=6
  return `${mm}.${dd} (${KR_DAY[dayIdx]})`;
}

function generateMonthDays(mm: string, year: number): ExcelDayData[] {
  const m = parseInt(mm, 10);
  const daysInMonth = new Date(year, m, 0).getDate();
  const result: ExcelDayData[] = [];
  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(year, m - 1, d);
    const dayIdx = (date.getDay() + 6) % 7; // Mon=0 … Sun=6
    const dayName = KR_DAY[dayIdx];
    result.push({
      dateLabel: makeDateKey(date),
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
  spare2FromTemporaryWork?: string[];
  excluded: string[];
  invalidStatusReasons?: Record<string, string>;
  // 위치 표시용 메타 (후출: 2부 뒤에서3번째, 조출: 1부 앞)
  조출List?: string[];
  후출List?: string[];
  // VIP 배정 (표시용)
  vip1List?: string[];
  vip2List?: string[];
  vipBothList?: string[];
  // 다음날 예상 순번: [스페어(앞번호순)] → [오늘 근무자 + 찾근자(일반 순번으로 포함)]
  nextDayQueue?: string[];
  // 대근 인원 목록 (1부·2부·투라운드 통합, bold 표시용)
  daegeunList?: string[];
}

function getPrioritySpares(result?: Pick<DayResult, "spare2" | "spare2FromTemporaryWork"> | null): string[] {
  if (!result) return [];
  const temporarySet = new Set(result.spare2FromTemporaryWork ?? []);
  return (result.spare2 ?? []).filter(n => !temporarySet.has(n)).slice(0, 2);
}

function getPrioritySparesFromSaved(spare2?: string[], spare2FromTemporaryWork?: string[]): string[] {
  const temporarySet = new Set(spare2FromTemporaryWork ?? []);
  return (spare2 ?? []).filter(n => !temporarySet.has(n)).slice(0, 2);
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
  // 순수 숫자 문자열 "1"~"31" — 달력 날짜 칸 (월 context 필수)
  m = s.match(/^(\d{1,2})$/);
  if (m && month) {
    const d = parseInt(m[1], 10);
    if (d >= 1 && d <= 31) return `${String(month).padStart(2, "0")}.${m[1].padStart(2, "0")}`;
  }
  // 앞 숫자 + 비숫자 혼합: "1 어린이날", "1\n노동절", "01 MBN", "1(공휴일)" 등
  m = s.match(/^(\d{1,2})\D/);
  if (m && month) {
    const d = parseInt(m[1], 10);
    if (d >= 1 && d <= 31) return `${String(month).padStart(2, "0")}.${m[1].padStart(2, "0")}`;
  }
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

// 셀에서 이름 목록 추출 — 쉼표/줄바꿈/공백 구분 포함
function extractNames(v: unknown): string[] {
  if (!v) return [];
  const s = String(v).trim();
  if (!s) return [];
  // 1차: 쉼표/줄바꿈/슬래시로 분리
  const commaParts = s.split(/[,，、\/\n\r]+/).map(p => p.trim()).filter(Boolean);
  const names: string[] = [];
  for (const part of commaParts) {
    if (isKoreanName(part)) {
      names.push(part);
    } else {
      // 2차: 공백으로 추가 분리 (공백 구분 이름 처리)
      const spaceParts = part.split(/\s+/).map(p => p.trim()).filter(p => isKoreanName(p));
      names.push(...spaceParts);
    }
  }
  if (names.length > 0) return names;
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

    // 병합 셀(merged cell) 처리: 셀이 비어 있으면 해당 셀이 속한 merge 범위의 좌상단 셀 값을 반환
    const merges: XLSX.Range[] = ws["!merges"] ?? [];
    const mergeTL = new Map<string, { r: number; c: number }>();
    for (const mg of merges) {
      for (let mr = mg.s.r; mr <= mg.e.r; mr++) {
        for (let mc = mg.s.c; mc <= mg.e.c; mc++) {
          if (mr === mg.s.r && mc === mg.s.c) continue; // 좌상단 자신은 제외
          mergeTL.set(XLSX.utils.encode_cell({ r: mr, c: mc }), { r: mg.s.r, c: mg.s.c });
        }
      }
    }
    const getCell = (r: number, c: number) => {
      const addr = XLSX.utils.encode_cell({ r, c });
      let cell = ws[addr];
      if ((!cell || cell.v === undefined || cell.v === null || cell.v === "") && mergeTL.has(addr)) {
        const tl = mergeTL.get(addr)!;
        cell = ws[XLSX.utils.encode_cell({ r: tl.r, c: tl.c })];
      }
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

    // ── 포맷G 강화: 달력형 — 요일 헤더 + 날짜/이름 혼합 행 ──
    // 구조: 요일 헤더행 → 각 열마다 [날짜번호] → [이름 N행] → [날짜번호] → ...
    // 핵심: r+=2 고정 stride 제거, 열별로 현재 날짜를 독립 추적하며 행을 1행씩 내려감
    const fmtG: Record<string, string[]> = {}; let gS = 0;
    const weekdayHeader = /^[일월화수목금토]$/;
    // 요일 헤더 행 탐색 (시트 앞 10행 이내)
    let headerRow = -1;
    for (let r = range.s.r; r <= Math.min(range.s.r + 9, range.e.r); r++) {
      let wdCount = 0;
      for (let c = range.s.c; c <= range.e.c; c++) {
        const v = String(getCell(r, c) ?? "").trim();
        if (weekdayHeader.test(v)) wdCount++;
      }
      if (wdCount >= 3) { headerRow = r; break; }
    }
    if (headerRow >= 0) {
      // 요일 헤더가 있는 열 위치 수집
      const weekdayCols: number[] = [];
      for (let c = range.s.c; c <= range.e.c; c++) {
        const v = String(getCell(headerRow, c) ?? "").trim();
        if (weekdayHeader.test(v)) weekdayCols.push(c);
      }
      // 열마다 현재 추적 중인 날짜 키 (독립 관리)
      const colCurKey: Record<number, string> = {};

      // 헤더 다음 행부터 끝까지 1행씩 스캔
      for (let r = headerRow + 1; r <= range.e.r; r++) {
        // 이 행이 "요일 헤더 반복"이면 skip (달력 다음 달 헤더 등)
        let wdRepeat = 0;
        for (const c of weekdayCols) {
          const v = String(getCell(r, c) ?? "").trim();
          if (weekdayHeader.test(v)) wdRepeat++;
        }
        if (wdRepeat >= 3) continue; // 요일 헤더 반복행 → skip

        for (const c of weekdayCols) {
          const v = getCell(r, c);
          const dateKey = dk(v);
          if (dateKey) {
            // 날짜값 → 이 열의 현재 날짜 갱신
            colCurKey[c] = dateKey;
            if (!fmtG[dateKey]) fmtG[dateKey] = [];
          } else {
            // 날짜가 아니면 이름 여부 확인 → 현재 날짜에 누적
            const names = extractNames(v);
            const curKey = colCurKey[c];
            if (names.length && curKey) {
              for (const name of names) {
                if (!fmtG[curKey].includes(name)) { fmtG[curKey].push(name); gS++; }
              }
            }
            // 빈 셀이어도 break 없이 continue
          }
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
  daegeunMap: Record<string, string> = {},  // 대근 유형 맵 (1부|2부|투라운드)
  statusOrder: string[] = []                // 찾근/조출/후출 클릭 순서 (앞에 배치)
): DayResult {
  const twoRound: string[] = [];   // 찾근 (1부+2부 투라운드)
  const 조출List: string[] = [];   // 조출 (1부 앞 고정, 최대 6명)
  const 후출List: string[] = [];   // 후출 (2부 뒤에서 3번째, 최대 6명)
  const 대기List: string[] = [];   // 대기 (1부 출근대기 → spare1로 2부 첫번째 고정)
  const 대근1부List: string[] = []; // 대근-1부: 1부만 근무 후 귀가
  const 대근2부List: string[] = []; // 대근-2부: 2부만 근무
  const vip1List: string[] = [];   // VIP 1부 전담
  const vip2List: string[] = [];   // VIP 2부 전담
  const vipBothList: string[] = []; // VIP 투근무 (1부+2부)
  const excluded: string[] = [];
  const autoQueue: string[] = [];  // 일반 순번 대기열

  // 클릭 순서 지정된 이름 먼저, 나머지는 roster 순서 fallback
  const orderSet = new Set(statusOrder);
  const order = [...statusOrder.filter(n => names.includes(n)), ...names.filter(n => !orderSet.has(n))];

  for (const name of order) {
    const s = statuses[name] ?? null;
    if (s === "VIP1부")   { vip1List.push(name); }
    else if (s === "VIP2부")    { vip2List.push(name); }
    else if (s === "VIP투근무") { vipBothList.push(name); }
    else if (s === "찾근")  { twoRound.push(name); }
    else if (s === "대기") {
      대기List.push(name);
    } else if (s === "조출") {
      if (조출List.length < 6) 조출List.push(name); else autoQueue.push(name);
    } else if (s === "후출") {
      if (후출List.length < 6) 후출List.push(name); else autoQueue.push(name);
    } else if (EXCLUDED_SET.has(s ?? "")) { excluded.push(name); }
    else {
      // status null(정상근무) — 대근 유형 확인
      const dg = daegeunMap[name];
      if (dg === "1부")           대근1부List.push(name);  // 1부만 출근
      else if (dg === "2부")      대근2부List.push(name);  // 2부만 출근
      else if (dg === "투라운드")  twoRound.push(name);    // 1부+2부 투라운드 대근
      else                        autoQueue.push(name);
    }
  }

  // ── 1부 배치: VIP→ 찾근 → 조출 → 대근1부 → 일반순번 ── (대기자는 1부 미포함)
  const fixed1 = [...vip1List, ...vipBothList, ...twoRound, ...조출List, ...대근1부List];
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
  // 순서: VIP → spare1(1부스페어) → 대근 → 일반순번 → 후출 → 순환보충(찾근+일반1부순번, 1부순서)
  const shift1Regular = autoQueue.slice(0, avail1);
  // vip2List + vipBothList 는 일반 순번과 별개로 2부 앞에 고정
  const vip2Fixed = [...vip2List, ...vipBothList];

  // 순환보충 후보: 찾근 + 일반1부순번 (조출·VIP·대근1부 제외), 1부 순서 그대로
  // → 찾근과 순환보충을 분리하지 않고 1부 흐름 그대로 2부에 재투입
  const circularQueue = [...shift1Regular.slice(0, 2), ...twoRound, ...shift1Regular.slice(2)];

  // 2부에서 normalFor2 + 순환보충이 채워야 할 총 자리 수
  const totalNeed2 = Math.max(0, shift2Size - vip2Fixed.length - spare1.length - 대근2부List.length - 후출List.length);
  const normalFor2 = remaining.slice(0, totalNeed2);
  const spare2fromRemaining = remaining.slice(totalNeed2);

  // 인원 부족 시 circularQueue(찾근+일반순번) 앞번호부터 순환 보충
  const extra2부Count = Math.max(0, totalNeed2 - normalFor2.length);
  const extra2부 = circularQueue.slice(0, extra2부Count);

  // ★ 2부 스페어: remaining 잔여 + 순환보충에서 못 들어간 사람
  const spare2fromShift1 = circularQueue.slice(extra2부Count);
  let spare2 = [...spare2fromRemaining, ...spare2fromShift1];
  let spare2FromTemporaryWork = spare2.filter(n => twoRound.includes(n));

  // 2부 순서: VIP → spare1(1부스페어) → 대근 → 일반순번 → 후출 → 순환보충(찾근포함, 1부순서)
  const shift2: string[] = [
    ...vip2Fixed,
    ...spare1,
    ...대근2부List,
    ...normalFor2,
    ...후출List,
    ...extra2부,
  ];

  // ── 다음날 예상 순번 ──────────────────────────────────
  let nextDayQueue: string[];
  {
    const exclSet2  = new Set(excluded);
    const spare1Set = new Set(spare1);
    const spare2Set = new Set(spare2);

    if (spare2.length > 0) {
      const twoRoundSet = new Set(twoRound);
      const firstSpares = getPrioritySparesFromSaved(spare2, spare2FromTemporaryWork); // 임시 투입 스페어 제외 첫번호
      const firstSpareSet = new Set(firstSpares);
      const restSpares  = spare2.filter(n => !firstSpareSet.has(n)); // 화면용 spare2 순서는 유지, 다음날 큐에서만 우선권 분리
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
      // ★ full rotation(한바퀴 완전 소화): 다음날 첫 2명을 2부스페어로 지정
      if (nextDayQueue.length >= 2) {
        spare2 = nextDayQueue.slice(0, 2);
        spare2FromTemporaryWork = spare2.filter(n => twoRound.includes(n));
      }
    }
  }

  // 대근 인원 통합 목록 (1부·2부·투라운드 — daegeunMap에 등록된 사람)
  const daegeunList = [...대근1부List, ...대근2부List, ...twoRound.filter(n => daegeunMap[n] === "투라운드")];
  return { twoRound, shift1, spare1, shift2, spare2, spare2FromTemporaryWork, excluded, 조출List, 후출List, vip1List, vip2List, vipBothList, nextDayQueue, daegeunList };
}

// ── 배정 엔진: 단부제 ─────────────────────────────
function assignSingle(
  names: string[],
  statuses: Record<string, StatusType>,
  teamSize: number
): DayResult {
  const twoRound: string[] = [];
  const vipFixed: string[] = [];  // VIP (단부제: 앞에 고정)
  const shift1: string[] = [];
  const spare2: string[] = [];
  const excluded: string[] = [];
  const autoQueue: string[] = [];

  for (const name of names) {
    const s = statuses[name] ?? null;
    if (VIP_STATUSES.has(s)) { vipFixed.push(name); }
    else if (s === "찾근") { twoRound.push(name); }
    else if (EXCLUDED_SET.has(s ?? "")) { excluded.push(name); }
    else { autoQueue.push(name); }
  }

  // VIP → 찾근 → 일반 순번
  const avail = Math.max(0, teamSize - vipFixed.length - twoRound.length);
  vipFixed.forEach(n => shift1.push(n));
  twoRound.forEach(n => shift1.push(n));
  autoQueue.forEach((n, i) => {
    if (i < avail) shift1.push(n); else spare2.push(n);
  });
  const spare2FromTemporaryWork = spare2.filter(n => twoRound.includes(n));

  // 다음날 순번: spare2=0 이면 마지막 근무자 다음부터 회전
  let nextDayQueue: string[];
  {
    const exclSet2  = new Set(excluded);
    const spare2Set = new Set(spare2);

    if (spare2.length > 0) {
      const rest  = names.filter(n => !spare2Set.has(n) && !exclSet2.has(n));
      const excls = names.filter(n => exclSet2.has(n));
      const firstSpares = getPrioritySparesFromSaved(spare2, spare2FromTemporaryWork);
      const firstSpareSet = new Set(firstSpares);
      const restSpares = spare2.filter(n => !firstSpareSet.has(n));
      nextDayQueue = [...firstSpares, ...restSpares, ...rest, ...excls];
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

  return { twoRound, shift1, spare1: [], shift2: [], spare2, spare2FromTemporaryWork, excluded, nextDayQueue, daegeunList: [] };
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
  excluded: string[],
  spare2FromTemporaryWork: string[] = []
): string[] {
  const spares = getPrioritySparesFromSaved(spare2, spare2FromTemporaryWork);
  const spareSet = new Set([...spare1, ...spares]);
  const exclSet  = new Set(excluded);

  // ① 오늘 2부스페어: 앞번호부터 (1부스페어는 오늘 2부도 나갔으므로 포함X)
  //    임시 투입 출신 스페어는 다음날 우선권에서만 제외한다.

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
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return parseCalendarOCR((data as any).words as OcrWord[]);
  } finally {
    await worker.terminate();
  }
}

// ── 메인 컴포넌트 ─────────────────────────────────
export default function SchedulePage() {
  const [, setLocation] = useLocation();
  const { user } = useAuth();
  const isAdmin = user?.role === "admin";
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

  // shift1Input: 빈 문자열 → 단부제 자동 판별, 숫자 입력 → 2부제 자동 판별
  const [shift1Input, setShift1Input] = useState<string>("");
  const mode: Mode = shift1Input.trim() !== "" ? "2부제" : "단부제";
  const [totalSize, setTotalSize] = useState<number>(60);
  const shift1Size = shift1Input.trim() !== "" ? Math.max(0, Number(shift1Input) || 0) : 0;
  const shift2Size = Math.max(0, totalSize - shift1Size);
  const singleSize = totalSize;  // 단부제: 총 팀수를 그대로 사용
  // 팀수 설정 잠금 (저장 완료 상태)
  const [teamsLocked, setTeamsLocked] = useState<boolean>(false);

  // 팀수 설정 저장 — ref로 동기 추적한 날짜 기준 저장 (React 상태 비동기 문제 방지)
  function saveTeamSettings() {
    const dateLabel = activeDateLabelRef.current || selectedDate?.dateLabel;
    if (!dateLabel) return;
    if (mode === "2부제" && shift1Size >= totalSize) return;  // 1부 >= 총 팀수 → 저장 불가
    _writeTeamForDate(dateLabel, { mode, totalSize, shift1Size, singleSize: totalSize, locked: true });
    setTeamsLocked(true);
  }
  function unlockTeamSettings() {
    const dateLabel = activeDateLabelRef.current || selectedDate?.dateLabel;
    if (!dateLabel) return;
    _writeTeamForDate(dateLabel, { mode, totalSize, shift1Size, singleSize: totalSize, locked: false });
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
      const raw: Record<string, Record<string, StatusType>> = saved ? JSON.parse(saved) : {};
      // 이름 키의 공백 제거 정규화 + null → "휴무해제" 마이그레이션
      // (resolveStatus 개편 후 null은 "no-op"으로 처리되므로 명시적 휴무해제로 변환)
      const normalized: typeof raw = {};
      for (const [dateKey, statuses] of Object.entries(raw)) {
        const fixedStatuses: Record<string, StatusType> = {};
        for (const [name, st] of Object.entries(statuses)) {
          const normName = normalize(name);
          if (!fixedStatuses[normName]) {
            // null → "휴무해제" 마이그레이션
            fixedStatuses[normName] = (st === null ? "휴무해제" : st) as StatusType;
          }
        }
        normalized[dateKey] = fixedStatuses;
      }
      return normalized;
    } catch { return {}; }
  });
  useEffect(() => {
    localStorage.setItem(DS_KEY, JSON.stringify(dateStatuses));
  }, [dateStatuses, DS_KEY]);

  // 날짜별 찾근/조출/후출 클릭 순서 (화면 표시 및 배치 순서 결정)
  const DSO_KEY = `lotto_dateStatusOrders_${new Date().getFullYear()}`;
  const [dateStatusOrders, setDateStatusOrders] = useState<Record<string, string[]>>(() => {
    try {
      const saved = localStorage.getItem(DSO_KEY);
      return saved ? JSON.parse(saved) : {};
    } catch { return {}; }
  });
  useEffect(() => {
    localStorage.setItem(DSO_KEY, JSON.stringify(dateStatusOrders));
  }, [dateStatusOrders, DSO_KEY]);

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

  const [holidayMap, setHolidayMap] = useState<Record<string, string[]>>(() => {
    try { return JSON.parse(localStorage.getItem(HM_KEY) ?? "{}"); } catch { return {}; }
  });
  const [holidayFileName, setHolidayFileName] = useState<string | null>(() =>
    localStorage.getItem("lotto_holidayFileName")
  );
  useEffect(() => {
    localStorage.setItem(HM_KEY, JSON.stringify(holidayMap));
  }, [holidayMap]);

  // ── 서버에서 휴무 데이터 자동 로드 (마운트 시 1회) ──
  // updatedAt 비교: 서버가 더 최신이면 localStorage 전체 덮어쓰기
  // 동일하거나 서버에 없는 달은 로컬 유지 (merge)
  useEffect(() => {
    const localHMRaw = localStorage.getItem("lotto_holidayMap");
    const localHM: Record<string, string[]> = (() => { try { return JSON.parse(localHMRaw ?? "{}"); } catch { return {}; } })();
    const localTs = localStorage.getItem("lotto_holidayMapUpdatedAt");

    fetch(`/api/holiday-map?_=${Date.now()}`, { cache: "no-store" })
      .then(r => r.json())
      .then((data: { fileName: string; holidayMap: Record<string, string[]>; updatedAt: string | null }) => {
        if (!data.fileName) return;

        const LOCAL_TS_KEY = "lotto_holidayMapUpdatedAt";
        const serverTs = data.updatedAt;

        setHolidayFileName(data.fileName);
        setHolidayMap(prev => {
          const serverMonths = new Set(Object.keys(data.holidayMap).map(k => k.slice(0, 2)));
          const next = { ...prev };
          for (const k of Object.keys(next)) {
            if (serverMonths.has(k.slice(0, 2))) delete next[k];
          }
          const merged = { ...next, ...data.holidayMap };
          localStorage.setItem(HM_KEY, JSON.stringify(merged));
          return merged;
        });
        localStorage.setItem("lotto_holidayFileName", data.fileName);
        if (serverTs) localStorage.setItem(LOCAL_TS_KEY, serverTs);
      })
      .catch(e => { console.error("[HolidaySync] 오류:", e); });
  }, []);

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
        // ── 업로드한 월만 교체, 다른 월 데이터는 유지 ──
        const uploadedMonths = new Set(Object.keys(map).map(k => k.slice(0, 2)));
        setHolidayMap(prev => {
          const next = { ...prev };
          for (const key of Object.keys(next)) {
            if (uploadedMonths.has(key.slice(0, 2))) delete next[key];
          }
          const merged = { ...next, ...map };

          // ── 서버에 저장 (모든 기기에서 공유) ──
          fetch("/api/holiday-map", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ fileName: file.name, holidayMap: merged }),
          })
            .then(r => r.json())
            .then((result: { ok?: boolean; months?: string[]; keyCount?: number; updatedAt?: string; error?: string }) => {
              if (result.error) {
                console.error("[HolidayUpload] 서버 오류:", result.error);
              } else if (result.updatedAt) {
                localStorage.setItem("lotto_holidayMapUpdatedAt", result.updatedAt);
              }
            })
            .catch(e => { console.error("[HolidayUpload] fetch 실패:", e); });

          localStorage.setItem("lotto_holidayMap", JSON.stringify(merged));
          return merged;
        });
        setHolidayFileName(file.name);
        localStorage.setItem("lotto_holidayFileName", file.name);

        // ── 업로드 월의 휴무 상태만 초기화 (다른 월·다른 상태는 유지) ──
        // viewMonth 대신 실제 업로드된 파일의 월을 기준으로 초기화
        const uploadMonthStr = [...uploadedMonths][0] ?? String(parseInt(viewMonth, 10)).padStart(2, "0");
        setDateStatuses(prev => {
          const next: typeof prev = {};
          for (const [dl, statuses] of Object.entries(prev)) {
            const dlMonth = dl.slice(5, 7); // "2026-05-01" → "05"
            const isUploadMonth = dlMonth === uploadMonthStr;
            const cleaned: Record<string, StatusType> = {};
            for (const [name, st] of Object.entries(statuses)) {
              // 업로드 월 날짜의 휴무만 제거, 다른 월·다른 상태 유지
              if (isUploadMonth && st === "휴무") continue;
              cleaned[name] = st;
            }
            if (Object.keys(cleaned).length > 0) next[dl] = cleaned;
          }
          return next;
        });

        // ── 최근 2개월만 유지: 업로드 월 기준 M-1 이전 데이터 삭제 ──
        // 예) 6월 업로드 → cutoff = 2026-05-01 → 5월·6월 유지, 4월 이전 삭제
        {
          const uploadYr = viewYear; // already number
          const uploadMo = parseInt(viewMonth, 10); // 1~12
          const keepFromMo = uploadMo - 1; // 이 달부터 유지 (M-1)
          const cutoffYr = keepFromMo <= 0 ? uploadYr - 1 : uploadYr;
          const cutoffMoNorm = keepFromMo <= 0 ? keepFromMo + 12 : keepFromMo;
          const cutoff = `${cutoffYr}-${String(cutoffMoNorm).padStart(2, "0")}-01`;

          const trimKeys = <T extends Record<string, unknown>>(obj: T): T => {
            const next = { ...obj };
            for (const k of Object.keys(next)) {
              if (k < cutoff) delete next[k];
            }
            return next;
          };

          // ── trimKeys 적용 범위 ──────────────────────────────────────
          // trimKeys는 ISO "YYYY-MM-DD" 형식 키에만 안전하게 동작합니다.
          // "MM.DD (요일)" 형식 키를 사용하는 state에 적용하면
          // "0..." < "2026-..." 비교로 모든 키가 삭제되는 버그가 발생합니다.
          //
          // ISO 키 state (trimKeys 적용 가능):
          //   - assignmentData
          //
          // MM.DD 형식 state (trimKeys 미적용 — 전부 삭제 버그 방지):
          //   - dateStatuses, savedSpare2, dateDaegeun,
          //     overrideStartByDate, dateStatusOrders
          setAssignmentData(prev => trimKeys(prev) as typeof prev);
        }

        const totalPeople = Object.values(map).reduce((s, a) => s + a.length, 0);
        const uploadedMonthList = [...uploadedMonths].sort().join("·");
        alert(`✅ 휴무 엑셀 업로드 완료!\n${dateCount}개 날짜 · 총 ${totalPeople}건\n${uploadedMonthList}월 데이터 갱신 (다른 월 유지)`);
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

  // VIP 날짜별 저장 (localStorage) — members/round는 dateStatuses로 이전, count만 유지
  interface VipEntry { count: number }
  const [vipData, setVipData] = useState<Record<string, VipEntry>>(() => {
    try {
      const saved = localStorage.getItem("lotto_vipData");
      const raw = saved ? JSON.parse(saved) : {};
      // 이전 구조 호환: count만 남기고 나머지 무시
      const clean: Record<string, VipEntry> = {};
      for (const [k, v] of Object.entries(raw)) {
        const entry = v as Record<string, unknown>;
        if (typeof entry.count === "number" && entry.count > 0) {
          clean[k] = { count: entry.count };
        }
      }
      return clean;
    } catch { return {}; }
  });
  useEffect(() => {
    localStorage.setItem("lotto_vipData", JSON.stringify(vipData));
  }, [vipData]);

  const [vipModalOpen, setVipModalOpen] = useState(false);
  const [vipSearch, setVipSearch] = useState("");
  // VIP 인원 선택 모달에서 하위 유형 선택 중인 사람 (이름 → 임시 선택 유형)
  const [vipSubPicking, setVipSubPicking] = useState<string | null>(null);

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

  // ── 날짜별 배정 결과 영구 저장 ──────────────────────────
  const ASSIGNMENT_KEY = `lotto_assignmentData_${new Date().getFullYear()}`;
  const [assignmentData, setAssignmentData] = useState<Record<string, DayResult>>(() => {
    try {
      const saved = localStorage.getItem(`lotto_assignmentData_${new Date().getFullYear()}`);
      return saved ? (JSON.parse(saved) as Record<string, DayResult>) : {};
    } catch { return {}; }
  });
  useEffect(() => {
    localStorage.setItem(ASSIGNMENT_KEY, JSON.stringify(assignmentData));
  }, [assignmentData, ASSIGNMENT_KEY]);

  // ── 배정 시각 기록 (달력 ✓완료 표시 7일 유지용) ────────────────────────────
  // 달력 날짜가 아닌 "배정을 실행한 시각"을 기준으로 7일 필터링
  const ASSIGN_TS_KEY = "lotto_assignmentTimestamps";
  const [assignmentTimestamps, setAssignmentTimestamps] = useState<Record<string, number>>(() => {
    try {
      const saved = localStorage.getItem("lotto_assignmentTimestamps");
      return saved ? JSON.parse(saved) : {};
    } catch { return {}; }
  });
  function recordAssignmentTimestamps(dateLabels: string[]) {
    const now = Date.now();
    setAssignmentTimestamps(prev => {
      const next = { ...prev };
      dateLabels.forEach(dl => { next[dl] = now; });
      localStorage.setItem(ASSIGN_TS_KEY, JSON.stringify(next));
      return next;
    });
  }

  // ── 첫 순번 표시 범위 (배정 완료 시 설정) ──────────────────────────────────
  // "↑ 이름" 힌트를 달력에 표시할 날짜 범위 (MM.DD 형식)
  const ASSIGNED_RANGE_KEY = "lotto_assignedRange";
  const [assignedRange, setAssignedRange] = useState<{ start: string; end: string } | null>(() => {
    try {
      const saved = localStorage.getItem("lotto_assignedRange");
      return saved ? JSON.parse(saved) : null;
    } catch { return null; }
  });
  function persistAssignedRange(v: { start: string; end: string } | null) {
    setAssignedRange(v);
    if (v) localStorage.setItem(ASSIGNED_RANGE_KEY, JSON.stringify(v));
    else localStorage.removeItem(ASSIGNED_RANGE_KEY);
  }

  // 이전날 날짜 레이블 찾기
  const prevDateLabel = useMemo(() => {
    if (!selectedDate) return null;
    // excelDays 우선 탐색
    const idx = excelDays.findIndex(d => d.dateLabel === selectedDate.dateLabel);
    if (idx > 0) return excelDays[idx - 1].dateLabel;
    // viewDays 폴백
    const vi = viewDays.findIndex(d => d.dateLabel === selectedDate.dateLabel);
    if (vi > 0) return viewDays[vi - 1].dateLabel;
    // 월 경계: excelDays/viewDays 첫 번째 날인 경우 → 날짜 산술로 전날 직접 계산
    const match = selectedDate.dateLabel.match(/^(\d{2})\.(\d{2})/);
    if (match) {
      const prev = new Date(viewYear, parseInt(match[1], 10) - 1, parseInt(match[2], 10) - 1);
      return makeDateKey(prev);
    }
    return null;
  }, [selectedDate, excelDays, viewDays, viewYear]);

  // 오늘 첫번호 힌트 = 전날 우선 스페어[0] (기존 데이터는 savedSpare2 fallback)
  const todayFirstHint = prevDateLabel
    ? (getPrioritySpares(assignmentData[prevDateLabel])[0] ?? savedSpare2[prevDateLabel]?.[0] ?? null)
    : null;

  // 현재 선택 날짜 키 (e.g. "04.01 (수)")
  const currentDateKey = selectedDate?.dateLabel ?? "";

  const currentVip: VipEntry = useMemo(
    () => vipData[currentDateKey] ?? { count: 0 },
    [vipData, currentDateKey]
  );
  function setCurrentVip(v: Partial<VipEntry>) {
    if (!currentDateKey) return;
    setVipData(prev => {
      const cur = prev[currentDateKey] ?? { count: 0 };
      const next = { ...cur, ...v };
      if (next.count === 0) {
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

  // dateStatuses에서 현재 날짜의 VIP 멤버 추출
  const currentVipMembers = useMemo(
    () => Object.entries(manualStatuses)
      .filter(([, s]) => VIP_STATUSES.has(s))
      .map(([name, s]) => ({ name, type: s as "VIP1부" | "VIP2부" | "VIP투근무" })),
    [manualStatuses]
  );

  // 현재 날짜의 대근 맵 (derived)
  const currentDaegeun = useMemo(
    () => dateDaegeun[currentDateKey] ?? {},
    [dateDaegeun, currentDateKey]
  );

  // 일주일/3일 생성 시 기존 배정 덮어쓰기 확인 모달 (0=닫힘, 3=3일, 7=7일)
  const [weekForceConfirm, setWeekForceConfirm] = useState(0);

  // 칩 드래그 상태 (드래그 중 인덱스 추적)
  const chipDragRef = useRef<{ fromIdx: number | null; didDrag: boolean }>({ fromIdx: null, didDrag: false });
  const [chipDragOver, setChipDragOver] = useState<number | null>(null);

  function reorderSelectedChips(fromIdx: number, toIdx: number, ordered: string[]) {
    if (fromIdx === toIdx || !currentDateKey) return;
    const next = [...ordered];
    const [moved] = next.splice(fromIdx, 1);
    next.splice(toIdx, 0, moved);
    setDateStatusOrders(prev => ({ ...prev, [currentDateKey]: next }));
  }

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
  // 배정하기 후 저장 전 임시 결과 (저장 버튼 누르기 전까지 dayResult를 덮어쓰지 않음)
  const [pendingResult, setPendingResult] = useState<DayResult | null>(null);
  // 화면에 표시할 결과 = 임시 결과 우선, 없으면 저장된 결과
  const displayResult = pendingResult ?? dayResult;
  const [weekly, setWeekly] = useState<{ day: string; result: DayResult; skipped?: boolean }[]>([]);
  // 재계산 완료 메시지
  const [recalcMessage, setRecalcMessage] = useState<string | null>(null);
  // 주간 근무표 날짜별 개별 토글 (기본: 요약 보기)
  const [expandedDays, setExpandedDays] = useState<Set<string>>(new Set());
  function toggleDayExpand(day: string) {
    setExpandedDays(prev => {
      const next = new Set(prev);
      if (next.has(day)) next.delete(day); else next.add(day);
      return next;
    });
  }

  // 통합 선택 모달 — StatusType | "VIP" 지원
  const [modalStatus, setModalStatus] = useState<StatusType | "VIP" | null>(null);
  const [modalSearch, setModalSearch] = useState("");
  const [showFullList, setShowFullList] = useState(false);

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
      const raw = saved ? (JSON.parse(saved) as PersonData[]) : ROSTER;
      return normalizeRoster(raw);
    } catch { return normalizeRoster(ROSTER); }
  });
  // 저장 (localStorage + 서버)
  const rosterInitRef = useRef(true);
  useEffect(() => {
    localStorage.setItem("lotto_customRoster", JSON.stringify(customRoster));
    // 초기 마운트 시 로드는 서버에서 받으므로 서버 저장 건너뜀
    if (rosterInitRef.current) { rosterInitRef.current = false; return; }
    // 변경 즉시 로컬 타임스탬프 업데이트 (서버 저장 성공 여부 무관하게 로컬 우선 보장)
    const optimisticNow = new Date().toISOString();
    localStorage.setItem("lotto_rosterUpdatedAt", optimisticNow);
    // 순번표가 변경되면 서버에도 저장
    fetch("/api/roster", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ roster: customRoster }),
    })
      .then(r => r.json())
      .then((res: { ok?: boolean; count?: number; updatedAt?: string }) => {
        if (res.updatedAt) localStorage.setItem("lotto_rosterUpdatedAt", res.updatedAt);
      })
      .catch(e => console.error("[RosterSync] 서버 저장 오류:", e));
  }, [customRoster]);
  // 이 기기 데이터로 서버 강제 덮어쓰기
  function forceUploadToServer() {
    if (!confirm("현재 이 기기의 순번표를 서버에 강제 저장합니다.\n다른 기기의 데이터는 덮어씌워집니다. 계속할까요?")) return;
    const now = new Date().toISOString();
    fetch("/api/roster", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ roster: customRoster }),
    })
      .then(r => r.json())
      .then((res: { ok?: boolean; updatedAt?: string }) => {
        if (res.updatedAt) localStorage.setItem("lotto_rosterUpdatedAt", res.updatedAt);
        else localStorage.setItem("lotto_rosterUpdatedAt", now);
        alert(`✅ 서버 덮어쓰기 완료 (${customRoster.length}명)`);
      })
      .catch(e => { console.error("[RosterSync] 강제 덮어쓰기 오류:", e); alert("⚠️ 서버 저장 실패"); });
  }

  // 마운트 시 서버 동기화 — 타임스탬프 비교: 더 최신 쪽 데이터 사용
  useEffect(() => {
    fetch(`/api/roster?_=${Date.now()}`, { cache: "no-store" })
      .then(r => r.json())
      .then((data: { roster: PersonData[]; updatedAt: string | null }) => {
        const serverRoster = Array.isArray(data.roster) ? normalizeRoster(data.roster as PersonData[]) : [];
        const serverTime = data.updatedAt ? new Date(data.updatedAt).getTime() : 0;
        const localTimeStr = localStorage.getItem("lotto_rosterUpdatedAt");
        const localTime = localTimeStr ? new Date(localTimeStr).getTime() : 0;

        if (serverRoster.length === 0) {
          // 서버가 비어있으면 로컬 데이터 올리기
          const localRoster = (() => { try { return normalizeRoster(JSON.parse(localStorage.getItem("lotto_customRoster") ?? "[]") as PersonData[]); } catch { return []; } })();
          if (localRoster.length > 0) {
            fetch("/api/roster", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({ roster: localRoster }),
            }).then(r => r.json()).then((res: { updatedAt?: string }) => {
              if (res.updatedAt) localStorage.setItem("lotto_rosterUpdatedAt", res.updatedAt);
            }).catch(e => console.error("[RosterSync] 초기 업로드 오류:", e));
          }
          return;
        }

        if (localTime > serverTime) {
          // 로컬이 더 최신 → 로컬 유지 + 서버에 업로드 (다른 기기와 동기화)
          const localRoster = (() => { try { return normalizeRoster(JSON.parse(localStorage.getItem("lotto_customRoster") ?? "[]") as PersonData[]); } catch { return []; } })();
          if (localRoster.length > 0) {
            fetch("/api/roster", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({ roster: localRoster }),
            }).then(r => r.json()).then((res: { updatedAt?: string }) => {
              if (res.updatedAt) localStorage.setItem("lotto_rosterUpdatedAt", res.updatedAt);
            }).catch(e => console.error("[RosterSync] 로컬→서버 업로드 오류:", e));
          }
          return;
        }

        // 서버가 더 최신이거나 로컬 타임스탬프 없음 → 서버 데이터 적용
        if (data.updatedAt) localStorage.setItem("lotto_rosterUpdatedAt", data.updatedAt);
        rosterInitRef.current = true;
        setCustomRoster(serverRoster);
      })
      .catch(e => console.error("[RosterSync] 서버 조회 오류:", e));
  }, []);

  // 조 순 정렬
  const sortedCustomRoster = useMemo(() =>
    [...customRoster].sort((a, b) => a.조 !== b.조 ? a.조 - b.조 : a.no - b.no),
    [customRoster]
  );
  // 이름 → PersonData 맵 (키는 normalize로 정규화, 조회 시 normalize(name) 사용)
  const customRosterMap = useMemo(() =>
    Object.fromEntries(customRoster.map(p => [normalize(p.name), p])),
    [customRoster]
  );
  // lookup helper — 조회 시 항상 normalize 경유
  const getRosterPerson = (name: string) => customRosterMap[normalize(name)];

  const getGroup = (name: string): GroupType =>
    customRosterMap[normalize(name)]?.group
    ?? NAME_GROUP_NORMALIZED[normalize(name)]
    ?? "하우스";

  // ── 순번표 편집 모달 ──
  const [rosterEditorOpen, setRosterEditorOpen] = useState(false);
  const [rosterEditorSearch, setRosterEditorSearch] = useState("");
  const rosterImportRef = useRef<HTMLInputElement>(null);
  const resultRef = useRef<HTMLDivElement>(null);
  const [rosterForm, setRosterForm] = useState<{ mode: "add"|"edit"; orig?: PersonData; name: string; 조: 1|2|3|4; group: GroupType } | null>(null);

  // ── 통합 상태 계산 헬퍼 ─────────────────────────────────────────────────────
  // 모든 배정/표시에서 동일 로직 사용 (holidayMap 항상 우선, dateStatuses는 override)
  function resolveStatus(
    name: string,
    dateKey: string,
    dayIdx: number,
    savedDay: Record<string, StatusType>,
    daegeunMap: Record<string, DaegeunType>
  ): StatusType {
    if (sickLeave[name]) return "병가";
    const override = savedDay[name] as StatusType | undefined;
    // 비휴무 명시 override (당번/조출/후출/찾근 등) → 그대로 반환
    if (override && override !== "휴무" && override !== "휴무해제") return override;
    // 대근 지정: 엑셀 휴무보다 우선 (주말반 등 대근 출근 허용)
    if (daegeunMap[name]) return null;
    // 휴무 판단: 해제 > 추가 > 엑셀 > 자동규칙
    if (override === "휴무해제") return null;
    if (override === "휴무") return "휴무";
    const dk = dateKey.slice(0, 5);
    if (new Set((holidayMap[dk] ?? []).map(n => normalize(n))).has(normalize(name))) return "휴무";
    const person = getRosterPerson(name);
    if (person && isAutoOff(person.group, dayIdx)) return "휴무";
    return null;
  }

  // 현재 날짜의 유효 상태 반환 (UI 표시용)
  function effectiveStatus(name: string, dayIdx: number = dayOfWeek): StatusType {
    const dg = currentDaegeun[name];
    // 대근 UI 표시: 투라운드 → 찾근, 1부/2부 → null (자동휴무 덮어쓰기)
    if (dg === "투라운드") return "찾근";
    if (dg === "1부" || dg === "2부") return null;
    return resolveStatus(name, currentDateKey, dayIdx, manualStatuses, currentDaegeun);
  }

  // ── 통일된 가용인원 계산 ────────────────────────────────────────────────────
  // 가용인원 = 총인원 - EXCLUDED_SET(휴무·당번·병가·하우스) 해당 인원 수
  // resolveStatus 경유 → 엑셀휴무·수동휴무·자동휴무·당번·병가·대근 모두 반영
  function calcAvailable(dateLabel: string, dayIdx: number): number {
    const baseNames = names.length > 0 ? names : sortedCustomRoster.map(p => p.name);
    if (baseNames.length === 0) return 0;
    const savedDay = dateStatuses[dateLabel] ?? {};
    const dgMap = dateDaegeun[dateLabel] ?? {};
    return baseNames.filter(name => {
      const st = resolveStatus(name, dateLabel, dayIdx, savedDay, dgMap);
      return !EXCLUDED_SET.has(st ?? "");
    }).length;
  }

  // ── displayDays: 통일된 가용인원 계산 (resolveStatus 기준, 엑셀 가용인원 값 미사용) ──
  // 달력 카드·하단 패널 모두 동일 계산식 사용
  const displayDays = useMemo(() => {
    if (customRoster.length === 0) return viewDays.map(d => ({ ...d, 가용인원: 0 }));
    return viewDays.map(d => ({
      ...d,
      가용인원: calcAvailable(d.dateLabel, d.dayIdx),
    }));
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [viewDays, names, sortedCustomRoster, dateStatuses, dateDaegeun, sickLeave, holidayMap, customRoster]);

  // 수동 상태 삭제 (키 자체 제거 → day-of-week 로직이 다시 적용됨)
  function clearStatus(name: string) {
    setManualStatuses(prev => {
      if (!(name in prev)) return prev;
      const next = { ...prev };
      delete next[name];
      return next;
    });
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
    // 모든 수동 상태: 클릭 순서 추적 (dateStatusOrders 업데이트)
    // 찾근/조출/후출 + 휴무/병가/당번/대기 모두 동일하게 적용
    if (currentDateKey) {
      const cur = effectiveStatus(name);
      if (cur === btn) {
        // 취소 → 순서 배열에서 제거
        setDateStatusOrders(prev => ({
          ...prev,
          [currentDateKey]: (prev[currentDateKey] ?? []).filter(n => n !== name),
        }));
      } else {
        // 추가 → 순서 배열 끝에 push (중복 방지)
        setDateStatusOrders(prev => ({
          ...prev,
          [currentDateKey]: [...(prev[currentDateKey] ?? []).filter(n => n !== name), name],
        }));
      }
    }

    // 병가 외 상태 토글 (기존 로직)
    setManualStatuses((prev) => {
      const cur = effectiveStatus(name);
      if (cur === btn && name in prev) {
        // 수동 override가 있는 경우: 삭제
        // 단, 엑셀 휴무인 사람은 삭제하면 holidayMap이 다시 "휴무"로 복구되므로
        // "휴무해제"로 명시적 해제
        const dk5 = currentDateKey.slice(0, 5);
        const inHolidayMap = new Set((holidayMap[dk5] ?? []).map(n => normalize(n))).has(normalize(name));
        if (btn === "휴무" && inHolidayMap) {
          return { ...prev, [name]: "휴무해제" };
        }
        const next = { ...prev };
        delete next[name];
        return next;
      } else if (cur === btn && !(name in prev)) {
        // 수동 override 없이 "휴무"인 경우 (holidayMap/autoOff 기반)
        // → "휴무해제"로 명시적 해제 (null은 resolveStatus에서 무시됨)
        return { ...prev, [name]: "휴무해제" };
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
      // 마이그레이션: 저장된 mode 기준으로 shift1Input 복원
      if (saved.mode === "2부제") {
        setShift1Input(String(saved.shift1Size ?? 35));
        setTotalSize(saved.totalSize ?? 60);
      } else {
        // 단부제: 기존 singleSize를 새 totalSize로 사용 (마이그레이션)
        setShift1Input("");
        setTotalSize(saved.singleSize ?? saved.totalSize ?? 60);
      }
      setTeamsLocked(saved.locked ?? false);
    } else {
      // 기본값으로 리셋
      setShift1Input("");
      setTotalSize(60);
      setTeamsLocked(false);
    }

    // 날짜 이동 시 미저장 임시 결과 항상 초기화
    setPendingResult(null);
    // dayResult를 선택 날짜 기준으로 항상 초기화/복원 (다른 날짜 결과가 남지 않도록)
    setDayResult(assignmentData[day.dateLabel] ?? null);
    // 이미 배정된 날짜면 weekly 초기화 → 1일 배정 결과 화면으로 전환
    if (assignmentData[day.dateLabel]) {
      setWeekly([]);
    }
  }

  // ── 첫번호 지정 (세션 전용 — localStorage 저장 X, 다음날 이어지지 않음) ───
  const _now = new Date();
  const TODAY_KEY = `lotto_queueStart_${_now.getFullYear()}-${String(_now.getMonth() + 1).padStart(2, "0")}-${String(_now.getDate()).padStart(2, "0")}`;
  const [queueStartName, setQueueStartName] = useState<string | null>(() =>
    localStorage.getItem(TODAY_KEY)
  );
  // 날짜별 수동 첫번호 override: { "2026-04-19": "김혜민", ... }
  const OVERRIDE_KEY = `lotto_overrideStart_${new Date().getFullYear()}`;
  const [overrideStartByDate, setOverrideStartByDate] = useState<Record<string, string>>(() => {
    try { return JSON.parse(localStorage.getItem(`lotto_overrideStart_${new Date().getFullYear()}`) ?? "{}") ?? {}; }
    catch { return {}; }
  });
  const [queueModal, setQueueModal] = useState<"ask" | "pick" | null>(null);
  const [queuePickSearch, setQueuePickSearch] = useState("");
  // visualViewport 높이 추적 (키보드 올라올 때 모달 재계산용)
  const [vvHeight, setVvHeight] = useState(() =>
    (typeof window !== "undefined" && window.visualViewport?.height) || window.innerHeight
  );
  useEffect(() => {
    const vv = window.visualViewport;
    if (!vv) return;
    const handler = () => setVvHeight(vv.height);
    vv.addEventListener("resize", handler);
    vv.addEventListener("scroll", handler);
    return () => { vv.removeEventListener("resize", handler); vv.removeEventListener("scroll", handler); };
  }, []);
  const queueListRef = useRef<HTMLDivElement>(null);
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
    const norm = (s: string) => s.trim();
    const idx = base.findIndex((n) => norm(n) === norm(startName));
    if (idx === -1) {
      console.warn("[rotateNames] startName not found in base:", startName);
      return base;
    }
    if (idx === 0) return base;
    return [...base.slice(idx), ...base.slice(0, idx)];
  }

  // 실제로 names 적용 (회전 포함)
  // saveOverride: true(기본) → 수동 선택, overrideStartByDate 저장
  //              false        → 자동 이어짐 적용, override 저장하지 않음
  function applyRoster(startName: string | null, { saveOverride = true }: { saveOverride?: boolean } = {}) {
    const base = sortedCustomRoster.map((p) => p.name);
    const rotated = rotateNames(base, startName);
    setNames(rotated);
    setRosterLoaded(true);
    setDayResult(null);
    setPendingResult(null);
    setWeekly([]);
    // 수동 선택일 때만 override 저장
    if (saveOverride && startName && currentDateKey) {
      setOverrideStartByDate(prev => {
        const next = { ...prev, [currentDateKey]: startName };
        localStorage.setItem(OVERRIDE_KEY, JSON.stringify(next));
        return next;
      });
    }
    setTimeout(() => {
      resultRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
    }, 100);
  }

  // 홈 이동 후 복귀 시 자동 재적용
  // — customRoster가 있고 queueStartName이 저장돼 있으면 바로 카드(배정) 화면으로
  useEffect(() => {
    if (sortedCustomRoster.length > 0 && queueStartName && names.length === 0) {
      const base = sortedCustomRoster.map((p) => p.name);
      setNames(rotateNames(base, queueStartName));
      setRosterLoaded(true);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // 순번표 불러오기 → 항상 ask 모달 열기 (시각적 피드백 보장)
  function loadRoster() {
    setQueueModalPos({ x: 0, y: 0 });
    setQueueModal("ask");
  }

  // 직접 입력으로 다음 단계
  function confirmNames() {
    const parsed = nameText.split("\n").map((n) => n.trim()).filter(Boolean);
    if (!parsed.length) return;
    setNames(parsed);
    setRosterLoaded(false);
    setDayResult(null);
    setPendingResult(null);
    setWeekly([]);
    setTimeout(() => {
      resultRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
    }, 100);
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
  function movePersonInJo(p: PersonData, direction: "up" | "down") {
    const sameJo = [...customRoster]
      .filter(x => x.조 === p.조)
      .sort((a, b) => a.no - b.no);
    const idx = sameJo.findIndex(x => x.name === p.name);
    const swapIdx = direction === "up" ? idx - 1 : idx + 1;
    if (swapIdx < 0 || swapIdx >= sameJo.length) return;
    const swapWith = sameJo[swapIdx];
    setCustomRoster(prev => prev.map(x => {
      if (x.name === p.name) return { ...x, no: swapWith.no };
      if (x.name === swapWith.name) return { ...x, no: p.no };
      return x;
    }));
  }
  function savePerson() {
    if (!rosterForm) return;
    const trimName = normalize(rosterForm.name); // 공백 제거 정규화 적용
    if (!trimName) { alert("이름을 입력하세요."); return; }
    if (rosterForm.mode === "add") {
      // 같은 이름 중복 체크
      if (customRoster.some(x => normalize(x.name) === trimName)) {
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
      const origNorm = normalize(rosterForm.orig?.name ?? "");
      if (trimName !== origNorm && customRoster.some(x => normalize(x.name) === trimName)) {
        alert("이미 같은 이름이 있습니다."); return;
      }
      setCustomRoster(prev => prev.map(x =>
        normalize(x.name) === origNorm
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
        const normalizedData = normalizeRoster(data);
        setCustomRoster(normalizedData);
        alert(`✅ ${data.length}명 복원 완료 (이름 공백 정규화 적용)`);
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

  // ── 현재 날짜 기준 실제 순번 ──────────────────────────
  // 현재 날짜 이전 마지막 저장된 날의 spare2[0] 기준으로 names를 rotate
  // (Day1: firstStarter 기준, Day2+: 전날 spare2[0] 체인 기준)
  // getStartNameForDate: override → spare2 체인 → queueStartName 순 우선
  // excelDays 의존 제거 → 날짜 계산으로 하루씩 거슬러 올라감 (달 바뀌어도 동작)
  function getStartNameForDate(dateLabel: string): string | null {
    if (!dateLabel) return queueStartName;
    // 1. 수동 override 최우선
    if (overrideStartByDate[dateLabel]) return overrideStartByDate[dateLabel];
    // 2. dateLabel 파싱 후 하루씩 거슬러 올라가며 spare2 탐색
    const match = dateLabel.match(/^(\d{2})\.(\d{2})/);
    if (match) {
      const mm = parseInt(match[1], 10);
      const dd = parseInt(match[2], 10);
      let cur = new Date(viewYear, mm - 1, dd);
      for (let i = 0; i < 400; i++) {
        cur = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate() - 1);
        const dl = makeDateKey(cur);
        const result = assignmentData[dl];
        const prioritySpares = result ? getPrioritySpares(result) : getPrioritySparesFromSaved(savedSpare2[dl]);
        if (prioritySpares.length > 0) return prioritySpares[0];
      }
    }
    // 3. 기본 첫번호
    return queueStartName;
  }

  const effectiveNames = useMemo(() => {
    if (names.length === 0) return [];
    const startName = getStartNameForDate(currentDateKey);
    if (!startName) return names;
    return rotateNames([...names], startName);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [names, currentDateKey, overrideStartByDate, viewYear, savedSpare2, assignmentData, queueStartName]);

  // ── 실시간 배정 미리보기 ──────────────────────────
  // generateWeek/recalculateFrom과 동일한 패턴: dateStatuses → isAutoOff 순으로 적용
  const livePreview = useMemo(() => {
    if (effectiveNames.length === 0) return null;
    const savedDay = currentDateKey ? (dateStatuses[currentDateKey] ?? {}) : {};
    const dayIdx = selectedDate?.dayIdx ?? dayOfWeek;
    // assign()과 동일한 검증 로직 적용 → 배정 전 미리보기와 배정 결과가 항상 일치
    const statuses: Record<string, StatusType> = {};
    effectiveNames.forEach((n) => {
      statuses[n] = resolveStatus(n, currentDateKey, dayIdx, savedDay, currentDaegeun);
    });
    const timingStatuses = new Set<StatusType>(["찾근", "조출", "후출"]);
    const baseSavedDay: Record<string, StatusType> = {};
    for (const [name, st] of Object.entries(savedDay)) {
      if (!timingStatuses.has(st)) baseSavedDay[name] = st;
    }
    const baseStatuses: Record<string, StatusType> = {};
    effectiveNames.forEach((n) => {
      baseStatuses[n] = resolveStatus(n, currentDateKey, dayIdx, baseSavedDay, currentDaegeun);
    });
    // 번호 유무 판정용 baseResult: 타이밍 상태(조출/후출/찾근)로 인한 순서 클릭은 제외하고
    // 나머지 수동 순서(휴무해제·당번 등)는 그대로 반영 → 실제 배정과 동일한 기준으로 판단
    const baseStatusOrder = (dateStatusOrders[currentDateKey] ?? [])
      .filter(n => !timingStatuses.has(savedDay[n] as StatusType));
    const baseResult = mode === "2부제"
      ? assignDouble(effectiveNames, baseStatuses, shift1Size, shift2Size, currentDaegeun, baseStatusOrder)
      : assignSingle(effectiveNames, baseStatuses, singleSize);
    const baseShift1Set = new Set(baseResult.shift1);
    const baseShift2Set = new Set(baseResult.shift2);
    const validatedStatuses = { ...statuses };
    effectiveNames.forEach((name) => {
      const st = statuses[name];
      if (st === "찾근") {
        const invalid = mode === "2부제" ? baseShift2Set.has(name) : baseShift1Set.has(name);
        if (invalid) validatedStatuses[name] = baseStatuses[name];
      } else if (st === "조출" || st === "후출") {
        // 기본 상태가 제외(휴무/당번 등)면 명시적 투입 → 검증 통과
        if (!EXCLUDED_SET.has(baseStatuses[name] ?? "")) {
          const hasOriginalNumber = mode === "2부제"
            ? baseShift1Set.has(name) || baseShift2Set.has(name)
            : baseShift1Set.has(name);
          if (!hasOriginalNumber) validatedStatuses[name] = baseStatuses[name];
        }
      }
    });
    return mode === "2부제"
      ? assignDouble(effectiveNames, validatedStatuses, shift1Size, shift2Size, currentDaegeun, dateStatusOrders[currentDateKey] ?? [])
      : assignSingle(effectiveNames, validatedStatuses, singleSize);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [effectiveNames, dateStatuses, currentDateKey, selectedDate, dayOfWeek, customRosterMap, currentDaegeun, mode, shift1Size, shift2Size, singleSize, dateStatusOrders, holidayMap, sickLeave]);

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

  function openStatusPicker(st: StatusType | "VIP") {
    if (names.length === 0) {
      const base = sortedCustomRoster.map((p) => p.name);
      setNames(rotateNames(base, queueStartName));
      setRosterLoaded(true);
    }
    setModalStatus(st);
    setModalSearch("");
    setVipSubPicking(null);
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

  // ── 공통: currentNames 하루 전진 (컴포넌트 레벨) ──
  function advanceNames(result: DayResult, curNames: string[]): string[] {
    const nextSpares = getPrioritySpares(result);
    const nextSpareSet = new Set(nextSpares);
    let nextRest = curNames.filter(n => !nextSpareSet.has(n));
    if (nextSpares.length === 0) {
      const todayLast = (mode === "2부제" ? result.shift2 : result.shift1).at(-1);
      if (todayLast) {
        const li = nextRest.indexOf(todayLast);
        if (li >= 0 && nextRest.length > 1) {
          const startAt = (li + 1) % nextRest.length;
          nextRest = [...nextRest.slice(startAt), ...nextRest.slice(0, startAt)];
        }
      }
    }
    return [...nextSpares, ...nextRest];
  }

  // ── 특정 날짜 이후 연속 재계산 ──
  // startDateLabel 다음날부터, assignmentData가 있는 날짜까지만 순차 재계산
  function recalculateFrom(
    startDateLabel: string,
    baseAssignment: Record<string, DayResult>,
    baseSpare2: Record<string, string[]>
  ): { updatedAssignment: Record<string, DayResult>; updatedSpare2: Record<string, string[]>; count: number } {
    const startIdx = excelDays.findIndex(d => d.dateLabel === startDateLabel);
    if (startIdx < 0) return { updatedAssignment: baseAssignment, updatedSpare2: baseSpare2, count: 0 };

    // 시작 기준: startDate 의 우선 스페어[0]
    const startSpares = baseAssignment[startDateLabel]
      ? getPrioritySpares(baseAssignment[startDateLabel])
      : getPrioritySparesFromSaved(baseSpare2[startDateLabel]);
    if (startSpares.length === 0) return { updatedAssignment: baseAssignment, updatedSpare2: baseSpare2, count: 0 };

    let currentNames = rotateNames([...names], startSpares[0]);
    const updatedAssignment = { ...baseAssignment };
    const updatedSpare2 = { ...baseSpare2 };
    let count = 0;

    for (let i = startIdx + 1; i < excelDays.length; i++) {
      const day = excelDays[i];
      // assignmentData가 없는 날짜에서 멈춤 (미래 새 생성 금지)
      if (!updatedAssignment[day.dateLabel]) break;

      const savedDay = dateStatuses[day.dateLabel] ?? {};
      const statuses: Record<string, StatusType> = {};
      const dgMap = dateDaegeun[day.dateLabel] ?? {};
      currentNames.forEach((n) => {
        statuses[n] = resolveStatus(n, day.dateLabel, day.dayIdx, savedDay, dgMap);
      });

      const s1 = shift1Size, s2 = shift2Size, ss = singleSize;

      const result = mode === "2부제"
        ? assignDouble(currentNames, statuses, s1, s2, dgMap, dateStatusOrders[day.dateLabel] ?? [])
        : assignSingle(currentNames, statuses, ss);

      updatedAssignment[day.dateLabel] = result;
      if (result.spare2.length > 0) updatedSpare2[day.dateLabel] = result.spare2;

      currentNames = advanceNames(result, currentNames);
      count++;
    }

    return { updatedAssignment, updatedSpare2, count };
  }

  function assign() {
    const en = effectiveNames.length > 0 ? effectiveNames : names;
    // generateWeek/recalculateFrom과 동일한 패턴으로 통일
    const savedDay = currentDateKey ? (dateStatuses[currentDateKey] ?? {}) : {};
    const dayIdx = selectedDate?.dayIdx ?? dayOfWeek;
    const statuses: Record<string, StatusType> = {};
    en.forEach((n) => {
      statuses[n] = resolveStatus(n, currentDateKey, dayIdx, savedDay, currentDaegeun);
    });
    const timingStatuses = new Set<StatusType>(["찾근", "조출", "후출"]);
    const baseSavedDay: Record<string, StatusType> = {};
    for (const [name, st] of Object.entries(savedDay)) {
      if (!timingStatuses.has(st)) baseSavedDay[name] = st;
    }
    const baseStatuses: Record<string, StatusType> = {};
    en.forEach((n) => {
      baseStatuses[n] = resolveStatus(n, currentDateKey, dayIdx, baseSavedDay, currentDaegeun);
    });
    // 번호 유무 판정용 baseResult: 타이밍 상태(조출/후출/찾근)로 인한 순서 클릭은 제외하고
    // 나머지 수동 순서(휴무해제·당번 등)는 그대로 반영 → livePreview와 동일한 기준 (sickLeave도 반영)
    const baseStatusOrder = (dateStatusOrders[currentDateKey] ?? [])
      .filter(n => !timingStatuses.has(savedDay[n] as StatusType));
    const baseResult = mode === "2부제"
      ? assignDouble(en, baseStatuses, shift1Size, shift2Size, currentDaegeun, baseStatusOrder)
      : assignSingle(en, baseStatuses, singleSize);
    const baseShift1Set = new Set(baseResult.shift1);
    const baseShift2Set = new Set(baseResult.shift2);
    const validatedStatuses = { ...statuses };
    const invalidStatusReasons: Record<string, string> = {};
    en.forEach((name) => {
      const st = statuses[name];
      if (st === "찾근") {
        const invalid = mode === "2부제" ? baseShift2Set.has(name) : baseShift1Set.has(name);
        if (invalid) {
          validatedStatuses[name] = baseStatuses[name];
          invalidStatusReasons[name] = mode === "2부제" ? "투번호 옴" : "번호 옴";
        }
      } else if (st === "조출" || st === "후출") {
        // 기본 상태가 제외(휴무/당번 등)면 명시적 투입 → 검증 통과 (livePreview와 동일)
        if (!EXCLUDED_SET.has(baseStatuses[name] ?? "")) {
          const hasOriginalNumber = mode === "2부제"
            ? baseShift1Set.has(name) || baseShift2Set.has(name)
            : baseShift1Set.has(name);
          if (!hasOriginalNumber) {
            validatedStatuses[name] = baseStatuses[name];
            invalidStatusReasons[name] = "번호 안옴";
          }
        }
      }
    });
    const invalidNameSet = new Set(Object.keys(invalidStatusReasons));
    const validatedStatusOrder = (dateStatusOrders[currentDateKey] ?? []).filter(n => !invalidNameSet.has(n));
    const result = mode === "2부제"
      ? assignDouble(en, validatedStatuses, shift1Size, shift2Size, currentDaegeun, validatedStatusOrder)
      : assignSingle(en, validatedStatuses, singleSize);
    if (Object.keys(invalidStatusReasons).length > 0) {
      result.invalidStatusReasons = invalidStatusReasons;
    }
    // 저장하지 않고 임시 결과만 보여줌 — 저장은 saveAssignment()에서
    setPendingResult(result);
    setWeekly([]);
    setTimeout(() => {
      resultRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
    }, 80);
  }

  function saveAssignment() {
    if (!pendingResult || !currentDateKey) return;

    setDayResult(pendingResult);

    // 해당 날짜만 확정 저장 — 이후 날짜는 변경하지 않음
    setAssignmentData(prev => ({ ...prev, [currentDateKey]: pendingResult }));
    if (pendingResult.spare2.length > 0) {
      setSavedSpare2(prev => ({ ...prev, [currentDateKey]: pendingResult.spare2 }));
    }
    setPendingResult(null);

    // 다음 날짜 override 자동 정리:
    // 오늘 spare2[0]이 체인으로 이어지므로, 다음 날에 수동 override가 남아 있으면
    // getStartNameForDate가 spare2 결과 대신 이전 값을 반환해 달력과 컷 요약이 불일치함
    const curIdx = excelDays.findIndex(d => d.dateLabel === currentDateKey);
    if (curIdx >= 0 && curIdx + 1 < excelDays.length) {
      const nextDateKey = excelDays[curIdx + 1].dateLabel;
      setOverrideStartByDate(prev => {
        if (!prev[nextDateKey]) return prev;
        const next = { ...prev };
        delete next[nextDateKey];
        localStorage.setItem(OVERRIDE_KEY, JSON.stringify(next));
        return next;
      });
      // ↑ 힌트 표시 범위: 배정 완료 다음 날 1일만
      persistAssignedRange({ start: nextDateKey.slice(0, 5), end: nextDateKey.slice(0, 5) });
    }
    // ✓완료 표시용 배정 시각 기록
    recordAssignmentTimestamps([currentDateKey]);
  }

  function saveAndRecalculate() {
    if (!pendingResult || !currentDateKey) return;

    setDayResult(pendingResult);

    const newAssignment = { ...assignmentData, [currentDateKey]: pendingResult };
    const newSpare2 = pendingResult.spare2.length > 0
      ? { ...savedSpare2, [currentDateKey]: pendingResult.spare2 }
      : { ...savedSpare2 };

    // 해당 날짜 저장 후 이후 날짜 연쇄 재계산
    const { updatedAssignment, updatedSpare2, count } = recalculateFrom(currentDateKey, newAssignment, newSpare2);
    setAssignmentData(updatedAssignment);
    setSavedSpare2(updatedSpare2);
    setPendingResult(null);

    // 재계산된 이후 날짜들의 override를 모두 정리
    // (spare2 체인이 새로 계산된 값으로 이어지므로 이전 수동 override는 모두 무효)
    const curIdx = excelDays.findIndex(d => d.dateLabel === currentDateKey);
    if (curIdx >= 0) {
      const subsequentKeys = new Set(excelDays.slice(curIdx + 1).map(d => d.dateLabel));
      setOverrideStartByDate(prev => {
        const hasAny = Object.keys(prev).some(k => subsequentKeys.has(k));
        if (!hasAny) return prev;
        const next = Object.fromEntries(Object.entries(prev).filter(([k]) => !subsequentKeys.has(k)));
        localStorage.setItem(OVERRIDE_KEY, JSON.stringify(next));
        return next;
      });
    }

    // ↑ 힌트 표시 범위: 배정+재계산 완료 다음 날 1일만
    const curIdx2 = excelDays.findIndex(d => d.dateLabel === currentDateKey);
    if (curIdx2 >= 0 && curIdx2 + 1 < excelDays.length) {
      const nextDk = excelDays[curIdx2 + 1].dateLabel;
      persistAssignedRange({ start: nextDk.slice(0, 5), end: nextDk.slice(0, 5) });
    }
    // ✓완료 표시용 배정 시각 기록
    recordAssignmentTimestamps([currentDateKey]);

    if (count > 0) {
      setRecalcMessage(`이 날짜 이후 ${count}일 스케줄이 업데이트되었습니다.`);
      setTimeout(() => setRecalcMessage(null), 4000);
    }
  }

  // force=false: N일 전체를 최신 dateStatuses로 재계산, 이후 날짜 연쇄 갱신 없음
  // force=true : N일 전체를 최신 dateStatuses로 재계산 후 이후 날짜도 recalculateFrom으로 연쇄 갱신
  // days: 생성할 날짜 수 (기본 7, 3일 생성 시 3)
  function generateWeek(force = false, days = 7) {
    if (!selectedDate) return;

    // ── 시작일: selectedDate 기준 Date 산술 (excelDays 없어도 동작, 월 경계 포함) ──
    const startDateLabel = selectedDate.dateLabel;
    const sm = startDateLabel.match(/^(\d{2})\.(\d{2})/);
    if (!sm) return;
    const startDate = new Date(viewYear, parseInt(sm[1], 10) - 1, parseInt(sm[2], 10));

    // ── force=false: days일 안에 기존 배정이 있으면 확인 모달 표시 후 중단 ──
    if (!force) {
      const hasExisting = Array.from({ length: days }, (_, di) => {
        const d = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate() + di);
        return makeDateKey(d);
      }).some(dl => !!assignmentData[dl]);
      if (hasExisting) {
        setWeekForceConfirm(days);
        return;
      }
    }

    const results: { day: string; result: DayResult; skipped?: boolean }[] = [];
    // 항상 각 날짜의 최신 dateStatuses 기준으로 재계산 (force 여부와 무관)
    const newAssignments: Record<string, DayResult> = {};

    // excelDays / viewDays 빠른 조회 맵 (① excelDays 우선, ② viewDays fallback)
    const excelMap = new Map(excelDays.map(d => [d.dateLabel, d]));
    const viewMap = new Map(viewDays.map(d => [d.dateLabel, d]));

    let lastDayLabel = startDateLabel;

    for (let di = 0; di < days; di++) {
      const date = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate() + di);
      const dateLabel = makeDateKey(date);
      lastDayLabel = dateLabel;

      // excelDays 우선, 없으면 viewDays fallback으로 dayIdx 얻기
      const weekDay = excelMap.get(dateLabel) ?? viewMap.get(dateLabel);
      const dayIdx = weekDay?.dayIdx ?? (date.getDay() + 6) % 7;

      // ── 시작 이름: 루프 내 직전 결과 우선 스페어[0] 우선 → getStartNameForDate → fallback ──
      // 직전 날짜를 이번 루프에서 방금 계산했다면 state 우회하여 직접 사용
      const prevLoopResult = results.length > 0 ? results[results.length - 1].result : null;
      const prevSpare2First = prevLoopResult ? (getPrioritySpares(prevLoopResult)[0] ?? null) : null;
      const dayStartName = prevSpare2First ?? getStartNameForDate(dateLabel);
      const dayNames = dayStartName ? rotateNames([...names], dayStartName) : [...names];

      // ── 배정 계산: 항상 해당 날짜의 최신 dateStatuses 기준으로 재계산 ──
      const savedDay = dateStatuses[dateLabel] ?? {};
      const statuses: Record<string, StatusType> = {};
      const dgMap = dateDaegeun[dateLabel] ?? {};
      dayNames.forEach((n) => {
        statuses[n] = resolveStatus(n, dateLabel, dayIdx, savedDay, dgMap);
      });

      const s1 = shift1Size, s2 = shift2Size, ss = singleSize;

      const result = mode === "2부제"
        ? assignDouble(dayNames, statuses, s1, s2, dateDaegeun[dateLabel] ?? {}, dateStatusOrders[dateLabel] ?? [])
        : assignSingle(dayNames, statuses, ss);

      results.push({ day: dateLabel, result, skipped: false });
      newAssignments[dateLabel] = result;
    }

    setWeekly(results);
    setDayResult(null);
    setPendingResult(null);

    // ── 배정 저장 (force: 7일 전체 / 非force: 새 날짜만) ──
    if (Object.keys(newAssignments).length > 0) {
      const mergedAssignment = { ...assignmentData, ...newAssignments };
      const newSpare2: Record<string, string[]> = {};
      Object.entries(newAssignments).forEach(([d, r]) => {
        if (d && r.spare2.length > 0) newSpare2[d] = r.spare2;
      });
      const mergedSpare2 = { ...savedSpare2, ...newSpare2 };

      if (force) {
        // 7일 이후 이미 배정된 날짜들도 새 spare2 체인으로 연쇄 재계산
        const { updatedAssignment, updatedSpare2, count } = recalculateFrom(lastDayLabel, mergedAssignment, mergedSpare2);
        setAssignmentData(updatedAssignment);
        setSavedSpare2(updatedSpare2);

        // 7일 범위 override 정리 (새 spare2 체인이 더 정확하므로)
        const forcedKeys = new Set(Object.keys(newAssignments));
        setOverrideStartByDate(prev => {
          const hasAny = Object.keys(prev).some(k => forcedKeys.has(k));
          if (!hasAny) return prev;
          const next = Object.fromEntries(Object.entries(prev).filter(([k]) => !forcedKeys.has(k)));
          localStorage.setItem(OVERRIDE_KEY, JSON.stringify(next));
          return next;
        });

        if (count > 0) {
          setRecalcMessage(`${days}일 재계산 완료 + 이후 ${count}일 연쇄 갱신되었습니다.`);
          setTimeout(() => setRecalcMessage(null), 4000);
        }
      } else {
        setAssignmentData(prev => ({ ...prev, ...newAssignments }));
        if (Object.keys(newSpare2).length > 0) {
          setSavedSpare2(prev => ({ ...prev, ...newSpare2 }));
        }
      }
    }

    // 힌트 표시 범위: 배정 마지막 날 다음날
    const afterDate = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate() + days);
    const afterDk = makeDateKey(afterDate).slice(0, 5);
    persistAssignedRange({ start: afterDk, end: afterDk });

    // ✓완료 표시용 배정 시각 기록 (새로 배정된 날짜만)
    const newDateLabels = Object.keys(newAssignments);
    if (newDateLabels.length > 0) recordAssignmentTimestamps(newDateLabels);
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
  const displayPrioritySpares = displayResult ? getPrioritySpares(displayResult) : [];
  const showFloatingBar = !!(displayPrioritySpares[0] && weekly.length === 0);
  return (
    <div style={{ ...S.page, paddingBottom: showFloatingBar ? "80px" : undefined }}>
      {/* 헤더 */}
      <div style={S.header}>
        <button onClick={() => setLocation(`${BASE}/`)} style={S.backBtn}>←</button>
        <img src={`${BASE}/char_dino.png`} alt="" style={{ width: 30, height: 30, objectFit: "contain" }} />
        <span style={S.headerTitle}>캐디 근무표</span>
        {names.length > 0 && (
          <button
            onClick={() => {
              setNames([]); setRosterLoaded(false); setDayResult(null); setPendingResult(null); setWeekly([]);
              window.scrollTo({ top: 0, behavior: "smooth" });
            }}
            style={S.smallBtn}
            className="text-[14px]">↩ 초기화</button>
        )}
        {names.length === 0 && customRoster.length > 0 && (
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
      {/* ─── 입력 단계 (항상 표시) ─── */}
        <div style={S.card}>
          {/* ── 날짜 선택 + 휴무 교체 버튼 (같은 행) ── */}
          <div style={{ display: "flex", alignItems: "center", gap: "8px", marginBottom: "4px" }}>
            <label style={{ ...S.label, margin: 0, flex: 1 }}>
              {selectedDate ? selectedDate.dateLabel : "날짜 선택"}
              {xlLoading && <span style={{ color: "#aaa", fontWeight: 400, marginLeft: "6px" }}>불러오는 중…</span>}
              {xlError && <span style={{ color: "#e53935", fontWeight: 400, marginLeft: "6px" }}>{xlError}</span>}
            </label>
            {isAdmin && (
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
            )}
          </div>

          {/* 휴무 파일 정보 */}
          <div style={{ fontSize: "0.72rem", color: "#555", marginBottom: "8px" }}>
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

          {/* ── 캐릭터 + 월 네비게이션 (엑셀 유무와 무관하게 항상 표시) ── */}
          <>
            <div style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "8px" }}>
              <img
                src={`${BASE}/char_cloud.png`}
                alt="cloud"
                style={{ width: 44, height: 44, objectFit: "contain", animation: "floatBob 3s ease-in-out infinite" }}
              />
              <div style={{ flex: 1 }}>
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
              {displayDays.map((d, idx) => {
                const isSelected = selectedDate?.dateLabel === d.dateLabel;
                const isWeekend = d.dayIdx === 5 || d.dayIdx === 6;
                const hasTeams = d.예약팀수 > 0;
                const savedStatuses = dateStatuses[d.dateLabel] ?? {};
                const hasManual = Object.keys(savedStatuses).length > 0;
                const hasExcel = d.가용인원 > 0 || hasTeams;

                // 달력 첫번호 힌트: 전날 우선 스페어[0] 직접 계산 (선택 여부 무관, 항상 렌더)
                // pendingResult가 있으면 최우선 (배정 직후 즉시 반영), 없으면 전날 spare2 실시간 탐색
                const calHint: string | null = (() => {
                  if (d.dateLabel === currentDateKey && pendingResult) {
                    return getPrioritySpares(pendingResult)[0] ?? null;
                  }
                  const m = d.dateLabel.match(/^(\d{2})\.(\d{2})/);
                  if (!m) return null;
                  const prev = new Date(viewYear, parseInt(m[1], 10) - 1, parseInt(m[2], 10) - 1);
                  const pDL = makeDateKey(prev);
                  return getPrioritySpares(assignmentData[pDL])[0] ?? savedSpare2[pDL]?.[0] ?? null;
                })();
                // 배정 표시: 배정을 실행한 시각 기준 7일 이내만 ✓완료 뱃지 & 초록 배경
                // (달력 날짜 기준 X → 배정 누른 시각 기준 O)
                const hasAssigned = (() => {
                  if (!assignmentData[d.dateLabel]) return false;
                  const ts = assignmentTimestamps[d.dateLabel];
                  if (!ts) return true; // 타임스탬프 없으면 (구버전 데이터) 항상 표시
                  return Date.now() - ts <= 7 * 24 * 60 * 60 * 1000;
                })();
                const hintVisible = !!calHint;

                return (
                  <button
                    key={d.dateLabel}
                    onClick={() => selectExcelDate(d)}
                    style={{
                      ...S.dateBtn,
                      background: isSelected
                        ? "linear-gradient(135deg, #1a1a2e 0%, #4e89ae 100%)"
                        : hasAssigned ? "#f0fdf4"
                        : hasManual ? "#eff6ff" : "#f8fafc",
                      color: isSelected ? "#fff" : isWeekend ? "#c62828" : "#1a1a2e",
                      border: isSelected
                        ? "2px solid #4e89ae"
                        : hasAssigned ? "2px solid #86efac"
                        : hasManual ? "2px solid #93c5fd"
                        : hasTeams ? "2px solid #60a5fa" : "1.5px solid #e5e7eb",
                      animation: isSelected ? "glowPulse 2s ease-in-out infinite" : "none",
                      transform: isSelected ? "scale(1.05)" : "scale(1)",
                      opacity: !hasExcel && !hasManual && !isSelected ? 0.75 : 1,
                      minHeight: hintVisible ? (isSelected ? "68px" : "64px") : "52px",
                    }}
                  >
                    <span style={{ fontSize: "0.72rem", fontWeight: 700 }}>{d.dateLabel.split(" ")[0]}</span>
                    <span style={{ fontSize: "0.65rem", opacity: 0.7 }}>{d.dayName}</span>
                    {hasAssigned && !isSelected && (
                      <span style={{
                        fontSize: "0.5rem", fontWeight: 800, lineHeight: 1,
                        color: "#15803d", background: "#dcfce7",
                        borderRadius: 4, padding: "1px 4px",
                      }}>✓완료</span>
                    )}
                    {d.가용인원 > 0 && (
                      <span style={{
                        fontSize: "0.6rem", fontWeight: 700,
                        color: isSelected ? "#fff" : "#2e7d32",
                        lineHeight: 1,
                      }}>
                        {customRoster.length}/{d.가용인원}
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
                    {/* 전날 spare2[0] → 이 날의 첫번호 힌트 (선택 여부 무관, 항상 표시 / 선택 시 강조) */}
                    {calHint && (
                      isSelected ? (
                        <span style={{
                          fontSize: "0.58rem", fontWeight: 900, lineHeight: 1,
                          color: "#fff",
                          background: "rgba(255,255,255,0.25)",
                          borderRadius: 5, padding: "2px 6px",
                          marginTop: 2,
                          maxWidth: "100%", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
                          border: "1px solid rgba(255,255,255,0.45)",
                          letterSpacing: "0.02em",
                        }}>
                          📌 {calHint}
                        </span>
                      ) : (
                        <span style={{
                          fontSize: "0.52rem", fontWeight: 800, lineHeight: 1,
                          color: "#b91c1c",
                          background: "#fef2f2",
                          borderRadius: 4, padding: "1px 4px",
                          marginTop: 1,
                          maxWidth: "100%", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
                        }}>
                          ↑{calHint}
                        </span>
                      )
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
              </div>
              {/* 총인원 · 가용인원(계산) · 투인원 · 대근 · 병가 — 5열 그리드 */}
              <div style={{ overflowX: "auto" }}>
              <div style={{ ...S.excelStatRow, alignItems: "stretch" }}>
                <StatBadge label="총 인원" value={customRoster.length} color="#1565c0" />
                {(() => {
                  // 통일 계산: 가용인원 = 총인원 - (휴무+당번+병가) — calcAvailable 경유
                  const availableCount = calcAvailable(selectedDate.dateLabel, dayOfWeek);
                  return <StatBadge label="가용인원" value={availableCount} color="#7c3aed" />;
                })()}
                {(() => {
                  // 투인원: calcAvailable이 이미 당번·병가 제외 → 직접 사용
                  const availableCount = calcAvailable(selectedDate.dateLabel, dayOfWeek);
                  const totalTeams = mode === "2부제" ? totalSize : singleSize;
                  const tuInwon = totalTeams > 0 ? totalTeams - availableCount : null;
                  return (
                    <StatBadge
                      label="투 인원"
                      value={tuInwon !== null ? tuInwon : "–"}
                      color={tuInwon !== null && tuInwon > 0 ? "#e65100" : "#9e9e9e"}
                    />
                  );
                })()}
                {(() => {
                  const daegeunBaseNames = names.length > 0 ? names : sortedCustomRoster.map(p => p.name);
                  const daegeunCandidates = daegeunBaseNames.filter(n => {
                    const p = getRosterPerson(n);
                    return p != null && isAutoOff(p.group, dayOfWeek) && !(n in manualStatuses);
                  });
                  const activeDaegeun = daegeunCandidates.filter(n => currentDaegeun[n]);
                  return (
                    <button
                      onClick={() => { setBatchDaegeunOpen(true); setBatchDaegeunSearch(""); }}
                      style={{
                        display: "flex", flexDirection: "column",
                        alignItems: "center", justifyContent: "center",
                        padding: "6px 4px",
                        borderRadius: "8px",
                        border: `1px solid ${activeDaegeun.length > 0 ? "#f59e0b" : "#f59e0b33"}`,
                        background: activeDaegeun.length > 0 ? "#fef3c7" : "#f59e0b15",
                        cursor: "pointer",
                        width: "100%",
                        boxSizing: "border-box",
                        fontFamily: "inherit",
                        lineHeight: 1,
                        outline: "none",
                        WebkitTapHighlightColor: "transparent",
                      }}>
                      <span style={{ fontSize: "0.62rem", fontWeight: 600, color: "#92400e", whiteSpace: "nowrap" }}>대근</span>
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
                        padding: "6px 4px",
                        borderRadius: "8px",
                        border: `1px solid ${active ? "#c62828" : "#c6282833"}`,
                        background: active ? "#ffebee" : "#c6282815",
                        cursor: "pointer",
                        width: "100%",
                        boxSizing: "border-box",
                        fontFamily: "inherit",
                        lineHeight: 1,
                        outline: "none",
                        WebkitTapHighlightColor: "transparent",
                      }}>
                      <span style={{ fontSize: "0.62rem", fontWeight: 600, color: "#b71c1c", whiteSpace: "nowrap" }}>병가</span>
                      <span style={{ fontSize: "1rem", fontWeight: 700, color: active ? "#c62828" : "#d1d5db" }}>
                        {active ? sickCnt : "–"}
                      </span>
                    </button>
                  );
                })()}
              </div>
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
                      {currentVip.count}팀
                    </span>
                  )}
                  {currentVipMembers.length > 0 && (
                    <span style={{ fontSize: "0.75rem", color: "#9c27b0", fontWeight: 600 }}>
                      · {currentVipMembers.map(m => m.name).join(", ")}
                    </span>
                  )}
                  <button
                    onClick={() => { setVipSearch(""); setVipModalOpen(true); }}
                    style={{
                      marginLeft: "auto", padding: "4px 12px", borderRadius: "10px",
                      border: "1.5px solid #ce93d8",
                      background: "#fce4ec", color: "#7b1fa2",
                      fontWeight: 700, fontSize: "0.78rem", cursor: "pointer",
                      flexShrink: 0,
                    }}>
                    {currentVip.count > 0 || currentVipMembers.length > 0 ? "수정" : "+ 추가"}
                  </button>
                </div>

                {/* VIP 인원 칩 (항상 표시) */}
                {currentVipMembers.length > 0 && (
                  <div style={{ display: "flex", flexWrap: "wrap", gap: "5px", marginBottom: "8px" }}>
                    {currentVipMembers.map(({ name, type }) => {
                      const sc = STATUS_COLOR[type] ?? { bg: "#f3e5f5", color: "#7b1fa2" };
                      const typeLabel = type === "VIP1부" ? "1부" : type === "VIP2부" ? "2부" : "투근무";
                      return (
                        <span
                          key={name}
                          style={{
                            display: "inline-flex", alignItems: "center", gap: "4px",
                            padding: "3px 10px", borderRadius: "20px",
                            background: sc.bg, color: sc.color, border: `1px solid ${sc.color}44`,
                            fontSize: "0.76rem", fontWeight: 700,
                          }}
                        >
                          {name}
                          <span style={{ opacity: 0.7, fontSize: "0.68rem" }}>({typeLabel})</span>
                          <button
                            onClick={() => clearStatus(name)}
                            style={{ background: "none", border: "none", color: sc.color, cursor: "pointer", fontSize: "0.85rem", lineHeight: 1, padding: 0 }}
                          >×</button>
                        </span>
                      );
                    })}
                  </div>
                )}

              </div>

            </div>
          )}

          {/* 팀수 입력 / 저장 */}
          {teamsLocked ? (
            /* ── 저장된 팀수 요약 카드 ── */
            (<div style={{
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
            </div>)
          ) : (
            /* ── 팀수 입력 폼 (통합: 총 팀수 필수, 1부 팀수 선택 → 자동 판별) ── */
            (<div style={{ marginBottom: "4px", marginTop: "14px" }}>
              <div style={{ display: "flex", gap: "10px", marginBottom: "8px" }}>
                <div style={{ flex: 1 }}>
                  <label style={S.label}>
                    총 팀수 <span style={{ color: "#e53935" }}>*</span>
                  </label>
                  <input type="number" value={totalSize} min={1}
                    onChange={(e) => setTotalSize(Number(e.target.value))} style={S.numInput} />
                </div>
                <div style={{ flex: 1 }}>
                  <label style={S.label}>
                    1부 팀수 <span style={{ color: "#9ca3af", fontSize: "0.7rem" }}>(선택)</span>
                  </label>
                  <input
                    type="number"
                    value={shift1Input}
                    placeholder="비우면 단부제"
                    min={1}
                    onChange={(e) => setShift1Input(e.target.value)}
                    style={S.numInput}
                  />
                </div>
              </div>
              {/* 1부 >= 총 팀수 오류 */}
              {mode === "2부제" && shift1Size >= totalSize && (
                <div style={{ color: "#c62828", fontSize: "0.78rem", marginBottom: "6px", fontWeight: 600 }}>
                  ⚠️ 1부 팀수는 총 팀수보다 작아야 합니다
                </div>
              )}
              {/* 계산식 요약 */}
              {mode === "2부제" ? (
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
              ) : (
                <div style={S.calcBox}>
                  <span style={{ fontWeight: 700, color: "#374151" }}>단부제</span>
                  <span style={{ color: "#aaa" }}>·</span>
                  <span style={{ fontWeight: 700 }}>총 {totalSize}팀</span>
                </div>
              )}
              <button
                onClick={saveTeamSettings}
                disabled={mode === "2부제" && shift1Size >= totalSize}
                style={{
                  width: "100%", marginTop: "10px", padding: "11px",
                  borderRadius: "12px", border: "none",
                  background: (mode === "2부제" && shift1Size >= totalSize) ? "#9ca3af" : "#2e7d32",
                  color: "#fff", fontWeight: 800, fontSize: "0.9rem",
                  cursor: (mode === "2부제" && shift1Size >= totalSize) ? "not-allowed" : "pointer",
                }}>
                💾 저장하기
              </button>
            </div>)
          )}

          {/* 순번표 불러오기 + 편집 */}
          <div style={{ display: "flex", gap: "8px", marginBottom: "8px" }}>
            <button onClick={loadRoster} style={{ ...S.primaryBtn, background: "#1565c0", flex: 1, marginBottom: 0, touchAction: "manipulation" }}>
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
      {/* ── 첫번호: 이 순번대로 가시겠습니까? ── */}
      {queueModal === "ask" && (() => {
        // 가장 최근 저장된 우선 스페어 첫번째 찾기
        const spare2First: string | null = (() => {
          // 선택된 날짜 전날 스페어 우선
          if (todayFirstHint) return todayFirstHint;
          // 없으면 savedSpare2에서 가장 최근 날짜의 첫번째
          const entries = Object.entries(savedSpare2);
          if (!entries.length) return null;
          entries.sort((a, b) => a[0].localeCompare(b[0]));
          const [dateLabel, spare2] = entries[entries.length - 1];
          return getPrioritySpares(assignmentData[dateLabel])[0] ?? spare2[0] ?? null;
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

              {/* 안하겠습니다 — spare2 체인으로 적용 + stale override 제거 */}
              <button
                onClick={() => {
                  if (spare2First) {
                    applyRoster(spare2First, { saveOverride: false });
                    // spare2 체인이 확정되면 해당 날짜의 stale override 제거
                    // → 달력 힌트도 spare2 체인값으로 일치
                    if (currentDateKey && overrideStartByDate[currentDateKey]) {
                      setOverrideStartByDate(prev => {
                        const next = { ...prev };
                        delete next[currentDateKey];
                        localStorage.setItem(OVERRIDE_KEY, JSON.stringify(next));
                        return next;
                      });
                    }
                  } else {
                    // 저장된 스페어 없으면 queueStartName 유지 (override 저장 X)
                    applyRoster(queueStartName, { saveOverride: false });
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
                  onClick={() => {
                    applyRoster(null);
                    saveQueueStart(null);
                    // 현재 날짜 override 삭제
                    if (currentDateKey) {
                      setOverrideStartByDate(prev => {
                        const next = { ...prev };
                        delete next[currentDateKey];
                        localStorage.setItem(OVERRIDE_KEY, JSON.stringify(next));
                        return next;
                      });
                    }
                    setQueueModal(null);
                  }}
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
          // 키보드가 올라온 만큼 하단 패딩 추가 → 시트가 키보드 위에 위치
          paddingBottom: Math.max(0, window.innerHeight - vvHeight),
        }}>
          <div style={{
            background: "#fff", borderRadius: "18px 18px 0 0",
            // 키보드 가시 영역 기준 최대 높이 (vvHeight 88%)
            maxHeight: Math.min(vvHeight * 0.88, vvHeight - 20),
            display: "flex", flexDirection: "column",
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
            {queueStartName && sortedCustomRoster.some(p => normalize(p.name) === normalize(queueStartName)) && (
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
                    const person = sortedCustomRoster.find(p => normalize(p.name) === normalize(queueStartName));
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
                onFocus={() => {
                  // 키보드 올라온 후 목록이 보이도록 약간의 지연 후 스크롤
                  setTimeout(() => {
                    queueListRef.current?.scrollIntoView({ behavior: "smooth", block: "nearest" });
                  }, 350);
                }}
                style={{
                  width: "100%", padding: "9px 12px", borderRadius: 10,
                  border: "1px solid #ddd", fontSize: 14, boxSizing: "border-box",
                }}
              />
            </div>
            {/* 목록 */}
            <div ref={queueListRef} style={{ overflowY: "auto", flex: 1, paddingBottom: 16 }}>
              {sortedCustomRoster
                .filter(p => !queuePickSearch || normalize(p.name).includes(normalize(queuePickSearch)))
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
                      background: normalize(p.name) === normalize(queueStartName ?? "") ? "#e3f2fd" : "transparent",
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
                    {normalize(p.name) === normalize(queueStartName ?? "") && (
                      <span style={{ color: "#1565c0", fontSize: 13, fontWeight: 700 }}>✓ 현재</span>
                    )}
                  </button>
                ))}
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
              <button onClick={forceUploadToServer} style={{
                padding: "5px 10px", borderRadius: "8px",
                border: "none", background: "#b71c1c", color: "#fff",
                fontSize: "0.75rem", fontWeight: 700, cursor: "pointer",
              }}>☁️ 서버 덮어쓰기</button>

              <button onClick={() => { setRosterEditorOpen(false); setRosterForm(null); }}
                style={{
                  background: "rgba(255,255,255,0.15)", border: "none",
                  borderRadius: "50%", width: "30px", height: "30px",
                  cursor: "pointer", fontSize: "1rem", color: "#fff", fontWeight: 700,
                }}>✕</button>
            </div>

            {rosterForm ? (
              /* ── 추가/수정 폼 ── */
              (<div style={{ padding: "20px 18px", overflowY: "auto" }}>
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
              </div>)
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
                      !q || normalize(p.name).includes(normalize(q)) || p.조.toString().includes(q) || p.group.includes(q)
                    );
                    const JO_COLORS: Record<number, { bg: string; color: string }> = {
                      1: { bg: "#fce4ec", color: "#c62828" },
                      2: { bg: "#e8f5e9", color: "#2e7d32" },
                      3: { bg: "#e3f2fd", color: "#1565c0" },
                      4: { bg: "#fff8e1", color: "#f57f17" },
                    };
                    let lastJo: number | null = null;
                    const items: React.ReactNode[] = [];

                    filtered.forEach((p) => {
                      const sameJoSorted = filtered.filter(x => x.조 === p.조).sort((a, b) => a.no - b.no);
                      const posInJo = sameJoSorted.findIndex(x => x.name === p.name);
                      const isFirst = posInJo === 0;
                      const isLast = posInJo === sameJoSorted.length - 1;

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
                          display: "flex", alignItems: "center", gap: "6px",
                          padding: "7px 14px", borderBottom: "1px solid #f9f9f9",
                        }}>
                          {/* 순서 번호 */}
                          <span style={{
                            fontSize: "0.65rem", color: "#aaa", fontWeight: 600,
                            minWidth: "16px", textAlign: "right", flexShrink: 0,
                          }}>{posInJo + 1}</span>

                          {/* ▲ ▼ 이동 버튼 */}
                          <div style={{ display: "flex", flexDirection: "column", gap: "1px", flexShrink: 0 }}>
                            <button
                              onClick={() => movePersonInJo(p, "up")}
                              disabled={isFirst || !!q}
                              style={{
                                width: "24px", height: "22px", border: "none",
                                borderRadius: "5px 5px 2px 2px",
                                background: isFirst || q ? "#f0f0f0" : "#e3f2fd",
                                color: isFirst || q ? "#ccc" : "#1565c0",
                                fontSize: "0.65rem", cursor: isFirst || q ? "default" : "pointer",
                                fontWeight: 800, lineHeight: 1, padding: 0,
                              }}>▲</button>
                            <button
                              onClick={() => movePersonInJo(p, "down")}
                              disabled={isLast || !!q}
                              style={{
                                width: "24px", height: "22px", border: "none",
                                borderRadius: "2px 2px 5px 5px",
                                background: isLast || q ? "#f0f0f0" : "#e3f2fd",
                                color: isLast || q ? "#ccc" : "#1565c0",
                                fontSize: "0.65rem", cursor: isLast || q ? "default" : "pointer",
                                fontWeight: 800, lineHeight: 1, padding: 0,
                              }}>▼</button>
                          </div>

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
      {/* ── 통합 선택 모달 (조출/후출/찾근/VIP/기타 공통) ── */}
      {modalStatus && (() => {
        const isVip = modalStatus === "VIP";
        const statusSc = isVip
          ? { bg: "#f3e5f5", color: "#7b1fa2" }
          : (STATUS_COLOR[modalStatus as string] ?? { bg: "#f5f5f5", color: "#333" });

        // 현재 선택 목록
        const _filtered = isVip
          ? currentVipMembers.map(m => m.name)
          : names.filter(n => effectiveStatus(n) === modalStatus as StatusType);
        // 클릭 순서(dateStatusOrders) 우선 → 나머지는 기존 순서 유지
        // 모든 상태(휴무/병가/당번/대기/찾근/조출/후출) 동일하게 적용
        const selectedNames = !isVip
          ? (() => {
              const order = dateStatusOrders[currentDateKey] ?? [];
              const orderSet = new Set(order);
              return [
                ...order.filter(n => _filtered.includes(n)),
                ..._filtered.filter(n => !orderSet.has(n)),
              ];
            })()
          : _filtered;

        const listNames = names.length > 0 ? names : sortedCustomRoster.map(p => p.name);
        const query = modalSearch.trim().toLowerCase();
        const filteredNames = listNames.filter(n =>
          !query || n.toLowerCase().includes(query) ||
          (getRosterPerson(n)?.조?.toString() ?? "").includes(query)
        );

        const JO_COLORS: Record<number, { bg: string; color: string }> = {
          1: { bg: "#fce4ec", color: "#c62828" },
          2: { bg: "#e8f5e9", color: "#2e7d32" },
          3: { bg: "#e3f2fd", color: "#1565c0" },
          4: { bg: "#fff8e1", color: "#f57f17" },
        };

        // ── 리스트 렌더 헬퍼 (검색결과 / 전체리스트 공유) ──
        const renderItems = (nameList: string[], showJoDivider: boolean) => {
          if (nameList.length === 0) {
            return <div style={{ textAlign: "center", color: "#bbb", padding: "20px" }}>검색 결과 없음</div>;
          }
          let lastJo: number | null = null;
          const items: React.ReactNode[] = [];
          nameList.forEach((name) => {
            const person = getRosterPerson(name);
            const joNum = person?.조;
            const effS = effectiveStatus(name);

            if (showJoDivider && !query && rosterLoaded && joNum !== undefined && joNum !== lastJo) {
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

            if (isVip) {
              const isVipSelected = VIP_STATUSES.has(effS);
              const vipType = isVipSelected ? effS as "VIP1부" | "VIP2부" | "VIP투근무" : null;
              const vipSc = vipType ? STATUS_COLOR[vipType] : null;
              const isExpanded = vipSubPicking === name;
              const nonVipStatus = !isVipSelected && effS !== null ? effS : null;
              items.push(
                <div key={name} style={{
                  borderBottom: "1px solid #fce4ec",
                  background: isVipSelected ? (vipSc?.bg ?? "#f3e5f5") + "44" : "transparent",
                  borderLeft: isVipSelected ? `3px solid ${vipSc?.color ?? "#7b1fa2"}` : "3px solid transparent",
                }}>
                  <div
                    onClick={() => setVipSubPicking(isExpanded ? null : name)}
                    style={{
                      display: "flex", padding: "10px 14px",
                      alignItems: "center", gap: "10px", cursor: "pointer",
                    }}
                  >
                    {person && (
                      <span style={{
                        background: isVipSelected ? (vipSc?.color ?? "#7b1fa2") : "#e8eaf6",
                        borderRadius: 6, padding: "2px 7px",
                        fontSize: "0.7rem", color: isVipSelected ? "#fff" : "#5c6bc0",
                        fontWeight: 700, minWidth: 40, textAlign: "center",
                      }}>
                        {person.조}조 {person.no}번
                      </span>
                    )}
                    <div style={{ flex: 1 }}>
                      <span style={{ fontWeight: 600, fontSize: "0.92rem" }}>{name}</span>
                      {(() => { const grp = person?.group ?? getGroup(name); const gs = GROUP_STYLE[grp]; return (
                        <span style={{ marginLeft: "6px", fontSize: "0.65rem", color: gs.color, fontWeight: 600 }}>
                          {gs.label}
                        </span>
                      ); })()}
                    </div>
                    {vipType && (
                      <span style={{
                        fontSize: "0.72rem", fontWeight: 700,
                        color: vipSc?.color, background: vipSc?.bg,
                        borderRadius: 6, padding: "2px 7px",
                        border: `1px solid ${vipSc?.color}44`,
                      }}>
                        {vipType === "VIP1부" ? "1부" : vipType === "VIP2부" ? "2부" : "투근무"}
                      </span>
                    )}
                    {nonVipStatus && !isExpanded && (
                      <span style={{
                        fontSize: "0.7rem", fontWeight: 700,
                        color: (STATUS_COLOR[nonVipStatus] ?? { color: "#888" }).color,
                        background: (STATUS_COLOR[nonVipStatus] ?? { bg: "#eee" }).bg,
                        borderRadius: 6, padding: "2px 7px",
                      }}>{nonVipStatus}</span>
                    )}
                    <span style={{ color: isExpanded ? "#7b1fa2" : "#ccc", fontSize: "0.85rem", flexShrink: 0 }}>
                      {isExpanded ? "▲" : "▼"}
                    </span>
                  </div>
                  {isExpanded && (
                    <div style={{ padding: "6px 14px 12px", display: "flex", gap: "8px" }}>
                      {(["VIP1부", "VIP2부", "VIP투근무"] as const).map(st => {
                        const btnSc = STATUS_COLOR[st];
                        const active = effS === st;
                        const lbl = st === "VIP1부" ? "1부" : st === "VIP2부" ? "2부" : "투근무";
                        return (
                          <button
                            key={st}
                            onClick={() => {
                              if (active) {
                                clearStatus(name);
                              } else {
                                setManualStatuses(prev => ({ ...prev, [name]: st }));
                              }
                              setVipSubPicking(null);
                            }}
                            style={{
                              flex: 1, padding: "9px 4px", borderRadius: 10,
                              border: `2px solid ${active ? btnSc.color : btnSc.color + "44"}`,
                              background: active ? btnSc.color : btnSc.bg,
                              color: active ? "#fff" : btnSc.color,
                              fontWeight: 800, fontSize: "0.85rem", cursor: "pointer",
                            }}
                          >
                            {lbl}
                          </button>
                        );
                      })}
                      {isVipSelected && (
                        <button
                          onClick={() => { clearStatus(name); setVipSubPicking(null); }}
                          style={{
                            padding: "9px 10px", borderRadius: 10,
                            border: "1px solid #ef9a9a", background: "#ffebee",
                            color: "#c62828", fontWeight: 700, fontSize: "0.8rem", cursor: "pointer",
                          }}
                        >해제</button>
                      )}
                    </div>
                  )}
                </div>
              );
            } else {
              const st = modalStatus as StatusType;
              const isSelected = effS === st;
              const isDifferent = effS !== null && effS !== st && !VIP_STATUSES.has(effS);
              const isVipDiff = VIP_STATUSES.has(effS);
              const isDisabled = (() => {
                if (isSelected) return false;
                if (st === "조출") return !cho가능 || cho현재수 >= 6;
                if (st === "후출") return hu현재수 >= 6;
                return false;
              })();
              items.push(
                <div key={name}
                  onClick={() => !isDisabled && toggleStatus(name, st)}
                  style={{
                    display: "flex", alignItems: "center", gap: "10px",
                    padding: "10px 14px",
                    background: isSelected ? statusSc.bg + "33" : "transparent",
                    borderLeft: isSelected ? `3px solid ${statusSc.color}` : "3px solid transparent",
                    cursor: isDisabled ? "not-allowed" : "pointer",
                    opacity: isDisabled ? 0.38 : 1,
                  }}
                >
                  <div style={{
                    width: "22px", height: "22px", borderRadius: "6px", flexShrink: 0,
                    border: `2px solid ${isSelected ? statusSc.color : "#ddd"}`,
                    background: isSelected ? statusSc.bg : "#fff",
                    display: "flex", alignItems: "center", justifyContent: "center",
                  }}>
                    {isSelected && <span style={{ fontSize: "0.9rem", color: statusSc.color, fontWeight: 900 }}>✓</span>}
                  </div>
                  <div style={{ flex: 1 }}>
                    <span style={{ fontWeight: 600, fontSize: "0.92rem" }}>{name}</span>
                    {(() => { const grp = person?.group ?? getGroup(name); const gs = GROUP_STYLE[grp]; return (
                      <span style={{ marginLeft: "6px", fontSize: "0.65rem", color: gs.color, fontWeight: 600 }}>
                        {person ? `${person.조}조 · ` : ""}{gs.label}
                      </span>
                    ); })()}
                  </div>
                  {(isDifferent || isVipDiff) && (
                    <span style={{
                      padding: "2px 8px", borderRadius: "12px", fontSize: "0.7rem", fontWeight: 700,
                      background: (STATUS_COLOR[effS as string] ?? { bg: "#eee" }).bg,
                      color: (STATUS_COLOR[effS as string] ?? { color: "#555" }).color,
                    }}>
                      {isVipDiff
                        ? (effS === "VIP1부" ? "VIP 1부" : effS === "VIP2부" ? "VIP 2부" : "VIP 투근무")
                        : effS}
                    </span>
                  )}
                  {isDisabled && !isSelected && (
                    <span style={{ fontSize: "0.65rem", color: "#bbb" }}>불가</span>
                  )}
                </div>
              );
            }
          });
          return items;
        };

        // ── 섹션 헤더 스타일 헬퍼 ──
        const SectionLabel = ({ num, title, count }: { num: string; title: string; count?: number }) => (
          <div style={{ display: "flex", alignItems: "center", gap: "6px", padding: "10px 14px 4px" }}>
            <span style={{
              fontSize: "0.65rem", fontWeight: 700, color: "#fff",
              background: "#c0c6d0", borderRadius: "4px", padding: "1px 5px",
              lineHeight: "1.4",
            }}>{num}</span>
            <span style={{ fontSize: "0.8rem", fontWeight: 700, color: "#444" }}>{title}</span>
            {count !== undefined && (
              <span style={{ fontSize: "0.75rem", color: "#9aa3b5", fontWeight: 600 }}>{count}명</span>
            )}
          </div>
        );

        return (
          <div style={{
            position: "fixed", inset: 0, zIndex: 200,
            background: "rgba(0,0,0,0.45)",
            display: "flex", flexDirection: "column", justifyContent: "flex-end",
          }}
            onClick={(e) => { if (e.target === e.currentTarget) { setModalStatus(null); setVipSubPicking(null); } }}
          >
            <div style={{
              background: "#fff", borderRadius: "18px 18px 0 0",
              maxHeight: "90vh", display: "flex", flexDirection: "column",
              overflow: "hidden",
            }}>
              {/* 헤더 */}
              <div style={{
                display: "flex", alignItems: "center", padding: "14px 16px 12px",
                borderBottom: "1px solid #f0f0f0",
                background: statusSc.bg, flexShrink: 0,
              }}>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 800, fontSize: "1rem", color: statusSc.color }}>
                    {isVip ? "👑 VIP 배정" : `${modalStatus} 배정`}
                  </div>
                  <div style={{ fontSize: "0.75rem", color: statusSc.color + "aa", marginTop: "1px" }}>
                    {isVip
                      ? "이름 탭 → 1부 / 2부 / 투근무 선택"
                      : "이름을 탭하면 선택/해제됩니다"}
                  </div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                  {selectedNames.length > 0 && (
                    <span style={{
                      padding: "3px 10px", borderRadius: "20px",
                      background: statusSc.color + "22", color: statusSc.color,
                      fontSize: "0.78rem", fontWeight: 800,
                    }}>{selectedNames.length}명</span>
                  )}
                  <button onClick={() => { setModalStatus(null); setVipSubPicking(null); }}
                    style={{
                      background: "rgba(255,255,255,0.4)", border: "none",
                      borderRadius: "50%", width: "30px", height: "30px",
                      cursor: "pointer", fontSize: "1rem", fontWeight: 700,
                      color: statusSc.color,
                    }}>✕</button>
                </div>
              </div>

              {/* ── 스크롤 영역 ── */}
              <div style={{ overflowY: "auto", flex: 1 }}>

                {/* 1. 선택된 인원 chip 영역 */}
                <SectionLabel num="1" title={`총 ${selectedNames.length}명`} />
                <div style={{
                  padding: "4px 14px 10px",
                  display: "flex", flexWrap: "wrap", gap: "6px",
                  minHeight: "36px", maxHeight: "200px", overflowY: "auto",
                  borderBottom: "1px solid #f0f0f0",
                }}>
                  {selectedNames.length === 0 ? (
                    <span style={{ fontSize: "0.8rem", color: "#ccc", alignSelf: "center" }}>선택된 인원 없음</span>
                  ) : isVip ? (
                    currentVipMembers.map(({ name, type }) => {
                      const sc = STATUS_COLOR[type] ?? { bg: "#f3e5f5", color: "#7b1fa2" };
                      const lbl = type === "VIP1부" ? "1부" : type === "VIP2부" ? "2부" : "투근무";
                      return (
                        <span key={name} style={{
                          display: "inline-flex", alignItems: "center", gap: "4px",
                          padding: "4px 10px", borderRadius: "20px",
                          background: sc.bg, color: sc.color, border: `1px solid ${sc.color}44`,
                          fontSize: "0.76rem", fontWeight: 700,
                        }}>
                          {name}<span style={{ opacity: 0.7, fontSize: "0.68rem" }}>({lbl})</span>
                          <button onClick={() => clearStatus(name)}
                            style={{ background: "none", border: "none", color: sc.color, cursor: "pointer", fontSize: "0.85rem", padding: 0, lineHeight: 1 }}>×</button>
                        </span>
                      );
                    })
                  ) : (
                    <>
                      <div style={{ width: "100%", fontSize: "0.72rem", color: "#aaa", marginBottom: "4px", paddingLeft: "2px" }}>
                        ↕ 드래그해서 순서 변경 (그룹 내)
                      </div>
                      {(["하우스", "주말", "주중"] as const).map(grp => {
                        const gs = GROUP_STYLE[grp];
                        // selectedNames 순서 유지하면서 이 그룹 이름만 추출 (인덱스 보존)
                        const grpEntries = selectedNames
                          .map((n, idx) => ({ n, idx }))
                          .filter(({ n }) => (getGroup(n) as string) === grp);
                        if (grpEntries.length === 0) return null;
                        return (
                          <div key={grp} style={{ width: "100%" }}>
                            <div style={{ display: "flex", alignItems: "center", gap: "5px", marginBottom: "5px" }}>
                              <span style={{ color: gs.color, fontSize: "0.85rem", lineHeight: 1 }}>●</span>
                              <span style={{ fontSize: "0.75rem", fontWeight: 800, color: gs.color }}>{grp} {grpEntries.length}명</span>
                            </div>
                            <div style={{ display: "flex", flexWrap: "wrap", gap: "6px", marginBottom: "6px" }}>
                              {grpEntries.map(({ n, idx }) => {
                                const isDragSrc = chipDragRef.current.fromIdx === idx;
                                const isDragOver = chipDragOver === idx;
                                return (
                                  <div
                                    key={n}
                                    data-chip-index={String(idx)}
                                    draggable
                                    onDragStart={() => {
                                      chipDragRef.current.fromIdx = idx;
                                      chipDragRef.current.didDrag = false;
                                    }}
                                    onDragOver={(e) => { e.preventDefault(); if (chipDragOver !== idx) setChipDragOver(idx); }}
                                    onDragLeave={() => setChipDragOver(null)}
                                    onDrop={(e) => {
                                      e.preventDefault();
                                      if (chipDragRef.current.fromIdx !== null) {
                                        reorderSelectedChips(chipDragRef.current.fromIdx, idx, selectedNames);
                                        chipDragRef.current.didDrag = true;
                                      }
                                      chipDragRef.current.fromIdx = null;
                                      setChipDragOver(null);
                                    }}
                                    onDragEnd={() => {
                                      chipDragRef.current.fromIdx = null;
                                      setChipDragOver(null);
                                    }}
                                    onTouchStart={() => {
                                      chipDragRef.current.fromIdx = idx;
                                      chipDragRef.current.didDrag = false;
                                    }}
                                    onTouchMove={(e) => {
                                      const touch = e.touches[0];
                                      const el = document.elementFromPoint(touch.clientX, touch.clientY);
                                      const chip = (el?.closest?.("[data-chip-index]")) as HTMLElement | null;
                                      if (chip) {
                                        const i = parseInt(chip.dataset.chipIndex ?? "-1");
                                        if (i >= 0 && chipDragOver !== i) setChipDragOver(i);
                                      }
                                    }}
                                    onTouchEnd={() => {
                                      const from = chipDragRef.current.fromIdx;
                                      const to = chipDragOver;
                                      if (from !== null && to !== null && from !== to) {
                                        reorderSelectedChips(from, to, selectedNames);
                                        chipDragRef.current.didDrag = true;
                                      }
                                      chipDragRef.current.fromIdx = null;
                                      setChipDragOver(null);
                                    }}
                                    onClick={() => {
                                      if (chipDragRef.current.didDrag) { chipDragRef.current.didDrag = false; return; }
                                      toggleStatus(n, modalStatus as StatusType);
                                    }}
                                    style={{
                                      display: "inline-flex", alignItems: "center", gap: "5px",
                                      padding: "5px 10px", borderRadius: "999px", background: "#fff",
                                      border: `1.5px solid ${isDragOver ? gs.color : gs.color + "55"}`,
                                      fontSize: "0.83rem", fontWeight: 700,
                                      cursor: "grab", userSelect: "none", touchAction: "none",
                                      opacity: isDragSrc ? 0.4 : 1,
                                      boxShadow: isDragOver ? `0 0 0 2.5px ${gs.color}88` : "none",
                                      transition: "box-shadow 0.1s, opacity 0.1s",
                                    }}>
                                    <span style={{ color: "#1a2035" }}>{n}</span>
                                    <span style={{ color: "#9aa3b5", fontWeight: 800, fontSize: "0.85rem", lineHeight: 1 }}>×</span>
                                  </div>
                                );
                              })}
                            </div>
                          </div>
                        );
                      })}
                    </>
                  )}
                </div>

                {/* 2. 검색 영역 */}
                <SectionLabel num="2" title="검색" />
                <div style={{ padding: "2px 14px 4px" }}>
                  {/* 2-1. 검색 결과 (그룹별) — 검색어 있을 때만 표시 */}
                  {modalSearch.trim() !== "" && <>
                  <div style={{ display: "flex", alignItems: "center", gap: "6px", padding: "6px 0 4px" }}>
                    <span style={{ fontSize: "0.68rem", fontWeight: 700, color: "#9aa3b5" }}>2-1</span>
                    <span style={{ fontSize: "0.76rem", color: "#777" }}>
                      검색결과 <strong style={{ color: "#444" }}>{filteredNames.length}</strong>명
                    </span>
                  </div>
                  <div style={{
                    border: "1px solid #eee", borderRadius: "10px",
                    overflow: "hidden", maxHeight: "260px", overflowY: "auto",
                    marginBottom: "10px",
                  }}>
                    {(() => {
                      const grouped: Record<"하우스" | "주말" | "주중", string[]> = { 하우스: [], 주말: [], 주중: [] };
                      filteredNames.forEach(n => {
                        const key = normalize(n);
                        const g = (customRosterMap[key]?.group ?? NAME_GROUP_NORMALIZED[key] ?? "하우스") as "하우스" | "주말" | "주중";
                        grouped[g].push(n);
                      });
                      if (filteredNames.length === 0) {
                        return <div style={{ textAlign: "center", color: "#bbb", padding: "20px" }}>검색 결과 없음</div>;
                      }
                      return (["하우스", "주말", "주중"] as const).map(grp => {
                        const grpNames = grouped[grp];
                        if (grpNames.length === 0) return null;
                        const gs = GROUP_STYLE[grp];
                        return (
                          <div key={grp}>
                            <div style={{
                              padding: "5px 14px 3px",
                              display: "flex", alignItems: "center", gap: "6px",
                              background: gs.bg + "66",
                              borderBottom: `1px solid ${gs.color}22`,
                            }}>
                              <span style={{
                                background: gs.bg, color: gs.color,
                                fontSize: "0.68rem", fontWeight: 800,
                                padding: "1px 8px", borderRadius: "20px",
                                border: `1px solid ${gs.color}44`,
                              }}>{grp} {grpNames.length}명</span>
                              <div style={{ flex: 1, height: 1, background: gs.color + "33" }} />
                            </div>
                            {renderItems(grpNames.slice(0, 2), false)}
                          </div>
                        );
                      });
                    })()}
                  </div>
                  </>}

                  {/* 2-2. 이름 검색 input */}
                  <input
                    value={modalSearch}
                    onChange={(e) => setModalSearch(e.target.value)}
                    placeholder="🔍 이름 검색..."
                    autoFocus
                    style={{
                      width: "100%", padding: "9px 14px", borderRadius: "12px",
                      border: `1.5px solid ${statusSc.color}44`, fontSize: "0.9rem",
                      outline: "none", boxSizing: "border-box",
                      background: statusSc.bg + "55",
                    }}
                  />
                </div>

              </div>{/* ── 스크롤 영역 끝 ── */}

              {/* 완료 버튼 - 하단 고정 */}
              <div style={{ padding: "10px 14px", borderTop: "1px solid #f0f0f0", flexShrink: 0, background: "#fff" }}>
                <button onClick={() => { setModalStatus(null); setVipSubPicking(null); }}
                  style={{
                    width: "100%", padding: "13px", borderRadius: "12px", border: "none",
                    background: "#1a2035",
                    color: "#fff",
                    fontWeight: 700, fontSize: "0.95rem", cursor: "pointer",
                  }}>
                  완료 — {selectedNames.length}명 선택됨
                </button>
              </div>
            </div>
          </div>
        );
      })()}
      {/* ── 명단 보기 모달 (당번/휴무/병가 해당자만 표시) ── */}
      {viewStatusModal && (() => {
        const st = viewStatusModal;
        const sc = STATUS_COLOR[st] ?? { bg: "#eee", color: "#333" };
        const _people = names.filter(n => effectiveStatus(n) === st);
        const people = st === "휴무"
          ? (() => {
              const dk = selectedDate?.dateLabel.slice(0, 5) ?? "";
              const order = holidayMap[dk] ?? [];
              return [..._people].sort((a, b) => {
                const ia = order.findIndex(h => normalize(h) === normalize(a));
                const ib = order.findIndex(h => normalize(h) === normalize(b));
                return (ia === -1 ? 9999 : ia) - (ib === -1 ? 9999 : ib);
              });
            })()
          : _people;
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
                    const person = getRosterPerson(name);
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
          const p = getRosterPerson(n);
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
                    const p = getRosterPerson(name);
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
      {/* ─── 일주일/3일 생성 덮어쓰기 확인 모달 ─── */}
      {weekForceConfirm > 0 && (
        <div style={{
          position: "fixed", inset: 0, zIndex: 600,
          background: "rgba(0,0,0,0.6)",
          display: "flex", alignItems: "center", justifyContent: "center",
        }}
          onClick={(e) => { if (e.target === e.currentTarget) setWeekForceConfirm(0); }}
        >
          <div style={{
            background: "#fff", borderRadius: 20, padding: "24px 20px 20px",
            maxWidth: 320, width: "90%",
            boxShadow: "0 8px 40px rgba(0,0,0,0.28)",
          }}>
            <div style={{ fontWeight: 800, fontSize: "1rem", color: "#1a1a2e", marginBottom: "6px", textAlign: "center" }}>
              📅 이미 배정된 날짜가 있습니다
            </div>
            <div style={{ fontSize: "0.82rem", color: "#666", marginBottom: "20px", textAlign: "center", lineHeight: 1.5 }}>
              선택한 {weekForceConfirm}일 범위 안에<br />이미 배정된 날짜가 포함되어 있습니다.<br />어떻게 처리할까요?
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: "10px", marginBottom: "10px" }}>
              <button
                onClick={() => {
                  const d = weekForceConfirm;
                  setWeekForceConfirm(0);
                  generateWeek(false, d);
                }}
                style={{
                  padding: "13px", borderRadius: "12px", border: "1.5px solid #374151",
                  background: "#f9fafb", color: "#374151",
                  fontWeight: 700, fontSize: "0.95rem", cursor: "pointer",
                  textAlign: "left",
                }}>
                <div>✅ 최신 상태로 재계산 (이후 날짜 유지)</div>
                <div style={{ fontSize: "0.72rem", fontWeight: 500, color: "#888", marginTop: "3px" }}>
                  선택 범위만 최신 상태로 재계산 · 이후 날짜 연쇄 갱신 없음
                </div>
              </button>
              <button
                onClick={() => {
                  const d = weekForceConfirm;
                  setWeekForceConfirm(0);
                  generateWeek(true, d);
                }}
                style={{
                  padding: "13px", borderRadius: "12px", border: "none",
                  background: "#374151", color: "#fff",
                  fontWeight: 700, fontSize: "0.95rem", cursor: "pointer",
                  textAlign: "left",
                }}>
                <div>🔄 선택 날짜부터 {weekForceConfirm}일 다시 계산</div>
                <div style={{ fontSize: "0.72rem", fontWeight: 500, color: "#9ca3af", marginTop: "3px" }}>
                  입력 데이터 유지 · 순번 체인만 새로 계산 · 이후 날짜 연쇄 갱신
                </div>
              </button>
            </div>
            <button
              onClick={() => setWeekForceConfirm(0)}
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
      {/* ─── VIP 관리 모달 ─── */}
      {vipModalOpen && (() => {
        const vipBaseNames = names.length > 0 ? names : sortedCustomRoster.map(p => p.name);
        const vipFiltered = vipSearch.trim()
          ? vipBaseNames.filter(n => n.includes(vipSearch.trim()))
          : vipBaseNames;
        return (
          <div style={{
            position: "fixed", inset: 0, zIndex: 490,
            background: "rgba(0,0,0,0.55)",
            display: "flex", flexDirection: "column", justifyContent: "flex-end",
          }}
            onClick={(e) => { if (e.target === e.currentTarget) setVipModalOpen(false); }}
          >
            <div style={{
              background: "#fff", borderRadius: "20px 20px 0 0",
              maxHeight: "82vh", display: "flex", flexDirection: "column",
              overflow: "hidden",
            }}>
              {/* 헤더 */}
              <div style={{
                padding: "16px 16px 10px",
                background: "linear-gradient(135deg, #f3e5f5 0%, #e1bee7 100%)",
                display: "flex", alignItems: "center", justifyContent: "space-between",
                flexShrink: 0,
              }}>
                <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                  <span style={{ fontWeight: 800, fontSize: "1rem", color: "#4a148c" }}>
                    👑 VIP 관리
                  </span>
                  {currentVipMembers.length > 0 && (
                    <span style={{
                      fontSize: "0.75rem", color: "#fff",
                      background: "#7b1fa2", padding: "2px 8px", borderRadius: "8px",
                      fontWeight: 700,
                    }}>
                      {currentVipMembers.length}명
                    </span>
                  )}
                  {/* VIP 팀수 컨트롤 */}
                  <div style={{ display: "flex", alignItems: "center", gap: "4px", marginLeft: "4px" }}>
                    <button
                      onClick={() => setCurrentVip({ count: Math.max(0, currentVip.count - 1) })}
                      style={{
                        width: "26px", height: "26px", borderRadius: "8px",
                        border: "1.5px solid #ce93d8", background: "#fff",
                        color: "#7b1fa2", fontSize: "1rem", fontWeight: 700, cursor: "pointer",
                        display: "flex", alignItems: "center", justifyContent: "center",
                      }}>−</button>
                    <span style={{ fontSize: "0.85rem", fontWeight: 900, color: "#7b1fa2", minWidth: "28px", textAlign: "center" }}>
                      {currentVip.count}팀
                    </span>
                    <button
                      onClick={() => setCurrentVip({ count: currentVip.count + 1 })}
                      style={{
                        width: "26px", height: "26px", borderRadius: "8px",
                        border: "1.5px solid #ce93d8", background: "#fff",
                        color: "#7b1fa2", fontSize: "1rem", fontWeight: 700, cursor: "pointer",
                        display: "flex", alignItems: "center", justifyContent: "center",
                      }}>＋</button>
                  </div>
                  {currentVipMembers.length > 0 && (
                    <button
                      onClick={() => {
                        currentVipMembers.forEach(({ name }) => clearStatus(name));
                        setCurrentVip({ count: 0 });
                      }}
                      style={{
                        padding: "3px 8px", borderRadius: "8px",
                        border: "1px solid #ef9a9a", background: "#ffebee",
                        color: "#c62828", fontSize: "0.7rem", fontWeight: 700, cursor: "pointer",
                      }}>전체삭제</button>
                  )}
                </div>
                <button onClick={() => setVipModalOpen(false)}
                  style={{
                    background: "transparent", border: "none",
                    fontSize: "1.3rem", cursor: "pointer", color: "#4a148c", lineHeight: 1,
                  }}>✕</button>
              </div>

              {/* 검색 */}
              <div style={{ padding: "10px 14px 6px", flexShrink: 0, borderBottom: "1px solid #f3f4f6" }}>
                <input
                  type="text"
                  placeholder="이름 검색..."
                  value={vipSearch}
                  onChange={(e) => setVipSearch(e.target.value)}
                  style={{
                    width: "100%", padding: "9px 12px", borderRadius: "10px",
                    border: "1.5px solid #e5e7eb", fontSize: "0.9rem",
                    outline: "none", boxSizing: "border-box",
                  }}
                />
              </div>

              {/* 인원 목록 */}
              <div style={{ overflowY: "auto", flex: 1, padding: "8px 14px 20px" }}>
                {vipFiltered.length === 0 ? (
                  <div style={{ textAlign: "center", color: "#bbb", padding: "40px 0", fontSize: "0.9rem" }}>
                    검색 결과 없음
                  </div>
                ) : (
                  vipFiltered.map((name) => {
                    const p = getRosterPerson(name);
                    const effS = effectiveStatus(name);
                    const isVipSelected = VIP_STATUSES.has(effS);
                    const vipType = isVipSelected ? effS as "VIP1부" | "VIP2부" | "VIP투근무" : null;
                    const vipSc = vipType ? STATUS_COLOR[vipType] : null;
                    return (
                      <div key={name} style={{
                        display: "flex", alignItems: "center", gap: "10px",
                        padding: "10px 12px", borderRadius: "12px", marginBottom: "7px",
                        background: isVipSelected ? "#fdf5ff" : "#fafafa",
                        border: isVipSelected ? `1.5px solid ${vipSc?.color ?? "#7b1fa2"}44` : "1px solid #f3f4f6",
                      }}>
                        {/* 아바타 */}
                        <div style={{
                          width: 32, height: 32, borderRadius: "50%", flexShrink: 0,
                          display: "flex", alignItems: "center", justifyContent: "center",
                          fontWeight: 800, fontSize: "0.82rem", color: "#fff",
                          background: isVipSelected ? (vipSc?.color ?? "#7b1fa2") : (p?.group === "주중" ? "#4e89ae" : "#9c27b0"),
                        }}>
                          {name.charAt(0)}
                        </div>

                        {/* 이름 + 그룹 */}
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontWeight: 700, fontSize: "0.9rem", color: "#1a1a2e" }}>
                            {name}
                            {vipType && (
                              <span style={{
                                marginLeft: "6px", fontSize: "0.68rem",
                                background: vipSc?.color ?? "#7b1fa2", color: "#fff",
                                padding: "1px 6px", borderRadius: "6px", fontWeight: 700,
                              }}>
                                VIP-{vipType === "VIP1부" ? "1부" : vipType === "VIP2부" ? "2부" : "투근무"}
                              </span>
                            )}
                          </div>
                          {p && (
                            <div style={{ fontSize: "0.65rem", color: GROUP_STYLE[p.group].color }}>
                              {p.조}조 · {GROUP_STYLE[p.group].label}
                            </div>
                          )}
                        </div>

                        {/* VIP 유형 버튼 */}
                        <div style={{ display: "flex", gap: "5px", flexShrink: 0 }}>
                          {(["VIP1부", "VIP2부", "VIP투근무"] as const).map((st) => {
                            const lbl = st === "VIP1부" ? "1부" : st === "VIP2부" ? "2부" : "투근무";
                            const btnSc = STATUS_COLOR[st];
                            const active = effS === st;
                            return (
                              <button key={st}
                                onClick={() => {
                                  if (active) {
                                    clearStatus(name);
                                  } else {
                                    setManualStatuses(prev => ({ ...prev, [name]: st }));
                                  }
                                }}
                                style={{
                                  padding: "5px 8px", borderRadius: "8px",
                                  border: `1.5px solid ${active ? btnSc.color : btnSc.color + "44"}`,
                                  background: active ? btnSc.color : btnSc.bg,
                                  color: active ? "#fff" : btnSc.color,
                                  fontWeight: 700, fontSize: "0.7rem", cursor: "pointer",
                                  whiteSpace: "nowrap",
                                }}>
                                {lbl}
                              </button>
                            );
                          })}
                        </div>
                      </div>
                    );
                  })
                )}
              </div>

              {/* 하단 닫기 */}
              <div style={{ padding: "10px 14px", borderTop: "1px solid #f0f0f0", flexShrink: 0 }}>
                <button onClick={() => setVipModalOpen(false)}
                  style={{
                    width: "100%", padding: "13px", borderRadius: "12px", border: "none",
                    background: "#f3e5f5", color: "#7b1fa2",
                    fontWeight: 700, fontSize: "0.95rem", cursor: "pointer",
                  }}>
                  닫기
                </button>
              </div>
            </div>
          </div>
        );
      })()}

      {/* ─── 배정 결과 (인라인 표시 — names 로드 시) ─── */}
      {names.length > 0 && (
        <>
          <div style={S.card}>
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
                        count: cho현재수, max: 6,
                        disabled: !cho가능,
                        hint: !cho가능 ? "1부 6팀+" : ""
                      },
                      {
                        status: "후출" as StatusType,
                        icon: "⬇", color: "#2196f3", bg: "#e8f4ff",
                        label: "후출", sub: `2부 뒤배치`,
                        count: hu현재수, max: 6, disabled: false, hint: ""
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
              <button
                onClick={assign}
                style={{
                  ...S.primaryBtn, flex: 1,
                  background: currentDateKey && assignmentData[currentDateKey] ? "#b45309" : undefined,
                }}>
                {currentDateKey && assignmentData[currentDateKey] ? "재배정 ⚠️" : "배정하기"}
              </button>
              <button onClick={() => generateWeek(false, 3)} style={{ ...S.primaryBtn, flex: 1, background: "#4b5563" }}>
                3일 생성
              </button>
              <button onClick={() => generateWeek()} style={{ ...S.primaryBtn, flex: 1, background: "#374151" }}>
                일주일 생성
              </button>
            </div>
            {/* 임시 결과 저장 / 취소 버튼 */}
            {pendingResult && (
              <div style={{ display: "flex", flexDirection: "column", gap: "6px", marginTop: "6px" }}>
                <div style={{ display: "flex", gap: "8px" }}>
                  <button
                    onClick={saveAssignment}
                    style={{
                      ...S.primaryBtn, flex: 1,
                      background: "linear-gradient(135deg, #16a34a, #15803d)",
                      fontSize: "0.9rem", fontWeight: 800,
                    }}>
                    💾 저장
                  </button>
                  <button
                    onClick={saveAndRecalculate}
                    style={{
                      ...S.primaryBtn, flex: 1,
                      background: "linear-gradient(135deg, #0369a1, #075985)",
                      fontSize: "0.9rem", fontWeight: 800,
                    }}>
                    💾↪ 저장+이후재계산
                  </button>
                </div>
                <button
                  onClick={() => setPendingResult(null)}
                  style={{
                    ...S.primaryBtn,
                    background: "#6b7280",
                    fontSize: "0.9rem",
                  }}>
                  ↩ 취소
                </button>
              </div>
            )}
          </div>

          {/* ── 컷 기준 요약 ── */}
          {(displayResult || (currentDateKey && assignmentData[currentDateKey]) || (livePreview && names.length > 0)) && (() => {
            // pendingResult → dayResult → livePreview → assignmentData 순 우선
            // 배정 결과(DayResultView)와 동일한 소스를 사용해야 컷 기준 요약과 배정 결과가 일치
            const cutSource = (pendingResult ?? dayResult ?? livePreview ?? (currentDateKey ? assignmentData[currentDateKey] : undefined))!;
            return (
            <div style={{
              background: "#f8f9ff", border: "1.5px solid #c5cae9", borderRadius: 12,
              padding: "10px 14px", display: "flex", flexDirection: "column", gap: 5,
            }}>
              <div style={{ fontWeight: 800, fontSize: "0.78rem", color: "#3949ab", marginBottom: 3 }}>
                ✂️ 컷 기준 요약
              </div>

              {mode === "2부제" ? (() => {
                const s1 = cutSource.shift1 ?? [];
                const s2 = cutSource.shift2 ?? [];
                const sp1 = cutSource.spare1 ?? [];
                const sp2 = cutSource.spare2 ?? [];
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
                const s1 = cutSource.shift1 ?? [];
                const sp2 = cutSource.spare2 ?? [];
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

                cutRows.push(
                  <div key="dan-sp" style={{
                    display: "flex", alignItems: "flex-start", gap: 8,
                    borderTop: "1px dashed #e0e7ff", paddingTop: 5, marginTop: 2,
                  }}>
                    <span style={{
                      fontSize: "0.7rem", fontWeight: 800, color: "#92400e",
                      background: "#fef3c7", borderRadius: 6, padding: "2px 8px",
                      minWidth: 76, textAlign: "center", flexShrink: 0,
                      marginTop: sp2.length > 0 ? 2 : 0,
                    }}>스페어</span>
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
              })()}
            </div>
            );
          })()}

          {/* 재계산 완료 배너 */}
          {recalcMessage && (
            <div style={{
              background: "linear-gradient(135deg, #dcfce7, #bbf7d0)",
              border: "1.5px solid #86efac", borderRadius: 12,
              padding: "12px 16px", display: "flex", alignItems: "center", gap: 10,
              animation: "fadeIn 0.3s ease",
            }}>
              <span style={{ fontSize: "1.2rem" }}>✅</span>
              <span style={{ fontSize: "0.85rem", fontWeight: 700, color: "#15803d" }}>
                {recalcMessage}
              </span>
            </div>
          )}

          {/* 1일 결과 */}
          {displayResult && weekly.length === 0 && (
            <div ref={resultRef} style={S.card} id="print-area">
              <div style={{ ...S.sectionTitle, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                <span>
                  📋 {selectedDate ? selectedDate.dateLabel : DAY_LABELS[dayOfWeek] + "요일"} 배정 결과
                  {pendingResult && (
                    <span style={{ marginLeft: 8, fontSize: "0.65rem", fontWeight: 700, color: "#d97706", background: "#fef3c7", borderRadius: 5, padding: "2px 7px" }}>
                      미저장
                    </span>
                  )}
                </span>
                <div style={{ display: "flex", gap: 6 }}>
                  {/* 미저장 상태면 취소, 저장 상태면 초기화 */}
                  {pendingResult ? (
                    <button
                      onClick={() => setPendingResult(null)}
                      style={{ ...S.smallBtn, fontSize: "0.7rem", padding: "4px 10px", background: "#f3f4f6", color: "#374151", border: "1px solid #d1d5db" }}
                    >
                      ↩ 취소
                    </button>
                  ) : currentDateKey && assignmentData[currentDateKey] ? (
                    <button
                      onClick={() => {
                        if (!confirm(`${currentDateKey} 배정을 초기화하시겠습니까?`)) return;
                        setAssignmentData(prev => {
                          const next = { ...prev };
                          delete next[currentDateKey];
                          return next;
                        });
                        setDayResult(null);
                      }}
                      style={{ ...S.smallBtn, fontSize: "0.7rem", padding: "4px 10px", background: "#fee2e2", color: "#b91c1c", border: "1px solid #fca5a5" }}
                    >
                      🗑 초기화
                    </button>
                  ) : null}
                  <button
                    onClick={() => window.print()}
                    style={{ ...S.smallBtn, fontSize: "0.75rem", padding: "4px 10px" }}
                  >
                    🖨️ 출력
                  </button>
                </div>
              </div>
              <DayResultView result={displayResult} mode={mode} />
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

              {weekly.map(({ day, result: r, skipped }, di) => {
                const isExpanded = expandedDays.has(day);
                const isWeekend  = di === 5 || di === 6;
                const chipBg = isWeekend
                  ? "linear-gradient(135deg, #c62828, #ef5350)"
                  : skipped
                  ? "linear-gradient(135deg, #374151, #6b7280)"
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
                        <span style={{ fontSize: "0.55rem", opacity: 0.8 }}>{skipped ? "기존" : "요일"}</span>
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
                      {skipped && (
                        <button
                          onClick={() => {
                            if (!confirm(`${day} 배정을 초기화하시겠습니까?`)) return;
                            setAssignmentData(prev => {
                              const next = { ...prev };
                              delete next[day];
                              return next;
                            });
                            setWeekly(prev => prev.filter(w => w.day !== day));
                          }}
                          style={{
                            flexShrink: 0, padding: "3px 8px",
                            border: "1px solid #fca5a5", borderRadius: 8,
                            background: "#fee2e2", color: "#b91c1c",
                            fontSize: "0.65rem", fontWeight: 700, cursor: "pointer",
                          }}
                          title="이 날 배정 초기화"
                        >
                          🗑
                        </button>
                      )}
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
                const person = getRosterPerson(name);
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
                      background: effS
                        ? (STATUS_COLOR[effS]?.bg ?? "#e5e7eb")
                        : person?.group === "하우스" ? "#52de97"
                        : person?.group === "주중" ? "#4e89ae" : "#f8b400",
                      boxShadow: "0 1px 4px rgba(0,0,0,0.12)",
                    }}>
                      <span style={{ color: effS ? (STATUS_COLOR[effS]?.color ?? "#333") : "#fff" }}>
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
                        <>
                        {STATUS_BUTTONS.filter((btn) =>
                          mode !== "단부제" || (btn !== "조출" && btn !== "후출")
                        ).map((btn) => {
                          const active = effS === btn;
                          const isAutoActive = active && isAutoHumu;
                          const disabled = (btn === "조출" && !cho가능 && effS !== "조출");
                          const col = active ? STATUS_COLOR[btn!] : null;
                          const maxReached =
                            (btn === "조출" && cho현재수 >= 6 && effS !== "조출") ||
                            (btn === "후출" && hu현재수 >= 6 && effS !== "후출");
                          return (
                            <button key={btn} disabled={disabled || maxReached}
                              onClick={() => toggleStatus(name, btn)}
                              title={
                                btn === "조출" && !cho가능 ? "1부 6팀 이상일 때만 사용 가능" :
                                btn === "조출" && maxReached ? "조출 최대 6명" :
                                btn === "후출" && maxReached ? "후출 최대 6명" : ""
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
                        {/* 엑셀 휴무 인원에게만 "해제" 버튼 표시 */}
                        {(() => {
                          const dk = currentDateKey.slice(0, 5);
                          const isInExcel = new Set((holidayMap[dk] ?? []).map(n => normalize(n))).has(normalize(name));
                          if (!isInExcel) return null;
                          const isReleased = effS === "휴무해제";
                          return (
                            <button
                              key="휴무해제"
                              onClick={() => toggleStatus(name, "휴무해제")}
                              title={isReleased ? "클릭 시 엑셀 휴무 복원" : "엑셀 휴무 해제 (가용인원에 포함)"}
                              style={{
                                ...S.statusBtn,
                                background: isReleased ? "#dcfce7" : "#fef2f2",
                                color: isReleased ? "#166534" : "#b91c1c",
                                border: isReleased ? "1.5px solid #86efac" : "1.5px solid #fca5a5",
                                fontWeight: 800,
                              }}>
                              {isReleased ? "↩복원" : "해제"}
                            </button>
                          );
                        })()}
                        </>
                      )}
                    </div>
                  </div>
                );
              });
              return rows;
            })()}

          </div>



        </>
      )}
      {/* ── 플로팅 바: 다음날 첫번호 ── */}
      {showFloatingBar && (
        <div style={S.floatingBar}>
          <img src={`${BASE}/char_smile.png`} alt="" style={{ width: 36, height: 36, objectFit: "contain" }} />
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: "0.68rem", color: "rgba(255,255,255,0.65)", marginBottom: 2 }}>🏁 내일 2부 첫번호</div>
            <div style={{ fontWeight: 800, fontSize: "1.1rem", color: "#f8b400" }}>{displayPrioritySpares[0]}</div>
          </div>
          {displayPrioritySpares[1] && (
            <div style={{ textAlign: "right" }}>
              <div style={{ fontSize: "0.65rem", color: "rgba(255,255,255,0.55)" }}>대기</div>
              <div style={{ fontWeight: 600, fontSize: "0.85rem", color: "rgba(255,255,255,0.8)" }}>{displayPrioritySpares[1]}</div>
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
      display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center",
      background: color + "15", borderRadius: "8px",
      padding: small ? "3px 4px" : "6px 4px",
      border: `1px solid ${color}33`,
      minWidth: small ? "44px" : "0",
      width: "100%",
      boxSizing: "border-box",
    }}>
      <span style={{ fontSize: small ? "0.6rem" : "0.62rem", color, fontWeight: 600, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: "100%" }}>{label}</span>
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
  const 조출Set  = new Set(result.조출List ?? []);
  const 후출Set  = new Set(result.후출List ?? []);
  const spare1Set = new Set(result.spare1 ?? []); // 1부스페어는 shift2 앞에 이미 배정 → 중복 제거용
  // 대근 인원 set (1부·2부·투라운드) — 배정 결과에서 bold 표시용
  const daegeunSet = new Set(result.daegeunList ?? []);

  function renderPeople(people: string[], key: string) {
    const isExcluded = key === "excluded";
    if (compact) {
      return people.map((n, i) => {
        const grp = isExcluded ? NAME_GROUP_NORMALIZED[normalize(n)] : undefined;
        const dot = grp ? GROUP_DOT[grp] : null;
        const invalidReason = result.invalidStatusReasons?.[n];
        return (
          <span key={n} style={{ fontWeight: daegeunSet.has(n) ? 800 : undefined }}>
            {i > 0 && "  ·  "}
            {n}
            {invalidReason && ` (${invalidReason})`}
            {dot && <span style={{ color: dot, marginLeft: "2px", fontSize: "0.65rem" }}>●</span>}
          </span>
        );
      });
    }
    return (
      <span style={{ fontSize: "0.88rem", color: "#333", lineHeight: 1.7 }}>
        {people.map((n, i) => {
          const isCho    = (key === "shift1") && 조출Set.has(n);
          const isHu     = (key === "shift2") && 후출Set.has(n);
          const isSpare1 = (key === "shift2") && spare1Set.has(n);
          const isTwoR   = key === "twoRound";
          const isDaegeun = daegeunSet.has(n); // 대근 인원 (bold 강조)
          const suffix   = isCho ? " [조출]" : isHu ? " [후출]" : isSpare1 ? " [1부스페어]" : "";
          const grp      = isExcluded ? NAME_GROUP_NORMALIZED[normalize(n)] : undefined;
          const dotColor = grp ? GROUP_DOT[grp] : null;
          const invalidReason = result.invalidStatusReasons?.[n];
          return (
            <span key={n}>
              {i > 0 && <span style={{ color: "#d1d5db" }}> · </span>}
              <span style={{
                fontWeight: (isCho || isHu || isTwoR || isSpare1 || isDaegeun) ? 800 : 500,
                color: isCho ? "#9a3412" : isHu ? "#5b21b6" : isTwoR ? "#164e63" : isSpare1 ? "#9a3412" : "#374151",
                background: isCho ? "#fed7aa" : isHu ? "#ddd6fe" : "transparent",
                borderRadius: 4, padding: (isCho || isHu) ? "1px 4px" : 0,
              }}>
                {n}{suffix}
                {invalidReason && ` (${invalidReason})`}
              </span>
              {dotColor && (
                <span style={{
                  display: "inline-block",
                  width: "7px", height: "7px",
                  borderRadius: "50%",
                  background: dotColor,
                  marginLeft: "3px",
                  verticalAlign: "middle",
                  flexShrink: 0,
                }} />
              )}
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
  page: {
    fontFamily: "'Noto Sans KR', 'Inter', sans-serif",
    background: "#eef1f8",
    minHeight: "100dvh",
    paddingBottom: "80px",
  },
  header: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
    padding: "14px 16px",
    background: "#ffffff",
    color: "#1a2035",
    position: "sticky",
    top: 0,
    zIndex: 20,
    boxShadow: "0 2px 12px rgba(100,110,180,0.10)",
    borderBottom: "1.5px solid #e8ecf4",
  },
  backBtn: {
    background: "#eef1f8",
    border: "1.5px solid #e0e5f0",
    color: "#5a6478",
    borderRadius: "10px",
    padding: "8px 14px",
    cursor: "pointer",
    fontSize: "1rem",
    minHeight: "44px",
    fontWeight: 700,
  },
  smallBtn: {
    marginLeft: "auto",
    background: "#eeebff",
    border: "1.5px solid #d4ceff",
    color: "#7c6ef7",
    borderRadius: "10px",
    padding: "8px 14px",
    cursor: "pointer",
    fontSize: "0.82rem",
    minHeight: "44px",
    fontWeight: 700,
  },
  headerTitle: {
    fontWeight: 800,
    fontSize: "1.05rem",
    letterSpacing: "-0.02em",
    color: "#1a2035",
  },

  /* ─── Cards ─── */
  card: {
    background: "#ffffff",
    borderRadius: "20px",
    padding: "18px 16px",
    margin: "10px 12px",
    boxShadow: "0 2px 16px rgba(100,110,180,0.09)",
  },
  "card1부": {
    background: "linear-gradient(135deg, #eef3ff 0%, #dde8ff 100%)",
    borderRadius: "16px",
    padding: "14px",
    border: "1.5px solid #b8ccff",
    marginBottom: "10px",
  },
  "card2부": {
    background: "linear-gradient(135deg, #f3eeff 0%, #e9e0ff 100%)",
    borderRadius: "16px",
    padding: "14px",
    border: "1.5px solid #c4b5fd",
    marginBottom: "10px",
  },
  cardSpare: {
    background: "linear-gradient(135deg, #fffbeb 0%, #fef3c7 100%)",
    borderRadius: "16px",
    padding: "14px",
    border: "1.5px solid #fcd34d",
  },

  /* ─── Typography & Labels ─── */
  label: {
    display: "block",
    fontSize: "0.7rem",
    color: "#9aa3b5",
    marginBottom: "8px",
    fontWeight: 700,
    textTransform: "uppercase",
    letterSpacing: "0.06em",
  },
  sectionTitle: {
    fontWeight: 800,
    fontSize: "0.95rem",
    marginBottom: "14px",
    color: "#1a2035",
    letterSpacing: "-0.01em",
  },

  /* ─── Mode Segment Control ─── */
  segmentTrack: {
    display: "flex",
    background: "#e8ecf4",
    borderRadius: "14px",
    padding: "4px",
    marginBottom: "18px",
    gap: "3px",
  },
  segmentBtn: {
    flex: 1,
    padding: "10px",
    border: "none",
    borderRadius: "11px",
    fontSize: "0.9rem",
    fontWeight: 700,
    cursor: "pointer",
    transition: "all 0.2s",
    minHeight: "44px",
  },

  /* ─── Day-of-week pills ─── */
  dayBtn: {
    padding: "8px 14px",
    border: "1.5px solid #e0e5f0",
    borderRadius: "12px",
    fontSize: "0.85rem",
    fontWeight: 700,
    cursor: "pointer",
    minWidth: "44px",
    minHeight: "44px",
    background: "#ffffff",
    color: "#5a6478",
  },

  /* ─── Date calendar grid ─── */
  dateGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(5, 1fr)",
    gap: "6px",
    marginBottom: "14px",
  },
  dateBtn: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "2px",
    padding: "8px 4px",
    borderRadius: "14px",
    cursor: "pointer",
    fontSize: "0.73rem",
    fontWeight: 600,
    border: "1.5px solid #e0e5f0",
    minHeight: "56px",
    transition: "all 0.15s",
    background: "#ffffff",
  },

  /* ─── Info rows ─── */
  infoRow: {
    display: "flex",
    gap: "6px",
    marginBottom: "12px",
    flexWrap: "wrap",
  },
  chip: {
    padding: "5px 12px",
    borderRadius: "20px",
    background: "#eef1f8",
    color: "#5a6478",
    fontSize: "0.78rem",
    fontWeight: 700,
    border: "1px solid #e0e5f0",
  },

  /* ─── Inputs ─── */
  textarea: {
    width: "100%",
    padding: "12px",
    borderRadius: "12px",
    border: "1.5px solid #e0e5f0",
    fontSize: "0.95rem",
    resize: "vertical",
    fontFamily: "'Noto Sans KR', 'Inter', sans-serif",
    marginBottom: "14px",
    boxSizing: "border-box",
    background: "#f8fafc",
    color: "#1a2035",
  },
  numInput: {
    width: "100%",
    padding: "12px",
    borderRadius: "12px",
    border: "1.5px solid #e0e5f0",
    fontSize: "1rem",
    boxSizing: "border-box",
    background: "#f8fafc",
    color: "#1a2035",
  },
  primaryBtn: {
    width: "100%",
    padding: "15px",
    background: "linear-gradient(135deg, #7c6ef7 0%, #5b4de8 100%)",
    color: "white",
    border: "none",
    borderRadius: "16px",
    fontSize: "1rem",
    fontWeight: 800,
    cursor: "pointer",
    boxShadow: "0 6px 20px rgba(124,110,247,0.35)",
    minHeight: "50px",
    letterSpacing: "0.01em",
  },

  /* ─── Person list ─── */
  personRow: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
    padding: "10px 0",
    borderBottom: "1.5px solid #f0f3fa",
    flexWrap: "wrap",
    transition: "opacity 0.2s",
  },
  personNum: {
    minWidth: "20px",
    fontSize: "0.72rem",
    color: "#c0c8da",
    textAlign: "right",
  },
  personName: {
    fontWeight: 700,
    fontSize: "0.88rem",
    marginBottom: "2px",
    color: "#1a2035",
  },
  btnGroup: {
    display: "flex",
    flexWrap: "wrap",
    gap: "4px",
    flex: 1,
  },
  statusBtn: {
    padding: "5px 10px",
    borderRadius: "8px",
    fontSize: "0.76rem",
    cursor: "pointer",
    fontWeight: 700,
    transition: "all 0.15s",
    minHeight: "32px",
    border: "1.5px solid #e0e5f0",
    background: "#f4f6fb",
    color: "#5a6478",
  },

  /* ─── Weekly view ─── */
  weekDay: {
    display: "flex",
    gap: "10px",
    padding: "12px 0",
    borderBottom: "1.5px solid #f0f3fa",
    alignItems: "flex-start",
  },
  dayChip: {
    minWidth: "56px",
    borderRadius: "12px",
    color: "white",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    fontWeight: 800,
    fontSize: "0.7rem",
    flexShrink: 0,
    padding: "6px 4px",
    lineHeight: 1.3,
  },

  /* ─── Info boxes ─── */
  calcBox: {
    display: "flex",
    gap: "8px",
    flexWrap: "wrap",
    alignItems: "center",
    background: "#f4f6fb",
    borderRadius: "12px",
    padding: "10px 14px",
    fontSize: "0.82rem",
    fontWeight: 600,
    border: "1.5px solid #e0e5f0",
    color: "#1a2035",
  },
  cutoffBox: {
    display: "flex",
    gap: "8px",
    flexWrap: "wrap",
    alignItems: "center",
    background: "#f8fafc",
    borderRadius: "12px",
    padding: "10px 14px",
    marginBottom: "8px",
    border: "1.5px solid #e0e5f0",
  },
  excelInfo: {
    background: "linear-gradient(135deg, #eef3ff 0%, #dde8ff 100%)",
    borderRadius: "14px",
    padding: "12px 14px",
    marginBottom: "14px",
    border: "1.5px solid #b8ccff",
  },
  excelInfoTitle: {
    fontSize: "0.78rem",
    fontWeight: 700,
    color: "#3b5fe8",
    marginBottom: "8px",
  },
  excelStatRow: {
    display: "grid",
    gridTemplateColumns: "repeat(5, 1fr)",
    gap: "6px",
  },
  dateBar: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    background: "#f4f6fb",
    borderRadius: "12px",
    padding: "10px 14px",
    marginBottom: "10px",
    flexWrap: "wrap",
    border: "1.5px solid #e0e5f0",
  },

  /* ─── Floating next-first bar (fixed bottom) ─── */
  floatingBar: {
    position: "fixed",
    bottom: 0,
    left: 0,
    right: 0,
    background: "linear-gradient(135deg, #7c6ef7 0%, #5b8dee 100%)",
    color: "white",
    padding: "12px 20px",
    zIndex: 30,
    display: "flex",
    alignItems: "center",
    gap: "10px",
    boxShadow: "0 -4px 20px rgba(124,110,247,0.30)",
  },
};
