import assert from "node:assert/strict";
import test from "node:test";
import * as XLSX from "xlsx";
import { calculateSchedule, type ScheduleStatus } from "./scheduleEngine.ts";
import {
  applyOcrAssignments,
  applyTimingImportState,
  clearManualHolidayImports,
  limitTimingAssignments,
  matchImportedStatuses,
  mergeHolidayImport,
  mergeTimingRequestOrder,
  parseOcrStatusText,
  parseTimingExcelBuffer,
} from "./scheduleInputImport.ts";

test("상태 Excel의 날짜·이름·조출/후출/찾근을 순서대로 읽는다", () => {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([
    ["날짜", "이름", "상태"],
    ["09.26", "김 철수", "조출"],
    ["09.26", "박영희", "후출"],
    ["09.26", "이민수", "찾근"],
  ]);
  XLSX.utils.book_append_sheet(workbook, sheet, "입력");
  const bytes = XLSX.write(workbook, { type: "array", bookType: "xlsx" }) as ArrayBuffer;
  assert.deepEqual(parseTimingExcelBuffer(bytes, 9), {
    "09.26": { 조출: ["김철수"], 후출: ["박영희"], 찾근: ["이민수"] },
  });
});

test("OCR은 상태별 이름을 분리한다", () => {
  assert.deepEqual(parseOcrStatusText("휴무: 김철수 박영희\n조출 이민수\n후출: 정하늘\n찾근 권희진 신현진"), {
    휴무: ["김철수", "박영희"], 조출: ["이민수"], 후출: ["정하늘"], 찾근: ["권희진", "신현진"],
  });
});

test("roster는 공백 제거 후 정확히 일치하는 이름만 자동 매칭한다", () => {
  const result = matchImportedStatuses({ 조출: ["김 철수", "김철슈"] }, ["김철수", "박영희"]);
  assert.deepEqual(result.matched, { 조출: ["김철수"] });
  assert.deepEqual(result.needsReview, { 조출: ["김철슈"] });
});

test("휴무 Excel 선택 날짜 적용은 선택 날짜만 교체한다", () => {
  const current = { "09.25": ["가"], "09.26": ["나"], "09.27": ["다"] };
  const incoming = { "09.26": ["라"], "09.27": ["마"] };
  assert.deepEqual(mergeHolidayImport(current, incoming, ["09.26"], "selected"), {
    "09.25": ["가"], "09.26": ["라"], "09.27": ["다"],
  });
  assert.deepEqual(current, { "09.25": ["가"], "09.26": ["나"], "09.27": ["다"] });
});

test("휴무 Excel 적용은 대상 휴무만 지우고 다른 날짜와 수동 상태를 보존한다", () => {
  const current = {
    "09.25 (금)": { 가: "휴무", 나: "조출" },
    "09.26 (토)": { 다: "휴무", 라: "후출", 마: "찾근" },
    "09.27 (일)": { 바: "휴무" },
  };
  assert.deepEqual(clearManualHolidayImports(current, ["09.26"], ["09.26"], "selected"), {
    "09.25 (금)": { 가: "휴무", 나: "조출" },
    "09.26 (토)": { 라: "후출", 마: "찾근" },
    "09.27 (일)": { 바: "휴무" },
  });
});

test("상태 Excel 선택 날짜 적용은 휴무와 다른 날짜를 변경하지 않는다", () => {
  const current = {
    "09.25 (금)": { 가: "조출" },
    "09.26 (토)": { 나: "휴무", 다: "후출" },
    "09.27 (일)": { 라: "찾근" },
  };
  const result = applyTimingImportState(
    current,
    { "09.26 (토)": { 다: "후출" } },
    { "09.26 (토)": { 다: "조출", 마: "찾근" } },
    ["09.26 (토)"],
    ["09"],
    "selected",
  );
  assert.deepEqual(result.statuses, {
    "09.25 (금)": { 가: "조출" },
    "09.26 (토)": { 나: "휴무", 다: "조출", 마: "찾근" },
    "09.27 (일)": { 라: "찾근" },
  });
});

test("상태 Excel 후 수동 변경·삭제한 값은 새 Excel이 덮어쓰지 않는다", () => {
  const result = applyTimingImportState(
    { "09.26 (토)": { 가: "찾근", 다: "휴무" } },
    { "09.26 (토)": { 가: "조출", 나: "후출" } },
    { "09.26 (토)": { 가: "후출", 나: "조출", 라: "찾근" } },
    ["09.26 (토)"],
    ["09"],
    "selected",
  );
  assert.deepEqual(result.statuses["09.26 (토)"], { 가: "찾근", 다: "휴무", 라: "찾근" });
  assert.deepEqual(result.source["09.26 (토)"], { 라: "찾근" });
});

test("OCR 검토와 roster 매칭은 적용 전 저장 데이터를 변경하지 않는다", () => {
  const stored = { "09.25 (금)": { 가: "휴무" } };
  const before = structuredClone(stored);
  const parsed = parseOcrStatusText("조출: 김철수 김철슈");
  const matched = matchImportedStatuses(parsed, ["김철수"]);
  assert.deepEqual(stored, before);
  assert.deepEqual(matched.matched, { 조출: ["김철수"] });
  assert.deepEqual(matched.needsReview, { 조출: ["김철슈"] });
});

test("OCR 적용은 촬영 시작 때 고정한 날짜 하나와 정확 매칭 이름만 변경한다", () => {
  const current = {
    "09.25 (금)": { 가: "휴무" },
    "09.26 (토)": { 나: "후출" },
    "09.27 (일)": { 다: "찾근" },
  };
  const matched = matchImportedStatuses({ 조출: ["김철수", "김철슈"] }, ["김철수"]);
  const assignments = Object.fromEntries(
    Object.entries(matched.matched).flatMap(([status, names]) => (names ?? []).map(name => [name, status])),
  );
  const result = applyOcrAssignments(current, "09.26 (토)", assignments as Record<string, "조출">);
  assert.deepEqual(result, {
    "09.25 (금)": { 가: "휴무" },
    "09.26 (토)": { 나: "후출", 김철수: "조출" },
    "09.27 (일)": { 다: "찾근" },
  });
  assert.equal(result["09.26 (토)"]["김철슈"], undefined);
});

test("Excel/OCR 적용 함수는 dateStatusOrders를 읽거나 변경하지 않는다", () => {
  const dateStatusOrders = { "09.26 (토)": ["다", "나", "가"] };
  const before = structuredClone(dateStatusOrders);
  applyTimingImportState({}, {}, { "09.26 (토)": { 가: "조출" } }, ["09.26 (토)"], ["09"], "selected");
  applyOcrAssignments({}, "09.26 (토)", { 나: "찾근" });
  assert.deepEqual(dateStatusOrders, before);
});

test("상태 입력 순서가 달라도 배정 결과는 canonical 번호 순서를 사용한다", () => {
  const canonicalQueue = ["A", "B", "C", "D", "E", "F", "G", "H"];
  const forward = Object.fromEntries(canonicalQueue.map(name => [name, null])) as Record<string, ScheduleStatus>;
  forward.B = "조출";
  forward.F = "후출";
  const reverse = Object.fromEntries([...canonicalQueue].reverse().map(name => [name, forward[name]])) as Record<string, ScheduleStatus>;
  const input = { canonicalQueue, mode: "2부제" as const, shift1Size: 4, shift2Size: 4, daegeun: {} };
  const first = calculateSchedule({ ...input, statuses: forward, baseStatuses: Object.fromEntries(canonicalQueue.map(name => [name, null])) });
  const second = calculateSchedule({ ...input, statuses: reverse, baseStatuses: Object.fromEntries(canonicalQueue.map(name => [name, null])) });
  assert.deepEqual(second, first);
});

test("Excel/OCR 조출·후출 입력은 기존 신청을 포함해 최대 6명까지만 받는다", () => {
  const current = { 기존조출: "조출", 기존후출: "후출" };
  const incoming: Array<[string, "조출" | "후출" | "찾근"]> = [
    ...Array.from({ length: 7 }, (_, index) => [`조출${index + 1}`, "조출"] as [string, "조출"]),
    ...Array.from({ length: 7 }, (_, index) => [`후출${index + 1}`, "후출"] as [string, "후출"]),
    ["찾근1", "찾근"],
  ];
  const result = limitTimingAssignments(current, incoming);
  assert.equal(Object.values(result.assignments).filter(status => status === "조출").length, 5);
  assert.equal(Object.values(result.assignments).filter(status => status === "후출").length, 5);
  assert.equal(result.assignments.조출6, undefined);
  assert.equal(result.assignments.후출6, undefined);
  assert.equal(result.assignments.찾근1, "찾근");
});

test("Excel/OCR 찾근 순서를 기존 dateStatusOrders에 이어서 보존한다", () => {
  const statuses = { 수동찾근: "찾근", 엑셀B: "찾근", 엑셀A: "찾근", 조출자: "조출" };
  const order = mergeTimingRequestOrder(
    ["조출자", "수동찾근"],
    statuses,
    ["엑셀B", "엑셀A"],
  );
  assert.deepEqual(order, ["조출자", "수동찾근", "엑셀B", "엑셀A"]);
  assert.deepEqual(order.filter(name => statuses[name as keyof typeof statuses] === "찾근"), ["수동찾근", "엑셀B", "엑셀A"]);
});
