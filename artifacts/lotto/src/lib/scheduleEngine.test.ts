import assert from "node:assert/strict";
import test from "node:test";
import { calculateSchedule, type ScheduleEngineInput, type ScheduleStatus } from "./scheduleEngine.ts";

const queue = "ABCDEFGHIJKL".split("");
const statuses = (overrides: Record<string, ScheduleStatus> = {}) =>
  Object.fromEntries(queue.map((name) => [name, overrides[name] ?? null]));
const doubleInput = (overrides: Partial<ScheduleEngineInput> = {}): ScheduleEngineInput => ({
  canonicalQueue: queue,
  mode: "2부제",
  shift1Size: 4,
  shift2Size: 4,
  statuses: statuses(),
  daegeun: {},
  ...overrides,
});

test("2부제 BASE: 1부 스페어 1명과 2부 스페어 1·2번을 분리한다", () => {
  const { base, final } = calculateSchedule(doubleInput());
  assert.deepEqual(base.shift1Membership, ["A", "B", "C", "D"]);
  assert.deepEqual(base.shift1Spare, ["E"]);
  assert.deepEqual(base.shift2Membership, ["E", "F", "G", "H"]);
  assert.deepEqual(base.shift2SpareQueue.slice(0, 2), ["I", "J"]);
  assert.deepEqual(final.nextDayQueue.slice(0, 2), ["I", "J"]);
});

test("40명 1부30 2부30에서 2부 화면도 circular membership 순서를 보존한다", () => {
  const forty = Array.from({ length: 40 }, (_, index) => String(index + 1));
  const baseStatuses = Object.fromEntries(forty.map((name) => [name, null]));
  const { final } = calculateSchedule({
    canonicalQueue: forty,
    mode: "2부제",
    shift1Size: 30,
    shift2Size: 30,
    statuses: baseStatuses,
    baseStatuses,
    daegeun: {},
  });
  const circularShift2 = [...forty.slice(30), ...forty.slice(0, 20)];
  assert.deepEqual(final.shift2Membership, circularShift2);
  assert.deepEqual(final.shift2DisplayOrder, circularShift2);
  assert.deepEqual(final.normalBothMembership, forty.slice(0, 20));
  assert.deepEqual(final.twoSpareQueue, forty.slice(20, 30));
  assert.deepEqual(final.shift2SpareQueue.slice(0, 2), ["21", "22"]);
  assert.deepEqual(final.nextDayQueue.slice(0, 2), ["21", "22"]);
});

test("원번일 2부 스페어 찾근은 1부 찾근으로 들어가고 circular를 다시 잇는다", () => {
  const requested = statuses({ I: "찾근" });
  const { final } = calculateSchedule(doubleInput({ statuses: requested, baseStatuses: statuses() }));
  assert.deepEqual(new Set(final.shift1Membership), new Set(["A", "B", "C", "I"]));
  assert.deepEqual(final.shift1Spare, ["D"]);
  assert.deepEqual(final.shift2Membership, ["D", "E", "F", "G"]);
  assert.deepEqual(final.shift2SpareQueue.slice(0, 2), ["H", "J"]);
  assert.deepEqual(final.nextDayQueue.slice(0, 2), ["H", "J"]);
  assert.deepEqual(final.appliedFinding, ["I"]);
});

test("1부 스페어 찾근은 FINAL을 다시 계산한다", () => {
  const requested = statuses({ I: "찾근" });
  const { base, final } = calculateSchedule(doubleInput({
    shift1Size: 8, shift2Size: 8, statuses: requested, baseStatuses: statuses(),
  }));
  assert.deepEqual(base.shift1Spare, ["I"]);
  assert.ok(base.shift2Membership.includes("I"));
  assert.ok(final.shift1Membership.includes("I"));
  assert.ok(final.shift2Membership.includes("I"));
  assert.deepEqual(final.appliedFinding, ["I"]);
});

test("찾근 성립자가 원래 1부 스페어 후보면 다음 정상 순번이 FINAL 1부 스페어가 된다", () => {
  const requested = statuses({ I: "찾근" });
  const { base, final } = calculateSchedule(doubleInput({
    shift1Size: 8, shift2Size: 8, statuses: requested, baseStatuses: statuses(),
  }));
  assert.deepEqual(base.shift1Spare, ["I"]);
  assert.deepEqual(final.appliedFinding, ["I"]);
  assert.deepEqual(final.shift1Spare, ["H"]);
  assert.ok(!final.shift1Spare.some((name) => final.appliedFinding.includes(name)));
});

test("찾근 성립자가 2부 스페어 후보면 다음 정상 후보가 FINAL 스페어로 승격된다", () => {
  const requested = statuses({ I: "찾근" });
  const { base, final } = calculateSchedule(doubleInput({ statuses: requested, baseStatuses: statuses() }));
  assert.deepEqual(base.shift2SpareQueue.slice(0, 2), ["I", "J"]);
  assert.deepEqual(final.shift2SpareQueue.slice(0, 2), ["H", "J"]);
  assert.ok(!final.shift2SpareQueue.includes("I"));
});

test("복수 찾근 성립자는 FINAL 1부·2부 스페어와 겹치지 않는다", () => {
  const requested = statuses({ J: "찾근", K: "찾근" });
  const { final } = calculateSchedule(doubleInput({
    shift2Size: 5, statuses: requested, baseStatuses: statuses(), previousSpare2: "F",
  }));
  const finding = new Set(final.appliedFinding);
  assert.ok(!final.shift1Spare.some((name) => finding.has(name)));
  assert.ok(!final.shift2SpareQueue.some((name) => finding.has(name)));
});

test("찾근 제외 후 새 FINAL 2부 스페어1·2가 다음날 1·2번으로 전달된다", () => {
  const requested = statuses({ I: "찾근" });
  const { final } = calculateSchedule(doubleInput({ statuses: requested, baseStatuses: statuses() }));
  assert.deepEqual(final.nextDayQueue.slice(0, 2), final.shift2SpareQueue.slice(0, 2));
  assert.deepEqual(final.nextDayQueue.slice(0, 2), ["H", "J"]);
});

test("원번일 찾근은 2부가 아니라 1부 찾근 블록에만 표시한다", () => {
  const requested = statuses({ I: "찾근" });
  const { final } = calculateSchedule(doubleInput({
    statuses: requested, baseStatuses: statuses(), previousSpare2: "F",
  }));
  assert.equal(final.shift1DisplayOrder.at(-1), "I");
  assert.ok(!final.shift2DisplayOrder.includes("I"));
});

test("1부 찾근자도 전일 스페어2 anchor 뒤 공통 순서를 따른다", () => {
  const requested = statuses({ J: "찾근" });
  const { final } = calculateSchedule(doubleInput({
    shift1Size: 8, shift2Size: 8, statuses: requested, baseStatuses: statuses(), previousSpare2: "F",
  }));
  assert.ok(final.appliedFinding.includes("J"));
  assert.equal(final.shift1DisplayOrder[final.shift1DisplayOrder.indexOf("F") + 1], "J");
});

test("복수 찾근자는 1부와 2부에서 canonical 순서의 연속 블록을 유지한다", () => {
  const requested = statuses({ G: "찾근", E: "찾근" });
  const { final } = calculateSchedule(doubleInput({
    shift1Size: 8, shift2Size: 8, statuses: requested, baseStatuses: statuses(), previousSpare2: "F",
  }));
  assert.deepEqual(final.appliedFinding, ["E", "G"]);
  const shift1Anchor = final.shift1DisplayOrder.indexOf("F");
  assert.deepEqual(final.shift1DisplayOrder.slice(shift1Anchor + 1, shift1Anchor + 3), ["E", "G"]);
  const shift2FindingStart = final.shift2DisplayOrder.indexOf("E");
  assert.deepEqual(final.shift2DisplayOrder.slice(shift2FindingStart, shift2FindingStart + 2), ["E", "G"]);
});

const findingAnchorResult = (excluded: Record<string, ScheduleStatus>, findings = ["30"]) => {
  const anchorQueue = Array.from({ length: 12 }, (_, index) => String(index + 21));
  const baseStatuses = Object.fromEntries(anchorQueue.map((name) => [name, excluded[name] ?? null]));
  const requested = { ...baseStatuses };
  findings.forEach((name) => { requested[name] = "찾근"; });
  return calculateSchedule({
    canonicalQueue: anchorQueue,
    mode: "2부제",
    shift1Size: 6,
    shift2Size: 10,
    statuses: requested,
    baseStatuses,
    requests: requested,
    daegeun: {},
    previousSpare1: "21",
    previousSpare2: "22",
  }).final;
};

test("전일 스페어2가 오늘 휴무면 다음 유효 두 번째 사람 뒤에 찾근을 배치한다", () => {
  const final = findingAnchorResult({ "22": "휴무" });
  assert.deepEqual(final.shift1DisplayOrder.slice(0, 4), ["21", "23", "30", "24"]);
});

test("전일 스페어1이 오늘 휴무면 다음 유효 두 명 중 두 번째 뒤에 찾근을 배치한다", () => {
  const final = findingAnchorResult({ "21": "휴무" });
  assert.deepEqual(final.shift1DisplayOrder.slice(0, 4), ["22", "23", "30", "24"]);
});

test("전일 스페어1·2가 모두 휴무면 다음 유효 두 명 뒤에 찾근을 배치한다", () => {
  const final = findingAnchorResult({ "21": "휴무", "22": "휴무" });
  assert.deepEqual(final.shift1DisplayOrder.slice(0, 4), ["23", "24", "30", "25"]);
});

test("전일 스페어의 병가도 유효 찾근 anchor 계산에서 제외한다", () => {
  const final = findingAnchorResult({ "22": "병가" });
  assert.deepEqual(final.shift1DisplayOrder.slice(0, 4), ["21", "23", "30", "24"]);
});

test("제외자가 없으면 전일 스페어1·2 뒤에 찾근을 배치한다", () => {
  const final = findingAnchorResult({});
  assert.deepEqual(final.shift1DisplayOrder.slice(0, 4), ["21", "22", "30", "23"]);
});

test("복수 찾근도 유효 전일 스페어2 뒤에 canonical 순서로 연속 배치한다", () => {
  const final = findingAnchorResult({ "22": "휴무" }, ["30", "31"]);
  assert.deepEqual(final.shift1DisplayOrder.slice(0, 5), ["21", "23", "30", "31", "24"]);
});

test("단부제 찾근과 뒤에서 3번째 후출 위치", () => {
  const singleQueue = "ABCDEFGH".split("");
  const baseStatuses = Object.fromEntries(singleQueue.map((n) => [n, null]));
  const requested = { ...baseStatuses, F: "찾근" as const, B: "후출" as const };
  const { final } = calculateSchedule({
    canonicalQueue: singleQueue, mode: "단부제", shift1Size: 5,
    statuses: requested, baseStatuses, daegeun: {},
  });
  assert.ok(final.shift1Membership.includes("F"));
  assert.deepEqual(final.shift2SpareQueue.slice(0, 2), ["E", "G"]);
  assert.equal(final.shift1DisplayOrder.at(-3), "B");
});

test("미성립 요청은 일반 순번으로 복귀하고 이유를 남긴다", () => {
  const requested = statuses({ A: "찾근", K: "조출", J: "후출" });
  const { final } = calculateSchedule(doubleInput({ statuses: requested, baseStatuses: statuses() }));
  assert.equal(final.invalidStatusReasons.A, "찾근 미성립 · 번호 옴");
  assert.equal(final.invalidStatusReasons.K, "조출 미성립 · 번호 안옴");
  assert.equal(final.invalidStatusReasons.J, "후출 미성립 · 번호 안옴");
  assert.ok(final.shift1Membership.includes("A"));
});

test("원번일 복수 찾근은 1부 블록에서 canonical 상대 순서를 유지한다", () => {
  const requested = statuses({ J: "찾근", K: "찾근" });
  const { final } = calculateSchedule(doubleInput({
    shift2Size: 5,
    statuses: requested,
    baseStatuses: statuses(),
    previousSpare2: "F",
  }));
  assert.deepEqual(final.shift1DisplayOrder.slice(-2), ["J", "K"]);
  assert.ok(!final.shift2DisplayOrder.includes("J") && !final.shift2DisplayOrder.includes("K"));
});

test("조출은 투스페어 최대 4번째 뒤에 표시되고 투스페어에는 섞이지 않는다", () => {
  const bigQueue = Array.from({ length: 40 }, (_, i) => String(i + 1));
  const baseStatuses = Object.fromEntries(bigQueue.map((n) => [n, null]));
  const requested = { ...baseStatuses, "5": "조출" as const };
  const { final } = calculateSchedule({
    canonicalQueue: bigQueue, mode: "2부제", shift1Size: 30, shift2Size: 30,
    statuses: requested, baseStatuses, daegeun: {},
  });
  assert.deepEqual(final.normalBothMembership, [...bigQueue.slice(0, 4), ...bigQueue.slice(5, 21)]);
  assert.deepEqual(final.twoSpareQueue.slice(0, 4), ["22", "23", "24", "25"]);
  assert.ok(!final.twoSpareQueue.includes("5"));
  const anchor = final.shift1DisplayOrder.indexOf("25");
  assert.equal(final.shift1DisplayOrder[anchor + 1], "5");
});

test("원번만 도는 날 조출은 전일 스페어2 canonical 위치 뒤에 표시한다", () => {
  const requested = statuses({ B: "조출" });
  const { final } = calculateSchedule(doubleInput({
    shift1Size: 4,
    shift2Size: 4,
    statuses: requested,
    baseStatuses: statuses(),
    previousSpare1: "K",
    previousSpare2: "C",
  }));
  assert.equal(final.bothMembership.length, 0);
  assert.equal(final.shift1DisplayOrder[final.shift1DisplayOrder.indexOf("C") + 1], "B");
});

test("투가 안 도는 2부제 후출은 정상 원번 마지막으로 이동한다", () => {
  const requested = statuses({ I: "찾근", G: "후출" });
  const { final } = calculateSchedule(doubleInput({
    statuses: requested,
    baseStatuses: statuses(),
    previousSpare2: "F",
  }));
  assert.deepEqual(final.shift2DisplayOrder, ["D", "E", "F", "G"]);
  assert.equal(final.shift2DisplayOrder.at(-1), "G");
});

test("투가 도는 날 후출은 정상 원번 구간 뒤와 정상 투근무 구간 앞에 표시된다", () => {
  const forty = Array.from({ length: 40 }, (_, index) => String(index + 1));
  const baseStatuses = Object.fromEntries(forty.map((name) => [name, null]));
  const requested = { ...baseStatuses, "5": "후출" as const };
  const { final } = calculateSchedule({
    canonicalQueue: forty,
    mode: "2부제",
    shift1Size: 30,
    shift2Size: 30,
    statuses: requested,
    baseStatuses,
    requests: requested,
    daegeun: {},
  });

  const expectedFlow = [...forty.slice(31), "5", ...forty.slice(0, 4), ...forty.slice(5, 21)];
  assert.deepEqual(final.shift2Membership, expectedFlow);
  assert.deepEqual(final.shift2DisplayOrder, expectedFlow);
  assert.deepEqual(final.normalBothMembership, [...forty.slice(0, 4), ...forty.slice(5, 21)]);
  assert.deepEqual(final.shift2SpareQueue.slice(0, 2), ["22", "23"]);
  assert.deepEqual(final.nextDayQueue.slice(0, 2), ["22", "23"]);
});

test("복수 후출은 정상 원번 구간 뒤에 canonical 순서로 연속 표시된다", () => {
  const forty = Array.from({ length: 40 }, (_, index) => String(index + 1));
  const baseStatuses = Object.fromEntries(forty.map((name) => [name, null]));
  const requested = { ...baseStatuses, "5": "후출" as const, "10": "후출" as const };
  const { final } = calculateSchedule({
    canonicalQueue: forty,
    mode: "2부제",
    shift1Size: 30,
    shift2Size: 30,
    statuses: requested,
    baseStatuses,
    requests: requested,
    daegeun: {},
  });

  const firstTwoRound = final.shift2DisplayOrder.findIndex((name) => final.normalBothMembership.includes(name));
  assert.deepEqual(final.shift2DisplayOrder.slice(firstTwoRound - 2, firstTwoRound), ["5", "10"]);
  const expectedFlow = [...forty.slice(32), "5", "10", ...forty.slice(0, 4), ...forty.slice(5, 9), ...forty.slice(10, 22)];
  assert.deepEqual(final.shift2Membership, expectedFlow);
  assert.deepEqual(final.shift2DisplayOrder, expectedFlow);
  assert.deepEqual(final.shift2SpareQueue.slice(0, 2), ["23", "24"]);
  assert.deepEqual(final.nextDayQueue.slice(0, 2), ["23", "24"]);
});

test("1부 스페어가 조출이면 스페어에서 빠지고 다음 정상 순번이 승격된다", () => {
  const requested = statuses({ E: "조출" });
  const { base, final } = calculateSchedule(doubleInput({
    statuses: requested, baseStatuses: statuses(), requests: requested,
  }));
  assert.deepEqual(base.shift1Spare, ["E"]);
  assert.deepEqual(final.appliedEarly, ["E"]);
  assert.ok(final.shift1Membership.includes("E"));
  assert.ok(!final.shift2Membership.includes("E"));
  assert.deepEqual(final.shift1Spare, ["D"]);
  assert.ok(!final.shift1Spare.some((name) => final.appliedEarly.includes(name)));
  assert.equal(final.shift2Membership[0], "D");
});

test("1부 스페어가 후출이면 새 스페어가 승격되고 후출 위치에서 근무한다", () => {
  const requested = statuses({ E: "후출" });
  const { base, final } = calculateSchedule(doubleInput({
    statuses: requested, baseStatuses: statuses(), requests: requested,
  }));
  assert.deepEqual(base.shift1Spare, ["E"]);
  assert.deepEqual(final.appliedLate, ["E"]);
  assert.ok(!final.shift1Membership.includes("E"));
  assert.ok(final.shift2Membership.includes("E"));
  assert.deepEqual(final.shift1Spare, ["F"]);
  assert.ok(!final.shift1Spare.some((name) => final.appliedLate.includes(name)));
  assert.equal(final.shift2DisplayOrder.at(-1), "E");
  assert.equal(final.shift2Membership[0], "F");
});

test("찾근·투라운드 대근·후출은 실제 위치에서 정원을 쓰고 정상 circular 중간을 건너뛰지 않는다", () => {
  const forty = Array.from({ length: 40 }, (_, index) => String(index + 1));
  const baseStatuses = Object.fromEntries(forty.map((name) => [name, null]));
  const requested = { ...baseStatuses, "25": "찾근" as const, "5": "후출" as const };
  const { final } = calculateSchedule({
    canonicalQueue: forty,
    mode: "2부제",
    shift1Size: 30,
    shift2Size: 30,
    statuses: requested,
    baseStatuses,
    requests: requested,
    requestOrder: ["25", "5"],
    daegeun: { "28": "투라운드" },
    previousSpare1: "31",
    previousSpare2: "32",
  });
  const expectedFlow = ["32", "25", "28", ...forty.slice(32), "5", ...forty.slice(0, 4), ...forty.slice(5, 19)];
  assert.deepEqual(final.shift2Membership, expectedFlow);
  assert.deepEqual(final.shift2DisplayOrder, expectedFlow);
  assert.deepEqual(final.shift1DisplayOrder.slice(0, 4), ["31", "25", "28", "1"]);
  assert.deepEqual(final.shift2SpareQueue.slice(0, 2), ["20", "21"]);
  assert.deepEqual(final.nextDayQueue.slice(0, 2), ["20", "21"]);
  assert.ok(!final.shift1Spare.some((name) => final.appliedFinding.includes(name)));
});

test("원번일 찾근은 2부 VIP 정원과 무관하게 남은 1부 자리에서 성립한다", () => {
  const base = statuses({ A: "VIP2부", B: "VIP2부", C: "VIP2부", D: "VIP2부" });
  const requested = { ...base, I: "찾근" as const };
  const { final } = calculateSchedule(doubleInput({ statuses: requested, baseStatuses: base }));
  assert.ok(final.appliedFinding.includes("I"));
  assert.ok(final.shift1Membership.includes("I"));
  assert.ok(!final.shift2Membership.includes("I"));
});

test("투라운드 날 1부만 근무하는 원번자는 찾근으로 2부에 들어간다", () => {
  const requested = statuses({ E: "찾근" });
  const { base, final } = calculateSchedule(doubleInput({
    shift1Size: 8, shift2Size: 8, statuses: requested, baseStatuses: statuses(),
  }));
  assert.ok(base.normalBothMembership.length > 0);
  assert.ok(base.shift1Membership.includes("E"));
  assert.ok(!base.shift2Membership.includes("E"));
  assert.ok(final.shift1Membership.includes("E"));
  assert.ok(final.shift2Membership.includes("E"));
  assert.ok(final.bothMembership.includes("E"));
  assert.deepEqual(final.appliedFinding, ["E"]);
});

test("투라운드 날 2부만 근무하는 원번자는 찾근으로 1부에 들어간다", () => {
  const requested = statuses({ J: "찾근" });
  const { base, final } = calculateSchedule(doubleInput({
    shift1Size: 8, shift2Size: 8, statuses: requested, baseStatuses: statuses(),
  }));
  assert.ok(base.normalBothMembership.length > 0);
  assert.ok(!base.shift1Membership.includes("J"));
  assert.ok(base.shift2Membership.includes("J"));
  assert.ok(final.shift1Membership.includes("J"));
  assert.ok(final.shift2Membership.includes("J"));
  assert.ok(final.bothMembership.includes("J"));
  assert.deepEqual(final.appliedFinding, ["J"]);
});

test("권희진·박수림 회귀: 조출 force1이 있어도 BASE 1부 찾근자는 양쪽 근무를 유지한다", () => {
  const regressionQueue = ["A", "B", "C", "D", "E", "F", "권희진", "박수림", "I", "윤다경", "이서온", "L"];
  const baseStatuses = Object.fromEntries(regressionQueue.map((name) => [name, null]));
  const requested = {
    ...baseStatuses,
    권희진: "찾근" as const,
    박수림: "찾근" as const,
    윤다경: "조출" as const,
    이서온: "조출" as const,
  };
  const { base, final } = calculateSchedule({
    canonicalQueue: regressionQueue,
    mode: "2부제",
    shift1Size: 8,
    shift2Size: 8,
    statuses: requested,
    baseStatuses,
    requests: requested,
    requestOrder: ["권희진", "박수림", "윤다경", "이서온"],
    daegeun: {},
  });

  for (const name of ["권희진", "박수림"]) {
    assert.ok(base.shift1Membership.includes(name));
    assert.ok(!base.shift2Membership.includes(name));
    assert.ok(final.shift1Membership.includes(name));
    assert.ok(final.shift2Membership.includes(name));
    assert.ok(final.bothMembership.includes(name));
    assert.ok(final.appliedFinding.includes(name));
  }
});

test("투라운드 날 정상 투스페어는 찾근이 성립한다", () => {
  const requested = statuses({ F: "찾근" });
  const { base, final } = calculateSchedule(doubleInput({
    shift1Size: 8, shift2Size: 8, statuses: requested, baseStatuses: statuses(),
  }));
  assert.ok(base.twoSpareQueue.includes("F"));
  assert.ok(final.shift2Membership.includes("F"));
  assert.ok(final.appliedFinding.includes("F"));
});

test("투라운드 날 BASE 양쪽 번호가 없는 신청자는 찾근 정원과 잔여 정원을 소비하지 않는다", () => {
  const names = Array.from({ length: 12 }, (_, index) => String(index + 1));
  const baseStatuses = Object.fromEntries(names.map((name) => [name, null]));
  baseStatuses["8"] = "대기";
  const requested = { ...baseStatuses, "4": "찾근" as const, "9": "찾근" as const };
  const { base, final } = calculateSchedule({
    canonicalQueue: names,
    mode: "2부제",
    shift1Size: 3,
    shift2Size: 8,
    statuses: requested,
    baseStatuses,
    requests: requested,
    requestOrder: ["4", "9"],
    daegeun: {},
  });

  assert.ok(base.normalBothMembership.length > 0);
  assert.ok(!base.shift1Membership.includes("4"));
  assert.ok(!base.shift2Membership.includes("4"));
  assert.equal(final.invalidStatusReasons["4"], "찾근 미성립 · 번호 안옴");
  assert.ok(!final.appliedFinding.includes("4"));
  assert.ok(!final.shift1Membership.includes("4"));
  assert.ok(!final.shift2Membership.includes("4"));
  assert.ok(final.appliedFinding.includes("9"));
  assert.equal(requested["4"], "찾근");
});

test("투라운드 날 이미 정상 투근무자는 찾근이 미성립한다", () => {
  const requested = statuses({ A: "찾근" });
  const { final } = calculateSchedule(doubleInput({
    shift1Size: 8, shift2Size: 8, statuses: requested, baseStatuses: statuses(),
  }));
  assert.equal(final.invalidStatusReasons.A, "찾근 미성립 · 이미 투번호 옴");
  assert.ok(!final.appliedFinding.includes("A"));
});

test("원번만 도는 날 원번 근무자는 찾근이 미성립한다", () => {
  const requested = statuses({ A: "찾근" });
  const { base, final } = calculateSchedule(doubleInput({ statuses: requested, baseStatuses: statuses() }));
  assert.equal(base.normalBothMembership.length, 0);
  assert.equal(final.invalidStatusReasons.A, "찾근 미성립 · 번호 옴");
});

test("원번만 도는 날 원번 미근무자는 찾근이 성립한다", () => {
  const requested = statuses({ I: "찾근" });
  const { base, final } = calculateSchedule(doubleInput({ statuses: requested, baseStatuses: statuses() }));
  assert.equal(base.normalBothMembership.length, 0);
  assert.ok(!base.shift1Membership.includes("I") && !base.shift2Membership.includes("I"));
  assert.ok(final.shift1Membership.includes("I"));
  assert.ok(!final.shift2Membership.includes("I"));
  assert.ok(final.appliedFinding.includes("I"));
});

test("함은진 회귀: 투라운드 날 1부 원번이지만 투근무가 아니면 찾근이 성립한다", () => {
  const regressionQueue = ["정상투1", "정상투2", "정상투3", "정상투4", "함은진", "F", "G", "H", "I", "J", "K", "L"];
  const baseStatuses = Object.fromEntries(regressionQueue.map((name) => [name, null]));
  const requested = { ...baseStatuses, 함은진: "찾근" as const };
  const { base, final } = calculateSchedule({
    canonicalQueue: regressionQueue, mode: "2부제", shift1Size: 8, shift2Size: 8,
    statuses: requested, baseStatuses, daegeun: {},
  });
  assert.ok(base.shift1Membership.includes("함은진"));
  assert.ok(!base.shift2Membership.includes("함은진"));
  assert.ok(base.normalBothMembership.length > 0);
  assert.ok(final.shift2Membership.includes("함은진"));
  assert.ok(final.appliedFinding.includes("함은진"));
  assert.equal(final.invalidStatusReasons.함은진, undefined);
});

test("투라운드 날 복수 찾근은 canonical 상대 순서로 성립한다", () => {
  const requested = statuses({ G: "찾근", E: "찾근" });
  const { final } = calculateSchedule(doubleInput({
    shift1Size: 8, shift2Size: 8, statuses: requested, baseStatuses: statuses(),
  }));
  assert.deepEqual(final.appliedFinding, ["E", "G"]);
  assert.ok(final.shift2Membership.includes("E"));
  assert.ok(final.shift2Membership.includes("G"));
});

test("팀 수 변경 후 같은 찾근 요청의 성립 여부를 BASE부터 다시 판정한다", () => {
  const requested = statuses({ E: "찾근" });
  const noRound = calculateSchedule(doubleInput({ statuses: requested, baseStatuses: statuses() })).final;
  const twoRound = calculateSchedule(doubleInput({
    shift1Size: 8, shift2Size: 8, statuses: requested, baseStatuses: statuses(),
  })).final;
  assert.equal(noRound.invalidStatusReasons.E, "찾근 미성립 · 번호 옴");
  assert.ok(twoRound.appliedFinding.includes("E"));
  assert.equal(twoRound.invalidStatusReasons.E, undefined);
});

test("조출은 FINAL이 아니라 BASE에서 번호가 오는지로 판정한다", () => {
  const requested = statuses({ A: "조출", I: "찾근" });
  const { base, final } = calculateSchedule(doubleInput({ statuses: requested, baseStatuses: statuses() }));
  assert.ok(base.shift1Membership.includes("A"));
  assert.ok(final.appliedEarly.includes("A"));
  assert.equal(final.invalidStatusReasons.A, undefined);
});

test("BASE에서 번호가 안 오는 조출 요청은 구체적인 이유와 함께 미성립한다", () => {
  const requested = statuses({ I: "조출" });
  const { final } = calculateSchedule(doubleInput({ statuses: requested, baseStatuses: statuses() }));
  assert.ok(!final.appliedEarly.includes("I"));
  assert.equal(final.invalidStatusReasons.I, "조출 미성립 · 번호 안옴");
});

test("2부제 후출은 어느 부든 BASE에서 번호가 오면 성립한다", () => {
  const requested = statuses({ A: "후출" });
  const { base, final } = calculateSchedule(doubleInput({ statuses: requested, baseStatuses: statuses() }));
  assert.ok(base.shift1Membership.includes("A"));
  assert.ok(!base.shift2Membership.includes("A"));
  assert.ok(final.appliedLate.includes("A"));
  assert.equal(final.invalidStatusReasons.A, undefined);
});

test("BASE에서 번호가 안 오는 후출 요청은 구체적인 이유와 함께 미성립한다", () => {
  const requested = statuses({ I: "후출" });
  const { final } = calculateSchedule(doubleInput({ statuses: requested, baseStatuses: statuses() }));
  assert.ok(!final.appliedLate.includes("I"));
  assert.equal(final.invalidStatusReasons.I, "후출 미성립 · 번호 안옴");
});

test("BASE 2부 전용 근무자는 조출 후 FINAL 1부에만 근무한다", () => {
  const requested = statuses({ E: "조출" });
  const { base, final } = calculateSchedule(doubleInput({ statuses: requested, baseStatuses: statuses() }));
  assert.ok(!base.shift1Membership.includes("E"));
  assert.ok(base.shift2Membership.includes("E"));
  assert.ok(final.shift1Membership.includes("E"));
  assert.ok(!final.shift2Membership.includes("E"));
  assert.ok(!final.bothMembership.includes("E"));
});

test("BASE 투근무자는 조출 후 FINAL 1부에만 근무한다", () => {
  const requested = statuses({ A: "조출" });
  const { base, final } = calculateSchedule(doubleInput({
    shift1Size: 8, shift2Size: 8, statuses: requested, baseStatuses: statuses(),
  }));
  assert.ok(base.bothMembership.includes("A"));
  assert.ok(final.shift1Membership.includes("A"));
  assert.ok(!final.shift2Membership.includes("A"));
  assert.ok(!final.bothMembership.includes("A"));
});

test("BASE 1부 전용 근무자는 후출 후 FINAL 2부에만 근무한다", () => {
  const requested = statuses({ A: "후출" });
  const { base, final } = calculateSchedule(doubleInput({ statuses: requested, baseStatuses: statuses() }));
  assert.ok(base.shift1Membership.includes("A"));
  assert.ok(!base.shift2Membership.includes("A"));
  assert.ok(!final.shift1Membership.includes("A"));
  assert.ok(final.shift2Membership.includes("A"));
  assert.ok(!final.bothMembership.includes("A"));
});

test("BASE 투근무자는 후출 후 FINAL 2부에만 근무한다", () => {
  const requested = statuses({ A: "후출" });
  const { base, final } = calculateSchedule(doubleInput({
    shift1Size: 8, shift2Size: 8, statuses: requested, baseStatuses: statuses(),
  }));
  assert.ok(base.bothMembership.includes("A"));
  assert.ok(!final.shift1Membership.includes("A"));
  assert.ok(final.shift2Membership.includes("A"));
  assert.ok(!final.bothMembership.includes("A"));
});

test("조출·후출 FINAL 재계산 뒤에도 2부 스페어와 다음날 1·2번은 같은 기존 계산식을 따른다", () => {
  const requested = statuses({ A: "후출", E: "조출" });
  const { final } = calculateSchedule(doubleInput({ statuses: requested, baseStatuses: statuses() }));
  assert.deepEqual(final.nextDayQueue.slice(0, 2), final.shift2SpareQueue.slice(0, 2));
  assert.equal(final.shift1Membership.length, 4);
  assert.equal(final.shift2Membership.length, 4);
});

test("VIP/대근 강제 좌석과 병가 차단을 BASE에 반영한다", () => {
  const base = statuses({ A: "VIP투근무", B: "병가" });
  const { final } = calculateSchedule(doubleInput({
    statuses: base,
    baseStatuses: base,
    daegeun: { C: "투라운드", B: "투라운드" },
  }));
  assert.ok(final.bothMembership.includes("A"));
  assert.ok(final.bothMembership.includes("C"));
  assert.ok(final.excluded.includes("B"));
  assert.ok(!final.shift1Membership.includes("B"));
});

test("병가가 요청을 덮어도 요청은 유지되고 구체적인 미성립 이유를 표시한다", () => {
  const resolved = statuses({ I: "병가" });
  const base = statuses({ I: "병가" });
  const requests = statuses({ I: "찾근" });
  const { final } = calculateSchedule(doubleInput({ statuses: resolved, baseStatuses: base, requests }));
  assert.equal(final.invalidStatusReasons.I, "찾근 미성립 · 병가");
  assert.ok(final.excluded.includes("I"));
});

test("팀 수 변경은 BASE부터 스페어와 다음날 큐를 다시 계산한다", () => {
  const four = calculateSchedule(doubleInput()).final;
  const five = calculateSchedule(doubleInput({ shift1Size: 5, shift2Size: 5 })).final;
  assert.notDeepEqual(four.shift1Spare, five.shift1Spare);
  assert.notDeepEqual(four.shift2SpareQueue.slice(0, 2), five.shift2SpareQueue.slice(0, 2));
  assert.deepEqual(five.nextDayQueue.slice(0, 2), five.shift2SpareQueue.slice(0, 2));
});

test("다음날 전체 순번은 2부 스페어 전체가 끊기지 않고 이어진다", () => {
  const { final } = calculateSchedule(doubleInput());
  assert.deepEqual(final.shift2SpareQueue, ["I", "J", "K", "L", "A", "B", "C", "D"]);
  assert.deepEqual(final.nextDayQueue, ["I", "J", "K", "L", "A", "B", "C", "D", "E", "F", "G", "H"]);
});

test("다음날 순번에서 휴무·병가 등 근무 불가자는 건너뛰어 뒤로 보낸다", () => {
  const baseStatuses = statuses({ B: "휴무", C: "병가" });
  const { final } = calculateSchedule(doubleInput({ statuses: baseStatuses, baseStatuses }));
  assert.deepEqual(final.nextDayQueue.slice(-2), ["B", "C"]);
  assert.ok(final.nextDayQueue.indexOf("D") < final.nextDayQueue.indexOf("B"));
});

test("VIP1부는 2부 근무와 2부 스페어 후보에서 제외한다", () => {
  const baseStatuses = statuses({ H: "VIP1부", K: "VIP2부" });
  const { final } = calculateSchedule(doubleInput({ statuses: baseStatuses, baseStatuses }));
  assert.ok(final.shift1Membership.includes("H"));
  assert.ok(!final.shift2Membership.includes("H"));
  assert.ok(!final.shift2SpareQueue.includes("H"));
});

test("대근1부는 2부 근무와 2부 스페어 후보에서 제외한다", () => {
  const { final } = calculateSchedule(doubleInput({ daegeun: { H: "1부", K: "2부", L: "투라운드" } }));
  assert.ok(final.shift1Membership.includes("H"));
  assert.ok(!final.shift2Membership.includes("H"));
  assert.ok(!final.shift2SpareQueue.includes("H"));
  assert.ok(final.shift2Membership.includes("K"));
  assert.ok(final.bothMembership.includes("L"));
});

test("2부제 full rotation은 마지막 실제 근무자 다음 유효 순번부터 이어간다", () => {
  const fullQueue = "ABCDEFGH".split("");
  const baseStatuses = Object.fromEntries(fullQueue.map((name) => [name, null]));
  const { final } = calculateSchedule({
    canonicalQueue: fullQueue, mode: "2부제", shift1Size: 4, shift2Size: 8,
    statuses: baseStatuses, baseStatuses, daegeun: {},
  });
  assert.deepEqual(final.shift2Membership, ["E", "F", "G", "H", "A", "B", "C", "D"]);
  assert.deepEqual(final.shift2SpareQueue, ["F", "G"]);
  assert.deepEqual(final.nextDayQueue, ["F", "G", "H", "A", "B", "C", "D"]);
});

test("단부제 full rotation은 마지막 실제 근무자 다음 유효 순번부터 이어간다", () => {
  const fullQueue = "ABCDEFGH".split("");
  const baseStatuses = Object.fromEntries(fullQueue.map((name) => [name, name === "C" ? "휴무" : null]));
  const { final } = calculateSchedule({
    canonicalQueue: fullQueue, mode: "단부제", shift1Size: 7,
    statuses: baseStatuses, baseStatuses, daegeun: {},
  });
  assert.deepEqual(final.shift2SpareQueue, []);
  assert.deepEqual(final.nextDayQueue, ["A", "B", "D", "E", "F", "G", "H", "C"]);
});

test("엔진은 조출과 후출 특수 효과를 각각 최대 6명까지만 적용한다", () => {
  const manyQueue = Array.from({ length: 20 }, (_, index) => String(index + 1));
  const plain = Object.fromEntries(manyQueue.map((name) => [name, null]));
  const earlyRequests = { ...plain };
  const lateRequests = { ...plain };
  manyQueue.slice(0, 7).forEach((name) => { earlyRequests[name] = "조출"; });
  manyQueue.slice(0, 7).forEach((name) => { lateRequests[name] = "후출"; });
  const early = calculateSchedule({
    canonicalQueue: manyQueue, mode: "2부제", shift1Size: 10, shift2Size: 10,
    statuses: earlyRequests, baseStatuses: plain, requests: earlyRequests, daegeun: {},
  }).final;
  const late = calculateSchedule({
    canonicalQueue: manyQueue, mode: "2부제", shift1Size: 10, shift2Size: 10,
    statuses: lateRequests, baseStatuses: plain, requests: lateRequests, daegeun: {},
  }).final;
  assert.deepEqual(early.appliedEarly, manyQueue.slice(0, 6));
  assert.deepEqual(late.appliedLate, manyQueue.slice(0, 6));
  assert.ok(!early.appliedEarly.includes("7"));
  assert.ok(!late.appliedLate.includes("7"));
});

test("3일/7일 연쇄는 전체 nextDayQueue를 다음 canonicalQueue로 전달할 수 있다", () => {
  let current = queue;
  for (let day = 0; day < 7; day++) {
    const result = calculateSchedule(doubleInput({ canonicalQueue: current })).final;
    assert.equal(result.nextDayQueue.length, queue.length);
    assert.equal(new Set(result.nextDayQueue).size, queue.length);
    current = result.nextDayQueue;
  }
});

test("3일 생성은 날짜별 팀 설정과 전날 FINAL 스페어 기준으로 하루 배정 3회와 일치한다", () => {
  const roster = Array.from({ length: 50 }, (_, index) => String(index + 1));
  const rotateFrom = (start?: string) => {
    if (!start) return [...roster];
    const index = roster.indexOf(start);
    return index < 0 ? [...roster] : [...roster.slice(index), ...roster.slice(0, index)];
  };
  const dayInputs = [
    {
      mode: "2부제" as const, shift1Size: 22, shift2Size: 20,
      requests: { "2": "휴무", "8": "조출", "24": "후출", "45": "찾근", "46": "찾근" } as Record<string, ScheduleStatus>,
      requestOrder: ["8", "24", "46", "45"], daegeun: { "47": "투라운드" as const },
    },
    {
      mode: "2부제" as const, shift1Size: 18, shift2Size: 14,
      requests: { "2": "휴무해제", "11": "조출", "31": "후출", "42": "찾근" } as Record<string, ScheduleStatus>,
      requestOrder: ["11", "31", "42"], daegeun: { "44": "2부" as const },
    },
    {
      mode: "단부제" as const, shift1Size: 27, shift2Size: 0,
      requests: { "5": "병가", "9": "조출", "33": "후출", "48": "찾근" } as Record<string, ScheduleStatus>,
      requestOrder: ["9", "33", "48"], daegeun: {},
    },
  ];
  const calculateDay = (
    day: typeof dayInputs[number],
    canonicalQueue: string[],
    previous?: ReturnType<typeof calculateSchedule>["final"],
  ) => {
    const statusesForDay = Object.fromEntries(roster.map(name => [name, day.requests[name] ?? null]));
    const baseStatuses = Object.fromEntries(roster.map(name => {
      const status = day.requests[name] ?? null;
      return [name, status === "조출" || status === "후출" || status === "찾근" ? null : status];
    }));
    return calculateSchedule({
      canonicalQueue,
      mode: day.mode,
      shift1Size: day.shift1Size,
      shift2Size: day.shift2Size,
      statuses: statusesForDay,
      baseStatuses,
      requests: statusesForDay,
      requestOrder: day.requestOrder,
      daegeun: day.daegeun,
      previousSpare1: previous?.shift2SpareQueue[0],
      previousSpare2: previous?.shift2SpareQueue[1],
    }).final;
  };
  const runLikeThreeManualAssignments = () => {
    const results = [];
    let previous: ReturnType<typeof calculateSchedule>["final"] | undefined;
    for (const day of dayInputs) {
      const canonicalQueue = rotateFrom(previous?.shift2SpareQueue[0]);
      previous = calculateDay(day, canonicalQueue, previous);
      results.push(previous);
    }
    return results;
  };
  const manual = runLikeThreeManualAssignments();
  const generated = runLikeThreeManualAssignments();
  for (let day = 0; day < 3; day++) {
    assert.deepEqual(generated[day].shift1DisplayOrder, manual[day].shift1DisplayOrder);
    assert.deepEqual(generated[day].shift1Spare, manual[day].shift1Spare);
    assert.deepEqual(generated[day].shift2DisplayOrder, manual[day].shift2DisplayOrder);
    assert.deepEqual(generated[day].shift2SpareQueue, manual[day].shift2SpareQueue);
    assert.deepEqual(generated[day].nextDayQueue, manual[day].nextDayQueue);
  }
  const legacySecondDay = calculateDay(dayInputs[1], manual[0].nextDayQueue, manual[0]);
  assert.notDeepEqual(legacySecondDay.shift1DisplayOrder, manual[1].shift1DisplayOrder);
});

test("정상 투근무 11명은 신청순서 앞 9명까지만 찾근이 성립한다", () => {
  const names = Array.from({ length: 40 }, (_, index) => String(index + 1));
  const plain = Object.fromEntries(names.map(name => [name, null]));
  const applicants = names.slice(11, 23).reverse();
  const requests = { ...plain };
  applicants.forEach(name => { requests[name] = "찾근"; });
  const { base, final } = calculateSchedule({
    canonicalQueue: names, mode: "2부제", shift1Size: 30, shift2Size: 21,
    statuses: requests, baseStatuses: plain, requests, requestOrder: applicants, daegeun: {},
  });
  assert.equal(base.normalBothMembership.length, 11);
  assert.deepEqual(final.appliedFinding, applicants.slice(0, 9));
  assert.equal(final.invalidStatusReasons[applicants[9]], "찾근 미성립 · 정원 초과");
});

test("정상 투근무 7명은 신청순서 앞 5명까지만 찾근이 성립한다", () => {
  const names = Array.from({ length: 40 }, (_, index) => String(index + 1));
  const plain = Object.fromEntries(names.map(name => [name, null]));
  const applicants = names.slice(7, 15).reverse();
  const requests = { ...plain };
  applicants.forEach(name => { requests[name] = "찾근"; });
  const { base, final } = calculateSchedule({
    canonicalQueue: names, mode: "2부제", shift1Size: 30, shift2Size: 17,
    statuses: requests, baseStatuses: plain, requests, requestOrder: applicants, daegeun: {},
  });
  assert.equal(base.normalBothMembership.length, 7);
  assert.deepEqual(final.appliedFinding, applicants.slice(0, 5));
  assert.equal(final.invalidStatusReasons[applicants[5]], "찾근 미성립 · 정원 초과");
});

test("이미 정상 투근무인 찾근 신청자는 동적 정원을 소비하지 않는다", () => {
  const names = Array.from({ length: 40 }, (_, index) => String(index + 1));
  const plain = Object.fromEntries(names.map(name => [name, null]));
  const applicants = ["1", "15", "14", "13", "12", "11"];
  const requests = { ...plain };
  applicants.forEach(name => { requests[name] = "찾근"; });
  const { final } = calculateSchedule({
    canonicalQueue: names, mode: "2부제", shift1Size: 30, shift2Size: 17,
    statuses: requests, baseStatuses: plain, requests, requestOrder: applicants, daegeun: {},
  });
  assert.equal(final.invalidStatusReasons["1"], "찾근 미성립 · 이미 투번호 옴");
  assert.deepEqual(final.appliedFinding, ["15", "14", "13", "12", "11"]);
});

test("팀 수 증감은 저장된 찾근 신청순서를 유지한 채 동적 정원을 다시 계산한다", () => {
  const names = Array.from({ length: 40 }, (_, index) => String(index + 1));
  const plain = Object.fromEntries(names.map(name => [name, null]));
  const applicants = names.slice(10, 18).reverse();
  const requests = { ...plain };
  applicants.forEach(name => { requests[name] = "찾근"; });
  const run = (shift2Size: number) => calculateSchedule({
    canonicalQueue: names, mode: "2부제", shift1Size: 30, shift2Size,
    statuses: requests, baseStatuses: plain, requests, requestOrder: applicants, daegeun: {},
  }).final;
  const small = run(16);
  const large = run(20);
  const smallAgain = run(16);
  assert.deepEqual(small.appliedFinding, applicants.slice(0, 4));
  assert.deepEqual(large.appliedFinding, applicants);
  assert.deepEqual(smallAgain.appliedFinding, applicants.slice(0, 4));
  assert.deepEqual(applicants, names.slice(10, 18).reverse());
});

test("조출·후출 미성립 요청은 입력을 변경하지 않고 조건 변경 후 다시 성립한다", () => {
  const requests = statuses({ I: "조출", J: "후출" });
  const before = structuredClone(requests);
  const small = calculateSchedule(doubleInput({ statuses: requests, baseStatuses: statuses(), requests })).final;
  const large = calculateSchedule(doubleInput({
    shift1Size: 10, shift2Size: 10, statuses: requests, baseStatuses: statuses(), requests,
  })).final;
  assert.equal(small.invalidStatusReasons.I, "조출 미성립 · 번호 안옴");
  assert.equal(small.invalidStatusReasons.J, "후출 미성립 · 번호 안옴");
  assert.ok(large.appliedEarly.includes("I"));
  assert.ok(large.appliedLate.includes("J"));
  assert.deepEqual(requests, before);
});

test("원번일 조출·후출·복수 찾근은 각 위치에서 정원을 쓰고 다음날 순번을 잇는다", () => {
  const names = Array.from({ length: 12 }, (_, index) => String(index + 1));
  const baseStatuses = Object.fromEntries(names.map((name) => [name, null]));
  const requested = {
    ...baseStatuses,
    "6": "조출" as const,
    "3": "후출" as const,
    "9": "찾근" as const,
    "10": "찾근" as const,
  };
  const { final } = calculateSchedule({
    canonicalQueue: names, mode: "2부제", shift1Size: 4, shift2Size: 4,
    statuses: requested, baseStatuses, requests: requested,
    requestOrder: ["6", "3", "9", "10"], daegeun: {},
    previousSpare1: "1", previousSpare2: "2",
  });
  assert.deepEqual(final.shift1Membership, ["6", "9", "10", "1"]);
  assert.deepEqual(final.shift1Spare, ["2"]);
  assert.deepEqual(final.shift2Membership, ["2", "4", "5", "3"]);
  assert.deepEqual(final.shift2SpareQueue.slice(0, 2), ["7", "8"]);
  assert.deepEqual(final.nextDayQueue.slice(0, 2), ["7", "8"]);
  assert.deepEqual(final.appliedEarly, ["6"]);
  assert.deepEqual(final.appliedLate, ["3"]);
  assert.deepEqual(final.appliedFinding, ["9", "10"]);
});

test("투라운드일 조출·후출·복수 찾근은 기존 위치와 BASE 찾근 정원을 유지한다", () => {
  const names = Array.from({ length: 12 }, (_, index) => String(index + 1));
  const baseStatuses = Object.fromEntries(names.map((name) => [name, null]));
  const requested = {
    ...baseStatuses,
    "10": "조출" as const,
    "7": "후출" as const,
    "5": "찾근" as const,
    "6": "찾근" as const,
  };
  const { base, final } = calculateSchedule({
    canonicalQueue: names, mode: "2부제", shift1Size: 8, shift2Size: 8,
    statuses: requested, baseStatuses, requests: requested,
    requestOrder: ["10", "7", "5", "6"], daegeun: {},
    previousSpare1: "9", previousSpare2: "10",
  });
  assert.equal(base.normalBothMembership.length - 2, 2);
  assert.deepEqual(final.shift1Spare, ["9"]);
  assert.deepEqual(final.shift2Membership, ["9", "5", "6", "11", "12", "7", "1", "2"]);
  assert.deepEqual(final.shift2SpareQueue, ["3", "4", "8"]);
  assert.deepEqual(final.nextDayQueue.slice(0, 2), ["3", "4"]);
  assert.deepEqual(final.appliedEarly, ["10"]);
  assert.deepEqual(final.appliedLate, ["7"]);
  assert.deepEqual(final.appliedFinding, ["5", "6"]);
});

test("찾근 재계산으로 번호를 잃은 조출은 미성립하고 후출 신청은 자기 효과 없이 재검증한다", () => {
  const names = Array.from({ length: 12 }, (_, index) => String(index + 1));
  const baseStatuses = Object.fromEntries(names.map((name) => [name, null]));
  const requested = {
    ...baseStatuses,
    "8": "조출" as const,
    "4": "후출" as const,
    "9": "찾근" as const,
    "10": "찾근" as const,
  };
  const { final } = calculateSchedule({
    canonicalQueue: names, mode: "2부제", shift1Size: 4, shift2Size: 4,
    statuses: requested, baseStatuses, requests: requested,
    requestOrder: ["8", "4", "9", "10"], daegeun: {},
    previousSpare1: "1", previousSpare2: "2",
  });
  assert.deepEqual(final.shift1Membership, ["9", "10", "1", "2"]);
  assert.deepEqual(final.shift1Spare, ["3"]);
  assert.deepEqual(final.shift2Membership, ["3", "5", "6", "4"]);
  assert.deepEqual(final.shift2SpareQueue.slice(0, 2), ["7", "8"]);
  assert.deepEqual(final.nextDayQueue.slice(0, 2), ["7", "8"]);
  assert.deepEqual(final.appliedEarly, []);
  assert.deepEqual(final.appliedLate, ["4"]);
  assert.deepEqual(final.appliedFinding, ["9", "10"]);
  assert.equal(final.invalidStatusReasons["8"], "조출 미성립 · 번호 안옴");
});

test("원번일 찾근 적용 후 번호를 잃은 기존 신청자를 다시 판정한다", () => {
  const names = Array.from({ length: 50 }, (_, index) => String(index + 1));
  const baseStatuses = Object.fromEntries(names.map((name) => [name, null]));
  const requested = { ...baseStatuses, "35": "찾근" as const, "36": "찾근" as const, "37": "찾근" as const };
  const { base, final } = calculateSchedule({
    canonicalQueue: names, mode: "2부제", shift1Size: 22, shift2Size: 13,
    statuses: requested, baseStatuses, requests: requested,
    requestOrder: ["35", "36", "37"], daegeun: {},
    previousSpare1: "36", previousSpare2: "37",
  });
  assert.ok(base.shift2Membership.includes("35"));
  assert.ok(!base.shift1Membership.includes("36") && !base.shift2Membership.includes("36"));
  assert.deepEqual(final.appliedFinding, ["35", "36", "37"]);
  assert.deepEqual(final.shift1Spare, ["20"]);
  assert.deepEqual(final.shift2SpareQueue.slice(0, 2), ["33", "34"]);
  assert.deepEqual(final.nextDayQueue.slice(0, 2), ["33", "34"]);
});

test("투라운드 찾근 적용 후 한 번 근무로 밀린 기존 신청자를 다시 판정한다", () => {
  const names = Array.from({ length: 40 }, (_, index) => String(index + 1));
  const baseStatuses = Object.fromEntries(names.map((name) => [name, null]));
  const requested = { ...baseStatuses, "20": "찾근" as const, "21": "찾근" as const };
  const { base, final } = calculateSchedule({
    canonicalQueue: names, mode: "2부제", shift1Size: 30, shift2Size: 30,
    statuses: requested, baseStatuses, requests: requested,
    requestOrder: ["20", "21"], daegeun: {},
    previousSpare1: "21", previousSpare2: "22",
  });
  assert.ok(base.normalBothMembership.includes("20"));
  assert.ok(!base.normalBothMembership.includes("21"));
  assert.deepEqual(final.appliedFinding, ["20", "21"]);
  assert.ok(final.bothMembership.includes("20") && final.bothMembership.includes("21"));
  assert.deepEqual(final.shift2SpareQueue.slice(0, 2), ["19", "22"]);
  assert.deepEqual(final.nextDayQueue.slice(0, 2), ["19", "22"]);
});
