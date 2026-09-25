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

test("2부 스페어 찾근은 정원을 유지하며 정상 마지막 근무자를 민다", () => {
  const requested = statuses({ I: "찾근" });
  const { final } = calculateSchedule(doubleInput({ statuses: requested, baseStatuses: statuses() }));
  assert.deepEqual(new Set(final.shift2Membership), new Set(["E", "F", "G", "I"]));
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

test("스페어 후보 제외 후에도 찾근 표시 위치는 전일 스페어2 뒤를 유지한다", () => {
  const requested = statuses({ I: "찾근" });
  const { final } = calculateSchedule(doubleInput({
    statuses: requested, baseStatuses: statuses(), previousSpare2: "F",
  }));
  const anchor = final.shift2DisplayOrder.indexOf("F");
  assert.equal(final.shift2DisplayOrder[anchor + 1], "I");
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

test("복수 찾근은 클릭 순서가 아닌 canonical 상대 순서로 anchor 뒤에 표시한다", () => {
  const requested = statuses({ J: "찾근", K: "찾근" });
  const { final } = calculateSchedule(doubleInput({
    shift2Size: 5,
    statuses: requested,
    baseStatuses: statuses(),
    previousSpare2: "F",
  }));
  const f = final.shift2DisplayOrder.indexOf("F");
  assert.deepEqual(final.shift2DisplayOrder.slice(f + 1, f + 3), ["J", "K"]);
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

test("2부제 후출은 찾근 배치 후 원번 마지막으로 이동한다", () => {
  const requested = statuses({ I: "찾근", G: "후출" });
  const { final } = calculateSchedule(doubleInput({
    statuses: requested,
    baseStatuses: statuses(),
    previousSpare2: "F",
  }));
  assert.deepEqual(final.shift2DisplayOrder, ["E", "F", "I", "G"]);
  assert.equal(final.shift2DisplayOrder.at(-1), "G");
});

test("VIP/대근이 정원을 채우면 찾근은 미성립한다", () => {
  const base = statuses({ A: "VIP2부", B: "VIP2부", C: "VIP2부", D: "VIP2부" });
  const requested = { ...base, I: "찾근" as const };
  const { final } = calculateSchedule(doubleInput({ statuses: requested, baseStatuses: base }));
  assert.equal(final.invalidStatusReasons.I, "찾근 미성립 · VIP/대근 근무로 정원 초과");
  assert.ok(!final.appliedFinding.includes("I"));
});

test("투라운드 날 1부만 근무하는 원번자는 찾근으로 2부에 들어간다", () => {
  const requested = statuses({ E: "찾근" });
  const { base, final } = calculateSchedule(doubleInput({
    shift1Size: 8, shift2Size: 8, statuses: requested, baseStatuses: statuses(),
  }));
  assert.ok(base.normalBothMembership.length > 0);
  assert.ok(base.shift1Membership.includes("E"));
  assert.ok(!base.shift2Membership.includes("E"));
  assert.ok(final.shift2Membership.includes("E"));
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
  assert.deepEqual(final.appliedFinding, ["J"]);
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
  assert.ok(final.shift2Membership.includes("I"));
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

test("3일/7일 연쇄는 전체 nextDayQueue를 다음 canonicalQueue로 전달할 수 있다", () => {
  let current = queue;
  for (let day = 0; day < 7; day++) {
    const result = calculateSchedule(doubleInput({ canonicalQueue: current })).final;
    assert.equal(result.nextDayQueue.length, queue.length);
    assert.equal(new Set(result.nextDayQueue).size, queue.length);
    current = result.nextDayQueue;
  }
});
