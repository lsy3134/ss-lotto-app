import assert from "node:assert/strict";
import test from "node:test";
import { calculateSchedule } from "./scheduleEngine";
import {
  buildBaseSavedDay,
  mergeLegacyHolidayReleases,
  setHolidayRelease,
  type HolidayReleaseMap,
} from "./scheduleHolidayRelease";

const queue = ["A", "B", "C", "D", "E", "F", "G", "H"];

function runReleasedTiming(status: "조출" | "후출" | "찾근") {
  const releases = setHolidayRelease({}, "09.30", "E", true);
  const savedDay = { E: status };
  const baseSavedDay = buildBaseSavedDay(savedDay, releases["09.30"]);
  const result = calculateSchedule({
    canonicalQueue: queue,
    mode: "2부제",
    shift1Size: 4,
    shift2Size: 8,
    statuses: savedDay,
    baseStatuses: baseSavedDay,
    requests: savedDay,
    requestOrder: ["E"],
    daegeun: {},
  });
  return { releases, baseSavedDay, result };
}

function roundTrip<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

for (const status of ["조출", "후출", "찾근"] as const) {
  test(`Excel 빨간 해제 → ${status}: release 기록이 유지되어 BASE에서 휴무가 되살아나지 않는다`, () => {
    const { releases, baseSavedDay, result } = runReleasedTiming(status);
    assert.equal(releases["09.30"].E, true);
    assert.equal(baseSavedDay.E, "휴무해제");
    assert.ok(!result.base.excluded.includes("E"));
    if (status === "조출") assert.deepEqual(result.final.appliedEarly, ["E"]);
    if (status === "후출") assert.deepEqual(result.final.appliedLate, ["E"]);
    if (status === "찾근") assert.deepEqual(result.final.appliedFinding, ["E"]);
  });
}

test("휴무해제 → 찾근 → 다시 휴무: 해제 기록을 제거하면 정상 근무 제외된다", () => {
  let releases: HolidayReleaseMap = setHolidayRelease({}, "09.30", "E", true);
  releases = setHolidayRelease(releases, "09.30", "E", false);
  const savedDay = { E: "휴무" as const };
  const result = calculateSchedule({
    canonicalQueue: queue,
    mode: "2부제",
    shift1Size: 4,
    shift2Size: 8,
    statuses: savedDay,
    baseStatuses: buildBaseSavedDay(savedDay, releases["09.30"]),
    requests: savedDay,
    requestOrder: ["E"],
    daegeun: {},
  });
  assert.equal(releases["09.30"], undefined);
  assert.ok(result.base.excluded.includes("E"));
  assert.deepEqual(result.final.appliedFinding, []);
});

test("휴무해제 + 찾근은 저장·새로고침 왕복 후에도 유지된다", () => {
  const before = mergeLegacyHolidayReleases(
    setHolidayRelease({}, "09.30", "E", true),
    { "09.30": { E: "찾근" } },
  );
  const after = JSON.parse(JSON.stringify(before)) as HolidayReleaseMap;
  assert.equal(after["09.30"].E, true);
  assert.equal(buildBaseSavedDay({ E: "찾근" }, after["09.30"]).E, "휴무해제");
});

for (const status of ["조출", "후출", "찾근"] as const) {
  test(`휴무 목록 이름 삭제 → ${status}: 저장·새로고침 후에도 release와 신청이 함께 유지된다`, () => {
    const releases = roundTrip(setHolidayRelease({}, "09.30", "E", true));
    const savedDay = roundTrip({ E: status });
    const baseSavedDay = buildBaseSavedDay(savedDay, releases["09.30"]);

    assert.equal(releases["09.30"].E, true);
    assert.equal(savedDay.E, status);
    assert.equal(baseSavedDay.E, "휴무해제");
  });
}

test("휴무 목록 이름 삭제와 빨간 해제는 동일한 release 상태와 찾근 결과를 만든다", () => {
  const listRemovalReleases = setHolidayRelease({}, "09.30", "E", true);
  const redButtonReleases = setHolidayRelease({}, "09.30", "E", true);
  const savedDay = { E: "찾근" as const };

  assert.deepEqual(listRemovalReleases, redButtonReleases);
  assert.deepEqual(
    buildBaseSavedDay(savedDay, listRemovalReleases["09.30"]),
    buildBaseSavedDay(savedDay, redButtonReleases["09.30"]),
  );
  assert.deepEqual(runReleasedTiming("찾근").result.final.appliedFinding, ["E"]);
});

test("휴무해제 후 휴무 목록에 다시 추가하면 release가 제거되고 다시 근무 제외된다", () => {
  const released = setHolidayRelease({}, "09.30", "E", true);
  const restored = setHolidayRelease(released, "09.30", "E", false);
  const savedDay = { E: "휴무" as const };
  const result = calculateSchedule({
    canonicalQueue: queue,
    mode: "2부제",
    shift1Size: 4,
    shift2Size: 8,
    statuses: savedDay,
    baseStatuses: buildBaseSavedDay(savedDay, restored["09.30"]),
    requests: savedDay,
    requestOrder: [],
    daegeun: {},
  });

  assert.equal(restored["09.30"], undefined);
  assert.ok(result.base.excluded.includes("E"));
});

for (const status of ["조출", "후출", "찾근"] as const) {
  test(`기존 정상 직원의 ${status}에는 휴무해제 보조 맵이 영향을 주지 않는다`, () => {
    const savedDay = { E: status };
    const baseSavedDay = buildBaseSavedDay(savedDay);
    const result = calculateSchedule({
      canonicalQueue: queue,
      mode: "2부제",
      shift1Size: 4,
      shift2Size: 8,
      statuses: savedDay,
      baseStatuses: baseSavedDay,
      requests: savedDay,
      requestOrder: ["E"],
      daegeun: {},
    });

    assert.deepEqual(baseSavedDay, {});
    assert.equal(result.base.excluded.includes("E"), false);
  });
}

test("기존 휴무해제 저장값은 보조 맵으로 호환 이관된다", () => {
  const migrated = mergeLegacyHolidayReleases({}, { "09.30": { E: "휴무해제" } });
  assert.equal(migrated["09.30"].E, true);
});
