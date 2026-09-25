export type ScheduleMode = "2부제" | "단부제";
export type ScheduleStatus =
  | "조출" | "후출" | "찾근" | "대기"
  | "당번" | "병가" | "휴무" | "하우스" | "휴무해제"
  | "VIP1부" | "VIP2부" | "VIP투근무"
  | null;
export type ScheduleDaegeun = "1부" | "2부" | "투라운드";

export interface ScheduleEngineInput {
  canonicalQueue: string[];
  mode: ScheduleMode;
  shift1Size: number;
  shift2Size?: number;
  statuses: Record<string, ScheduleStatus>;
  baseStatuses?: Record<string, ScheduleStatus>;
  requests?: Record<string, ScheduleStatus>;
  requestOrder?: string[];
  daegeun: Record<string, ScheduleDaegeun>;
  previousSpare1?: string;
  previousSpare2?: string;
}

export interface ScheduleEngineResult {
  shift1Membership: string[];
  shift2Membership: string[];
  shift1Spare: string[];
  shift2SpareQueue: string[];
  bothMembership: string[];
  normalBothMembership: string[];
  twoSpareQueue: string[];
  shift1DisplayOrder: string[];
  shift2DisplayOrder: string[];
  appliedFinding: string[];
  appliedEarly: string[];
  appliedLate: string[];
  invalidStatusReasons: Record<string, string>;
  excluded: string[];
  vip1: string[];
  vip2: string[];
  vipBoth: string[];
  daegeunNames: string[];
  nextDayQueue: string[];
}

type Allocation = Omit<ScheduleEngineResult,
  "appliedFinding" | "appliedEarly" | "appliedLate" | "invalidStatusReasons" |
  "shift1DisplayOrder" | "shift2DisplayOrder" | "nextDayQueue"> & {
    shift2RegularOrder: string[];
  };

const BLOCKED = new Set<ScheduleStatus>(["당번", "병가", "휴무", "하우스"]);
const TIMING = new Set<ScheduleStatus>(["찾근", "조출", "후출"]);

function unique(names: string[]): string[] {
  return [...new Set(names)];
}

function canonicalSort(names: Iterable<string>, queue: string[]): string[] {
  const set = new Set(names);
  return queue.filter((name) => set.has(name));
}

function rotateAfter(queue: string[], anchor?: string): string[] {
  if (!anchor) return [...queue];
  const index = queue.indexOf(anchor);
  if (index < 0) return [...queue];
  return [...queue.slice(index + 1), ...queue.slice(0, index + 1)];
}

function addAfterAnchor(
  order: string[], movers: string[], canonicalQueue: string[], anchor?: string,
): string[] {
  const moving = canonicalSort(movers, canonicalQueue);
  if (!moving.length) return order;
  const movingSet = new Set(moving);
  const rest = order.filter((name) => !movingSet.has(name));
  if (!anchor) return [...rest, ...moving];
  const direct = rest.indexOf(anchor);
  if (direct >= 0) return [...rest.slice(0, direct + 1), ...moving, ...rest.slice(direct + 1)];
  const anchorIndex = canonicalQueue.indexOf(anchor);
  if (anchorIndex < 0) return [...rest, ...moving];
  const insertion = rest.findIndex((name) => canonicalQueue.indexOf(name) > anchorIndex);
  const at = insertion < 0 ? rest.length : insertion;
  return [...rest.slice(0, at), ...moving, ...rest.slice(at)];
}

function validFindingAnchor(
  canonicalQueue: string[], excluded: Iterable<string>, previousSpare1?: string, previousSpare2?: string,
): string | undefined {
  if (!previousSpare2) return previousSpare1;
  if (!previousSpare1) return previousSpare2;
  const start = canonicalQueue.indexOf(previousSpare1);
  if (start < 0) return previousSpare2;
  const excludedSet = new Set(excluded);
  const previousOrder = [...canonicalQueue.slice(start), ...canonicalQueue.slice(0, start)];
  const valid = previousOrder.filter((name) => !excludedSet.has(name));
  return valid[1] ?? valid[0];
}

function makeNextDayQueue(spares: string[], canonicalQueue: string[], excluded: string[]): string[] {
  const spareOrder = unique(spares);
  const spareSet = new Set(spareOrder);
  const excludedSet = new Set(excluded);
  return [
    ...spareOrder,
    ...canonicalQueue.filter((name) => !spareSet.has(name) && !excludedSet.has(name)),
    ...canonicalQueue.filter((name) => excludedSet.has(name)),
  ];
}

function baseStatus(status: ScheduleStatus): ScheduleStatus {
  return TIMING.has(status) ? null : status;
}

function allocateSingle(
  input: ScheduleEngineInput,
  forcedFinding: Set<string> = new Set(),
  excludeFromSpares: Set<string> = new Set(),
): Allocation {
  const queue = input.canonicalQueue;
  const excluded = queue.filter((name) => BLOCKED.has(baseStatus(input.statuses[name] ?? null)));
  const excludedSet = new Set(excluded);
  const vip = queue.filter((name) => {
    const status = baseStatus(input.statuses[name] ?? null);
    return !excludedSet.has(name) && (status === "VIP1부" || status === "VIP2부" || status === "VIP투근무");
  });
  const waiting = queue.filter((name) => baseStatus(input.statuses[name] ?? null) === "대기" && !excludedSet.has(name));
  const fixed = unique([...vip, ...canonicalSort(forcedFinding, queue)]).slice(0, input.shift1Size);
  const fixedSet = new Set(fixed);
  const waitingSet = new Set(waiting);
  const regular = queue.filter((name) => !excludedSet.has(name) && !fixedSet.has(name) && !waitingSet.has(name));
  const shift1Membership = unique([...fixed, ...regular.slice(0, Math.max(0, input.shift1Size - fixed.length))]);
  const shift1Set = new Set(shift1Membership);
  const spareCandidates = [...waiting, ...queue.filter((name) => !excludedSet.has(name) && !shift1Set.has(name) && !waitingSet.has(name))]
    .filter((name) => !excludeFromSpares.has(name));
  return {
    shift1Membership,
    shift2Membership: [],
    shift1Spare: [],
    shift2SpareQueue: spareCandidates,
    bothMembership: [],
    normalBothMembership: [],
    twoSpareQueue: [],
    excluded,
    vip1: vip,
    vip2: [],
    vipBoth: [],
    daegeunNames: [],
    shift2RegularOrder: [],
  };
}

function allocateDouble(
  input: ScheduleEngineInput,
  forceShift1: Set<string> = new Set(),
  forceShift2: Set<string> = new Set(),
  excludeFromShift1: Set<string> = new Set(),
  excludeFromShift2: Set<string> = new Set(),
  excludeFromSpares: Set<string> = new Set(),
): Allocation {
  const queue = input.canonicalQueue;
  const s1Capacity = Math.max(0, input.shift1Size);
  const s2Capacity = Math.max(0, input.shift2Size ?? 0);
  const excluded = queue.filter((name) => BLOCKED.has(baseStatus(input.statuses[name] ?? null)));
  const excludedSet = new Set(excluded);
  const statusOf = (name: string) => baseStatus(input.statuses[name] ?? null);
  const vip1 = queue.filter((n) => !excludedSet.has(n) && statusOf(n) === "VIP1부");
  const vip2 = queue.filter((n) => !excludedSet.has(n) && statusOf(n) === "VIP2부");
  const vipBoth = queue.filter((n) => !excludedSet.has(n) && statusOf(n) === "VIP투근무");
  const dg1 = queue.filter((n) => !excludedSet.has(n) && input.daegeun[n] === "1부");
  const dg2 = queue.filter((n) => !excludedSet.has(n) && input.daegeun[n] === "2부");
  const dgBoth = queue.filter((n) => !excludedSet.has(n) && input.daegeun[n] === "투라운드");
  const waiting = queue.filter((n) => !excludedSet.has(n) && statusOf(n) === "대기");
  const dedicated1 = new Set([...vip1, ...dg1]);
  const dedicated2 = new Set([...vip2, ...dg2]);
  const fixed1 = unique([...vip1, ...vipBoth, ...dg1, ...dgBoth, ...canonicalSort(forceShift1, queue)])
    .filter((n) => !excludedSet.has(n) && !dedicated2.has(n) && !excludeFromShift1.has(n)).slice(0, s1Capacity);
  const fixed1Set = new Set(fixed1);
  const waitingSet = new Set(waiting);
  const normal1Candidates = queue.filter((n) => !excludedSet.has(n) && !fixed1Set.has(n) && !dedicated2.has(n) && !waitingSet.has(n) && !excludeFromShift1.has(n));
  const normal1Selected = normal1Candidates.slice(0, Math.max(0, s1Capacity - fixed1.length));
  const shift1Membership = unique([...fixed1, ...normal1Selected]);
  const shift1Set = new Set(shift1Membership);
  const shift1Spare = [...waiting, ...queue.filter((n) => !excludedSet.has(n) && !shift1Set.has(n) && !waitingSet.has(n))]
    .filter((n) => !excludeFromSpares.has(n))
    .slice(0, 1);

  const fixed2 = unique([...vip2, ...vipBoth, ...dg2, ...dgBoth, ...canonicalSort(forceShift2, queue), ...shift1Spare])
    .filter((n) => !excludedSet.has(n) && !excludeFromShift2.has(n)).slice(0, s2Capacity);
  const fixed2Set = new Set(fixed2);
  const regularCandidateOrder = rotateAfter(queue, shift1Spare[0])
    .filter((n) => !excludedSet.has(n) && !dedicated1.has(n) && !fixed2Set.has(n) && !excludeFromShift2.has(n));
  const regularSelected = regularCandidateOrder.slice(0, Math.max(0, s2Capacity - fixed2.length));
  const shift2Membership = unique([...fixed2, ...regularSelected]);
  const shift2Set = new Set(shift2Membership);
  const shift2SpareQueue = regularCandidateOrder
    .slice(regularSelected.length)
    .filter((n) => !shift2Set.has(n) && !excludeFromSpares.has(n));
  const bothMembership = queue.filter((n) => shift1Set.has(n) && shift2Set.has(n));
  const regularSelectedSet = new Set(regularSelected);
  const normalBothMembership = bothMembership.filter((n) => regularSelectedSet.has(n));
  const forcedCause = new Set([...vip1, ...vip2, ...vipBoth, ...dg1, ...dg2, ...dgBoth, ...forceShift1, ...forceShift2]);
  const normal1Set = new Set(normal1Selected);
  const twoSpareQueue = queue.filter((n) => normal1Set.has(n) && !shift2Set.has(n) && !forcedCause.has(n));
  return {
    shift1Membership,
    shift2Membership,
    shift1Spare,
    shift2SpareQueue,
    bothMembership,
    normalBothMembership,
    twoSpareQueue,
    excluded,
    vip1,
    vip2,
    vipBoth,
    daegeunNames: unique([...dg1, ...dg2, ...dgBoth]),
    shift2RegularOrder: regularSelected,
  };
}

function fixedCount(input: ScheduleEngineInput, shift: 1 | 2): number {
  return input.canonicalQueue.filter((name) => {
    if (BLOCKED.has(baseStatus(input.statuses[name] ?? null))) return false;
    const status = baseStatus(input.statuses[name] ?? null);
    const dg = input.daegeun[name];
    return shift === 1
      ? status === "VIP1부" || status === "VIP투근무" || dg === "1부" || dg === "투라운드"
      : status === "VIP2부" || status === "VIP투근무" || dg === "2부" || dg === "투라운드";
  }).length;
}

export function calculateSchedule(input: ScheduleEngineInput): { base: ScheduleEngineResult; final: ScheduleEngineResult } {
  const queue = unique(input.canonicalQueue);
  const normalized: ScheduleEngineInput = { ...input, canonicalQueue: queue };
  const allocationInput: ScheduleEngineInput = {
    ...normalized,
    statuses: input.baseStatuses ?? input.statuses,
  };
  const baseAllocation = input.mode === "2부제" ? allocateDouble(allocationInput) : allocateSingle(allocationInput);
  const requestSource = input.requests ?? input.statuses;
  const requestOrder = unique([...(input.requestOrder ?? []), ...queue]).filter((name) => queue.includes(name));
  const requestedFinding = requestOrder.filter((n) => requestSource[n] === "찾근");
  const requestedEarly = requestOrder.filter((n) => requestSource[n] === "조출").slice(0, 6);
  const requestedLate = requestOrder.filter((n) => requestSource[n] === "후출").slice(0, 6);
  const invalidStatusReasons: Record<string, string> = {};
  const force1 = new Set<string>();
  const force2 = new Set<string>();
  const exclude1 = new Set<string>();
  const exclude2 = new Set<string>();
  let appliedFinding: string[] = [];
  let free1 = Math.max(0, input.shift1Size - fixedCount(allocationInput, 1));
  let free2 = Math.max(0, (input.shift2Size ?? 0) - fixedCount(allocationInput, 2) - baseAllocation.shift1Spare.length);
  const normalTwoRoundDay = input.mode === "2부제" && baseAllocation.normalBothMembership.length > 0;
  const findingLimit = normalTwoRoundDay
    ? Math.max(0, baseAllocation.normalBothMembership.length - 2)
    : Number.POSITIVE_INFINITY;
  const baseS1 = new Set(baseAllocation.shift1Membership);
  const baseS2 = new Set(baseAllocation.shift2Membership);
  const baseNumbered = new Set([...baseS1, ...baseS2]);
  const baseBoth = new Set(baseAllocation.bothMembership);
  const appliedEarly = requestedEarly.filter((name) => {
    const valid = baseNumbered.has(name);
    if (!valid) invalidStatusReasons[name] = "조출 미성립 · 번호 안옴";
    else if (input.mode === "2부제") {
      force1.add(name);
      exclude2.add(name);
    }
    return valid;
  });
  const appliedLate = requestedLate.filter((name) => {
    const valid = baseNumbered.has(name);
    if (!valid) invalidStatusReasons[name] = "후출 미성립 · 번호 안옴";
    else if (input.mode === "2부제") {
      force2.add(name);
      exclude1.add(name);
    }
    return valid;
  });

  for (const name of requestedFinding) {
    if (baseAllocation.excluded.includes(name)) {
      invalidStatusReasons[name] = `찾근 미성립 · ${input.statuses[name] === "병가" ? "병가" : "근무 제외"}`;
    } else if (input.mode === "단부제") {
      if (baseS1.has(name)) invalidStatusReasons[name] = "찾근 미성립 · 번호 옴";
      else if (free1 <= 0) invalidStatusReasons[name] = "찾근 미성립 · VIP/대근 근무로 정원 초과";
      else { force1.add(name); appliedFinding.push(name); free1--; }
    } else if (normalTwoRoundDay) {
      if (baseBoth.has(name)) {
        invalidStatusReasons[name] = "찾근 미성립 · 이미 투번호 옴";
      } else if (appliedFinding.length >= findingLimit) {
        invalidStatusReasons[name] = "찾근 미성립 · 정원 초과";
      } else if (baseS1.has(name)) {
        if (free2 <= 0) invalidStatusReasons[name] = "찾근 미성립 · VIP/대근 근무로 정원 초과";
        else { force2.add(name); appliedFinding.push(name); free2--; }
      } else if (baseS2.has(name)) {
        if (free1 <= 0) invalidStatusReasons[name] = "찾근 미성립 · VIP/대근 근무로 정원 초과";
        else { force1.add(name); appliedFinding.push(name); free1--; }
      } else if (free2 <= 0) {
        invalidStatusReasons[name] = "찾근 미성립 · VIP/대근 근무로 정원 초과";
      } else {
        force2.add(name); appliedFinding.push(name); free2--;
      }
    } else if (baseS1.has(name) || baseS2.has(name)) {
      invalidStatusReasons[name] = "찾근 미성립 · 번호 옴";
    } else if (free2 <= 0) {
      invalidStatusReasons[name] = "찾근 미성립 · VIP/대근 근무로 정원 초과";
    } else {
      force2.add(name); appliedFinding.push(name); free2--;
    }
  }

  let finalAllocation = input.mode === "2부제"
    ? allocateDouble(allocationInput, force1, force2, exclude1, exclude2, new Set(appliedFinding))
    : allocateSingle(allocationInput, force1, new Set(appliedFinding));
  const failedFinding = appliedFinding.filter((name) =>
    (force1.has(name) && !finalAllocation.shift1Membership.includes(name)) ||
    (force2.has(name) && !finalAllocation.shift2Membership.includes(name))
  );
  if (failedFinding.length > 0) {
    failedFinding.forEach((name) => {
      force1.delete(name);
      force2.delete(name);
      invalidStatusReasons[name] = "찾근 미성립 · VIP/대근 근무로 정원 초과";
    });
    appliedFinding = appliedFinding.filter((name) => !failedFinding.includes(name));
    finalAllocation = input.mode === "2부제"
      ? allocateDouble(allocationInput, force1, force2, exclude1, exclude2, new Set(appliedFinding))
      : allocateSingle(allocationInput, force1, new Set(appliedFinding));
  }
  const finalS1 = new Set(finalAllocation.shift1Membership);
  const finalS2 = new Set(finalAllocation.shift2Membership);
  const findingAnchor = validFindingAnchor(
    queue, baseAllocation.excluded, input.previousSpare1, input.previousSpare2,
  );

  const findingDisplayOrder = input.mode === "2부제" && appliedFinding.length
    ? addAfterAnchor(queue, appliedFinding, queue, findingAnchor)
    : queue;
  let shift1DisplayOrder = findingDisplayOrder.filter((n) => finalS1.has(n));
  let shift2DisplayOrder = [...finalAllocation.shift2Membership];
  if (input.mode === "2부제" && appliedEarly.length) {
    const earlyAnchor = finalAllocation.bothMembership.length > 0 && finalAllocation.twoSpareQueue.length > 0
      ? finalAllocation.twoSpareQueue[Math.min(3, finalAllocation.twoSpareQueue.length - 1)]
      : input.previousSpare2 ?? input.previousSpare1;
    shift1DisplayOrder = addAfterAnchor(shift1DisplayOrder, appliedEarly, queue, earlyAnchor);
  }
  if (input.mode === "2부제" && appliedFinding.length) {
    shift2DisplayOrder = addAfterAnchor(
      shift2DisplayOrder,
      appliedFinding.filter((n) => finalS2.has(n)),
      queue,
      findingAnchor,
    );
  }
  if (appliedLate.length) {
    if (input.mode === "단부제") {
      const lateSet = new Set(appliedLate);
      const rest = shift1DisplayOrder.filter((n) => !lateSet.has(n));
      const at = Math.max(0, rest.length - 2);
      shift1DisplayOrder = [...rest.slice(0, at), ...canonicalSort(appliedLate, queue), ...rest.slice(at)];
    } else {
      const lateSet = new Set(appliedLate);
      let rest = shift2DisplayOrder.filter((name) => !lateSet.has(name));
      const normalBothSet = new Set(finalAllocation.normalBothMembership);
      const firstOriginalIndex = rest.findIndex((name) => !normalBothSet.has(name));
      if (firstOriginalIndex > 0 && rest.slice(0, firstOriginalIndex).every((name) => normalBothSet.has(name))) {
        rest = [...rest.slice(firstOriginalIndex), ...rest.slice(0, firstOriginalIndex)];
      }
      const firstTwoRoundIndex = rest.findIndex((name) => normalBothSet.has(name));
      const at = firstTwoRoundIndex >= 0 ? firstTwoRoundIndex : rest.length;
      shift2DisplayOrder = [
        ...rest.slice(0, at),
        ...canonicalSort(appliedLate, queue),
        ...rest.slice(at),
      ];
    }
  }

  const decorate = (allocation: Allocation, isFinal: boolean): ScheduleEngineResult => {
    let shift2SpareQueue = allocation.shift2SpareQueue;
    let nextDayQueue: string[];
    if (shift2SpareQueue.length > 0) {
      nextDayQueue = makeNextDayQueue(shift2SpareQueue, queue, allocation.excluded);
    } else {
      const excludedSet = new Set(allocation.excluded);
      const shift1SpareSet = new Set(input.mode === "2부제" ? allocation.shift1Spare : []);
      const validQueue = queue.filter((name) => !excludedSet.has(name) && !shift1SpareSet.has(name));
      const lastWorked = input.mode === "2부제"
        ? allocation.shift2Membership.at(-1)
        : allocation.shift1Membership.at(-1);
      const rotated = rotateAfter(validQueue, lastWorked);
      nextDayQueue = [...rotated, ...queue.filter((name) => excludedSet.has(name))];
      if (input.mode === "2부제") shift2SpareQueue = rotated.slice(0, 2);
    }
    return {
      ...allocation,
      shift2SpareQueue,
      shift1DisplayOrder: isFinal ? shift1DisplayOrder : allocation.shift1Membership,
      shift2DisplayOrder: isFinal ? shift2DisplayOrder : allocation.shift2Membership,
      appliedFinding: isFinal ? appliedFinding : [],
      appliedEarly: isFinal ? appliedEarly : [],
      appliedLate: isFinal ? appliedLate : [],
      invalidStatusReasons: isFinal ? invalidStatusReasons : {},
      nextDayQueue,
    };
  };
  return { base: decorate(baseAllocation, false), final: decorate(finalAllocation, true) };
}
