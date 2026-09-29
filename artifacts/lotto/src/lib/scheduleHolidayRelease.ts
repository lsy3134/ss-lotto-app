export type HolidayReleaseMap = Record<string, Record<string, true>>;

const TIMING_STATUSES = new Set(["조출", "후출", "찾근"]);

export function setHolidayRelease(
  releases: HolidayReleaseMap,
  dateKey: string,
  name: string,
  released: boolean,
): HolidayReleaseMap {
  const day = { ...(releases[dateKey] ?? {}) };
  if (released) day[name] = true;
  else delete day[name];

  const next = { ...releases };
  if (Object.keys(day).length > 0) next[dateKey] = day;
  else delete next[dateKey];
  return next;
}

export function mergeLegacyHolidayReleases(
  releases: HolidayReleaseMap,
  statuses: Record<string, Record<string, string | null>>,
): HolidayReleaseMap {
  let next = releases;
  for (const [dateKey, day] of Object.entries(statuses)) {
    for (const [name, status] of Object.entries(day)) {
      if (status === "휴무해제" && !next[dateKey]?.[name]) {
        next = setHolidayRelease(next, dateKey, name, true);
      }
    }
  }
  return next;
}

export function buildBaseSavedDay<T extends string | null>(
  savedDay: Record<string, T>,
  releasedNames: Record<string, true> = {},
): Record<string, T | "휴무해제"> {
  const baseDay: Record<string, T | "휴무해제"> = {};
  for (const [name, status] of Object.entries(savedDay)) {
    if (status != null && TIMING_STATUSES.has(status)) {
      if (releasedNames[name]) baseDay[name] = "휴무해제";
      continue;
    }
    baseDay[name] = status;
  }
  for (const name of Object.keys(releasedNames)) {
    if (!(name in baseDay)) baseDay[name] = "휴무해제";
  }
  return baseDay;
}
