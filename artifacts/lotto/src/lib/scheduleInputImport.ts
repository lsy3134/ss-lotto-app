import * as XLSX from "xlsx";

export type ImportStatus = "휴무" | "조출" | "후출" | "찾근";
export type ImportStatusMap = Record<string, Partial<Record<ImportStatus, string[]>>>;

const STATUSES: ImportStatus[] = ["휴무", "조출", "후출", "찾근"];

export const normalizeImportName = (name: string) => name.replace(/\s+/g, "").trim();

const extractKoreanNames = (value: unknown): string[] => {
  if (value == null) return [];
  return String(value)
    .split(/[,，、/\n\r]+/)
    .flatMap(part => {
      const trimmed = part.trim();
      if (/^[가-힣 ]{2,10}$/.test(trimmed)) return [trimmed];
      return trimmed.split(/\s+/).filter(token => /^[가-힣]{2,6}$/.test(token));
    })
    .map(normalizeImportName)
    .filter(name => /^[가-힣]{2,6}$/.test(name) && !STATUSES.includes(name as ImportStatus));
};

const toDateKey = (value: unknown, contextMonth?: number): string | null => {
  if (value == null || value === "") return null;
  if (value instanceof Date) {
    return `${String(value.getMonth() + 1).padStart(2, "0")}.${String(value.getDate()).padStart(2, "0")}`;
  }
  if (typeof value === "number") {
    if (value >= 1 && value <= 31 && contextMonth) {
      return `${String(contextMonth).padStart(2, "0")}.${String(value).padStart(2, "0")}`;
    }
    if (value > 35000 && value < 60000) {
      const date = new Date(Math.round((value - 25569) * 86400 * 1000));
      return `${String(date.getUTCMonth() + 1).padStart(2, "0")}.${String(date.getUTCDate()).padStart(2, "0")}`;
    }
  }
  const text = String(value).trim();
  let match = text.match(/^\d{4}[-.](\d{1,2})[-.](\d{1,2})/);
  if (match) return `${match[1].padStart(2, "0")}.${match[2].padStart(2, "0")}`;
  match = text.match(/^(\d{1,2})[.\-/](\d{1,2})/);
  if (match) return `${match[1].padStart(2, "0")}.${match[2].padStart(2, "0")}`;
  match = text.match(/(\d{1,2})월\s*(\d{1,2})일/);
  if (match) return `${match[1].padStart(2, "0")}.${match[2].padStart(2, "0")}`;
  match = text.match(/^(\d{1,2})(?:일|\s*[（(][일월화수목금토][)）])?$/);
  if (match && contextMonth) {
    const day = Number(match[1]);
    if (day >= 1 && day <= 31) return `${String(contextMonth).padStart(2, "0")}.${String(day).padStart(2, "0")}`;
  }
  return null;
};

const statusFrom = (value: unknown): ImportStatus | null => {
  const text = String(value ?? "").replace(/\s+/g, "");
  return STATUSES.find(status => text === status || text.startsWith(status)) ?? null;
};

const pushNames = (map: ImportStatusMap, dateKey: string, status: ImportStatus, names: string[]) => {
  const day = map[dateKey] ?? (map[dateKey] = {});
  const list = day[status] ?? (day[status] = []);
  for (const name of names) if (!list.includes(name)) list.push(name);
};

/** 1차 조출·후출·찾근 Excel 파서. 표형과 상태별 열 형식을 지원한다. */
export function parseTimingExcelBuffer(buffer: ArrayBuffer, contextMonth?: number): ImportStatusMap {
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const result: ImportStatusMap = {};

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet?.["!ref"]) continue;
    const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, raw: true, defval: "" });
    if (rows.length === 0) continue;

    // 날짜 | 이름 | 상태 형태
    const headerIndex = rows.slice(0, 10).findIndex(row => {
      const headers = row.map(cell => String(cell).replace(/\s+/g, ""));
      return headers.some(h => h.includes("날짜")) && headers.some(h => h.includes("이름")) && headers.some(h => h.includes("상태"));
    });
    if (headerIndex >= 0) {
      const headers = rows[headerIndex].map(cell => String(cell).replace(/\s+/g, ""));
      const dateColumn = headers.findIndex(h => h.includes("날짜"));
      const nameColumn = headers.findIndex(h => h.includes("이름"));
      const statusColumn = headers.findIndex(h => h.includes("상태"));
      for (const row of rows.slice(headerIndex + 1)) {
        const dateKey = toDateKey(row[dateColumn], contextMonth);
        const status = statusFrom(row[statusColumn]);
        if (!dateKey || !status || status === "휴무") continue;
        pushNames(result, dateKey, status, extractKoreanNames(row[nameColumn]));
      }
      continue;
    }

    // 날짜 | 조출 | 후출 | 찾근 형태
    const matrixHeaderIndex = rows.slice(0, 10).findIndex(row => row.some(cell => statusFrom(cell) !== null));
    if (matrixHeaderIndex >= 0) {
      const header = rows[matrixHeaderIndex];
      const statusColumns = header
        .map((cell, index) => ({ index, status: statusFrom(cell) }))
        .filter((entry): entry is { index: number; status: ImportStatus } => entry.status !== null && entry.status !== "휴무");
      if (statusColumns.length > 0) {
        for (const row of rows.slice(matrixHeaderIndex + 1)) {
          const dateKey = row.map(cell => toDateKey(cell, contextMonth)).find(Boolean) ?? null;
          if (!dateKey) continue;
          for (const { index, status } of statusColumns) pushNames(result, dateKey, status, extractKoreanNames(row[index]));
        }
        continue;
      }
    }

    // 날짜·상태가 같은 행에 있는 단순 블록 형태
    for (const row of rows) {
      const dateKey = row.map(cell => toDateKey(cell, contextMonth)).find(Boolean) ?? null;
      const status = row.map(statusFrom).find(Boolean) ?? null;
      if (!dateKey || !status || status === "휴무") continue;
      const names = row.flatMap(extractKoreanNames);
      pushNames(result, dateKey, status, names);
    }
  }

  return result;
}

export function parseOcrStatusText(text: string): ImportStatusMap[string] {
  const result: ImportStatusMap[string] = {};
  let currentStatus: ImportStatus | null = null;
  for (const rawLine of text.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line) continue;
    const detected = statusFrom(line.split(/[\s:：|]/)[0]);
    if (detected) currentStatus = detected;
    if (!currentStatus) continue;
    const withoutStatus = line.replace(new RegExp(`^\\s*${currentStatus}\\s*[:：|-]?\\s*`), "");
    const names = withoutStatus
      .split(/[\s,，、/]+/)
      .map(normalizeImportName)
      .filter(name => /^[가-힣]{2,6}$/.test(name) && !STATUSES.includes(name as ImportStatus));
    const list = result[currentStatus] ?? (result[currentStatus] = []);
    for (const name of names) if (!list.includes(name)) list.push(name);
  }
  return result;
}

export function matchImportedStatuses(
  statuses: ImportStatusMap[string],
  rosterNames: string[],
): { matched: Partial<Record<ImportStatus, string[]>>; needsReview: Partial<Record<ImportStatus, string[]>> } {
  const rosterByNormalized = new Map(rosterNames.map(name => [normalizeImportName(name), name]));
  const matched: Partial<Record<ImportStatus, string[]>> = {};
  const needsReview: Partial<Record<ImportStatus, string[]>> = {};
  for (const status of STATUSES) {
    for (const rawName of statuses[status] ?? []) {
      const rosterName = rosterByNormalized.get(normalizeImportName(rawName));
      const target = rosterName ? matched : needsReview;
      const value = rosterName ?? rawName;
      const list = target[status] ?? (target[status] = []);
      if (!list.includes(value)) list.push(value);
    }
  }
  return { matched, needsReview };
}

export type DateStatusStore = Record<string, Record<string, string | null>>;
export type TimingSourceStore = Record<string, Record<string, ImportStatus>>;

export function mergeHolidayImport(
  current: Record<string, string[]>,
  incoming: Record<string, string[]>,
  selectedDates: string[],
  scope: "all" | "selected",
): Record<string, string[]> {
  const next = { ...current };
  if (scope === "all") {
    const months = new Set(Object.keys(incoming).map(key => key.slice(0, 2)));
    for (const key of Object.keys(next)) if (months.has(key.slice(0, 2))) delete next[key];
    for (const [key, names] of Object.entries(incoming)) next[key] = [...names];
  } else {
    for (const key of selectedDates) next[key] = [...(incoming[key] ?? [])];
  }
  return next;
}

export function clearManualHolidayImports(
  current: DateStatusStore,
  incomingDates: string[],
  selectedDates: string[],
  scope: "all" | "selected",
): DateStatusStore {
  const months = new Set(incomingDates.map(key => key.slice(0, 2)));
  const selected = new Set(selectedDates);
  const next: DateStatusStore = {};
  for (const [dateLabel, statuses] of Object.entries(current)) {
    const shortKey = dateLabel.slice(0, 5);
    const targeted = scope === "all" ? months.has(shortKey.slice(0, 2)) : selected.has(shortKey);
    const cleaned = Object.fromEntries(Object.entries(statuses).filter(([, status]) => !(targeted && status === "휴무")));
    if (Object.keys(cleaned).length > 0) next[dateLabel] = cleaned;
  }
  return next;
}

export function applyTimingImportState(
  currentStatuses: DateStatusStore,
  previousSource: TimingSourceStore,
  incoming: TimingSourceStore,
  targetDateLabels: string[],
  uploadedMonths: string[],
  scope: "all" | "selected",
): { statuses: DateStatusStore; source: TimingSourceStore } {
  const months = new Set(uploadedMonths);
  const targets = scope === "all"
    ? [...new Set([...Object.keys(previousSource).filter(label => months.has(label.slice(0, 2))), ...targetDateLabels])]
    : [...targetDateLabels];
  const statuses: DateStatusStore = Object.fromEntries(
    Object.entries(currentStatuses).map(([dateLabel, day]) => [dateLabel, { ...day }]),
  );
  const source: TimingSourceStore = Object.fromEntries(
    Object.entries(previousSource).map(([dateLabel, day]) => [dateLabel, { ...day }]),
  );

  for (const dateLabel of targets) {
    const day = { ...(statuses[dateLabel] ?? {}) };
    const oldImported = previousSource[dateLabel] ?? {};
    const manuallyChanged = new Set(
      Object.entries(oldImported)
        .filter(([name, oldStatus]) => day[name] !== oldStatus)
        .map(([name]) => name),
    );
    for (const [name, oldStatus] of Object.entries(oldImported)) if (day[name] === oldStatus) delete day[name];

    const acceptedSource: Record<string, ImportStatus> = {};
    for (const [name, status] of Object.entries(incoming[dateLabel] ?? {})) {
      if (manuallyChanged.has(name)) continue;
      day[name] = status;
      acceptedSource[name] = status;
    }
    if (Object.keys(day).length > 0) statuses[dateLabel] = day;
    else delete statuses[dateLabel];
    if (Object.keys(acceptedSource).length > 0) source[dateLabel] = acceptedSource;
    else delete source[dateLabel];
  }
  return { statuses, source };
}

export function applyOcrAssignments(
  current: DateStatusStore,
  targetDateLabel: string,
  assignments: Record<string, ImportStatus>,
): DateStatusStore {
  return {
    ...current,
    [targetDateLabel]: { ...(current[targetDateLabel] ?? {}), ...assignments },
  };
}
