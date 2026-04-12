export type GroupType = "하우스" | "주중" | "주말";

export interface PersonData {
  no: number;
  name: string;
  cart?: number;
  group: GroupType;
  조: 1 | 2 | 3 | 4;
}

// ───────────────────────────────────────────────────
// 캐디 순번표 (총 63명: 하우스 52, 주중 2, 주말 9)
// ───────────────────────────────────────────────────
export const ROSTER: PersonData[] = [
  // ── 1조 하우스 (13명) ──
  { no: 1,  name: "도현지", cart: 57, group: "하우스", 조: 1 },
  { no: 2,  name: "서혜정", cart: 24, group: "하우스", 조: 1 },
  { no: 3,  name: "함은진", cart: 33, group: "하우스", 조: 1 },
  { no: 4,  name: "류지현", cart: 2,  group: "하우스", 조: 1 },
  { no: 5,  name: "강아연", cart: 45, group: "하우스", 조: 1 },
  { no: 6,  name: "김보미", cart: 21, group: "하우스", 조: 1 },
  { no: 7,  name: "정은희", cart: 59, group: "하우스", 조: 1 },
  { no: 8,  name: "김우빈", cart: 11, group: "하우스", 조: 1 },
  { no: 9,  name: "유인매", cart: 46, group: "하우스", 조: 1 },
  { no: 11, name: "이수경", cart: 55, group: "하우스", 조: 1 },
  { no: 13, name: "황은곤", cart: 6,  group: "하우스", 조: 1 },
  { no: 14, name: "이은미", cart: 28, group: "하우스", 조: 1 },
  { no: 15, name: "전푸름", cart: 18, group: "하우스", 조: 1 },

  // ── 2조 하우스 (12명) ──
  { no: 16, name: "김하늘", cart: 52, group: "하우스", 조: 2 },
  { no: 17, name: "황은지", cart: 5,  group: "하우스", 조: 2 },
  { no: 18, name: "성소진", cart: 37, group: "하우스", 조: 2 },
  { no: 19, name: "박수림", cart: 23, group: "하우스", 조: 2 },
  { no: 20, name: "양승원", cart: undefined, group: "하우스", 조: 2 },
  { no: 24, name: "김현정", cart: 8,  group: "하우스", 조: 2 },
  { no: 25, name: "김승희", cart: 34, group: "하우스", 조: 2 },
  { no: 26, name: "유미선", cart: 41, group: "하우스", 조: 2 },
  { no: 27, name: "이수애", cart: 25, group: "하우스", 조: 2 },
  { no: 28, name: "신현진", cart: 9,  group: "하우스", 조: 2 },
  { no: 29, name: "도의지", cart: 44, group: "하우스", 조: 2 },
  { no: 30, name: "천예솔", cart: 20, group: "하우스", 조: 2 },

  // ── 3조 하우스 (13명) ──
  { no: 31, name: "박청민", cart: 30, group: "하우스", 조: 3 },
  { no: 32, name: "진유진", cart: 35, group: "하우스", 조: 3 },
  { no: 33, name: "이주영", cart: 51, group: "하우스", 조: 3 },
  { no: 34, name: "이사야", cart: 47, group: "하우스", 조: 3 },
  { no: 35, name: "김태리", cart: 17, group: "하우스", 조: 3 },
  { no: 38, name: "정희진", cart: 31, group: "하우스", 조: 3 },
  { no: 39, name: "이아영", cart: 40, group: "하우스", 조: 3 },
  { no: 40, name: "임한솔", cart: 14, group: "하우스", 조: 3 },
  { no: 41, name: "안  솔", cart: 43, group: "하우스", 조: 3 },
  { no: 42, name: "이매지", cart: 15, group: "하우스", 조: 3 },
  { no: 43, name: "정  문", cart: 7,  group: "하우스", 조: 3 },
  { no: 44, name: "윤다경", cart: 4,  group: "하우스", 조: 3 },
  { no: 45, name: "이재은", cart: 56, group: "하우스", 조: 3 },

  // ── 4조 하우스 (14명) ──
  { no: 46, name: "이서온", cart: 42, group: "하우스", 조: 4 },
  { no: 47, name: "최보미", cart: 48, group: "하우스", 조: 4 },
  { no: 48, name: "김예진", cart: 53, group: "하우스", 조: 4 },
  { no: 49, name: "백건희", cart: 50, group: "하우스", 조: 4 },
  { no: 50, name: "김경진", cart: 58, group: "하우스", 조: 4 },
  { no: 51, name: "이선희", cart: 54, group: "하우스", 조: 4 },
  { no: 52, name: "류  라", cart: 60, group: "하우스", 조: 4 },
  { no: 53, name: "박민지", cart: 12, group: "하우스", 조: 4 },
  { no: 54, name: "성지영", cart: 3,  group: "하우스", 조: 4 },
  { no: 55, name: "서매진", cart: 19, group: "하우스", 조: 4 },
  { no: 56, name: "김예솔", cart: 13, group: "하우스", 조: 4 },
  { no: 58, name: "박지연", cart: 38, group: "하우스", 조: 4 },
  { no: 59, name: "조현진", cart: 39, group: "하우스", 조: 4 },
  { no: 60, name: "박서윤", cart: 22, group: "하우스", 조: 4 },

  // ── 주중반 (2명, 월~금 근무) ──
  { no: 64, name: "김혜민", cart: 36, group: "주중", 조: 2 },
  { no: 67, name: "백지은", cart: 32, group: "주중", 조: 3 },

  // ── 주말반 (9명, 금~일 근무) ──
  { no: 73, name: "하세진", cart: 10, group: "주말", 조: 1 },
  { no: 77, name: "이보란", cart: undefined, group: "주말", 조: 2 },
  { no: 78, name: "박은선", cart: 16, group: "주말", 조: 2 },
  { no: 79, name: "노  아", cart: 49, group: "주말", 조: 2 },
  { no: 80, name: "신세영", cart: 26, group: "주말", 조: 2 },
  { no: 81, name: "홍다해", cart: 27, group: "주말", 조: 3 },
  { no: 82, name: "박혜진", cart: 29, group: "주말", 조: 3 },
  { no: 85, name: "이지해", cart: undefined, group: "주말", 조: 4 },
  { no: 88, name: "정여진", cart: 1,  group: "주말", 조: 4 },
];

// 이름 → 데이터 맵
export const ROSTER_MAP: Record<string, PersonData> =
  Object.fromEntries(ROSTER.map((p) => [p.name, p]));

// 요일 인덱스(0=월~6=일) + 그룹 → 자동 휴무 여부
export function isAutoOff(group: GroupType, dayIdx: number): boolean {
  if (group === "주중" && (dayIdx === 5 || dayIdx === 6)) return true; // 토,일
  if (group === "주말" && dayIdx <= 3) return true;                    // 월~목
  return false;
}
