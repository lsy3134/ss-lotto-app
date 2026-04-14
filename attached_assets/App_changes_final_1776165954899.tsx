// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// App.tsx 에 적용할 변경사항 전체
// 총 3단계만 따라하면 끝!
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// [STEP 1] C 상수 교체
// 기존 const C = { ... } 전체를 아래로 교체
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const C = {
  bgPage:      "#eef1f8",
  cardBg:      "#ffffff",
  purple:      "#7c6ef7",
  purpleLight: "#eeebff",
  blue:        "#5b8dee",
  blueLight:   "#eef3ff",
  green:       "#3db882",
  greenLight:  "#e6f9f1",
  yellow:      "#f0b429",
  yellowLight: "#fff8e6",
  red:         "#f76e6e",
  redLight:    "#fff0f0",
  textPrimary:   "#1a2035",
  textSecondary: "#5a6478",
  textMuted:     "#9aa3b5",
  border:      "#e0e5f0",
  // 기존 코드 호환용
  white:       "#ffffff",
  bgDark:      "#7c6ef7",
};


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// [STEP 2] BottomTabBar 컴포넌트 추가
// App() 함수 바로 위에 아래 코드 전체를 붙여넣기
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function BottomTabBar() {
  const [location, navigate] = useLocation();

  const tabs = [
    { path: "/",         icon: "🏠", label: "홈"    },
    { path: "/schedule", icon: "📋", label: "근무표" },
    { path: "/lotto",    icon: "🎰", label: "로또"  },
  ];

  // 스플래시/인트로 화면에서는 탭 바 숨김
  if (location === "/intro" || location === "/splash" || location === "/intro-page") {
    return null;
  }

  return (
    <nav
      style={{
        position: "fixed",
        bottom: 0,
        left: 0,
        right: 0,
        background: "#ffffff",
        borderTop: "1.5px solid #e8ecf4",
        display: "flex",
        zIndex: 100,
        boxShadow: "0 -2px 16px rgba(100,110,180,0.10)",
        paddingBottom: "env(safe-area-inset-bottom)",
      }}
    >
      {tabs.map((tab) => {
        const isActive = location === tab.path;
        return (
          <button
            key={tab.path}
            onClick={() => navigate(tab.path)}
            style={{
              flex: 1,
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              justifyContent: "center",
              gap: "3px",
              padding: "10px 0 8px",
              border: "none",
              background: "transparent",
              cursor: "pointer",
              transition: "all 0.15s",
            }}
          >
            {/* 아이콘 */}
            <span
              style={{
                fontSize: "22px",
                lineHeight: 1,
                filter: isActive ? "none" : "grayscale(1) opacity(0.35)",
                transition: "all 0.2s",
              }}
            >
              {tab.icon}
            </span>

            {/* 레이블 */}
            <span
              style={{
                fontSize: "10px",
                fontWeight: 700,
                color: isActive ? "#7c6ef7" : "#b0b8ca",
                transition: "all 0.2s",
                fontFamily: "'Noto Sans KR', sans-serif",
              }}
            >
              {tab.label}
            </span>

            {/* 활성 인디케이터 점 */}
            {isActive && (
              <div
                style={{
                  width: "4px",
                  height: "4px",
                  borderRadius: "50%",
                  background: "#7c6ef7",
                  marginTop: "1px",
                }}
              />
            )}
          </button>
        );
      })}
    </nav>
  );
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// [STEP 3] App() return 에 <BottomTabBar /> 추가
//
// 기존 return 이 이런 형태일 거예요:
//
//   return (
//     <div style={{ background: C.bgDark, ... }}>
//       <Switch>
//         ...
//       </Switch>
//     </div>
//   );
//
// 아래처럼 두 가지만 수정:
//   1. 최상위 div background 를 C.bgPage 로 변경
//   2. </Switch> 바로 아래에 <BottomTabBar /> 추가
//
//   return (
//     <div style={{ background: C.bgPage, minHeight: "100dvh" }}>
//       <Switch>
//         ...
//       </Switch>
//       <BottomTabBar />    ← 이 줄만 추가!
//     </div>
//   );
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// [STEP 4] 인트로/스플래시 화면 배경색 변경
// App.tsx 에서 스플래시·홈·인트로 인라인 스타일 찾아서
// 아래 값으로 교체 (Ctrl+H 찾아바꾸기)
//
//   "#1a1a2e"  →  "#eef1f8"   (배경)
//   "#16213e"  →  "#ffffff"   (카드)
//   "#4e89ae"  →  "#7c6ef7"   (포인트)
//   "#52de97"  →  "#3db882"   (초록)
//   "#f8b400"  →  "#f0b429"   (골드)
//
// 입장하기 버튼 background 값:
//   기존 → background: "linear-gradient(135deg, #7c6ef7, #5b4de8)"
//   (또는 현재 어두운 색 찾아서 위 값으로 교체)
//
// 로또 생성 버튼 background 값:
//   기존 → background: "linear-gradient(135deg, #f0b429, #d08000)"
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// ※ 로또 경로 확인 필수!
// 현재 App.tsx 에서 로또 라우트 경로 확인 후
// tabs 배열의 path: "/lotto" 를 실제 경로로 수정
//
// 확인 방법: App.tsx 에서 Ctrl+F → "lotto" 검색
// 예) <Route path="/lotto-page"> 면 → path: "/lotto-page"
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
