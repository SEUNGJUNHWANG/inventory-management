"""
재고관리 시스템 - 상수 및 버전 정보
향후 업데이트 시 버전 번호만 변경하면 앱 전체에 반영됩니다.
"""

# ─────────────────────────────────────────
# 앱 정보
# ─────────────────────────────────────────
APP_NAME = "재고관리 시스템"
APP_VERSION = "2.1.0"
APP_TITLE = f"{APP_NAME} v{APP_VERSION}"
APP_AUTHOR = "Tenova"
APP_BUILD_DATE = "2026-04-29"

# ─────────────────────────────────────────
# GitHub 저장소 정보 (자동 업데이트용)
# ※ 본인의 GitHub 계정명과 저장소명으로 변경하세요
# ─────────────────────────────────────────
GITHUB_OWNER = "SEUNGJUNHWANG"   # ← GitHub 계정명으로 변경
GITHUB_REPO  = "inventory-management"   # ← GitHub 저장소명으로 변경

# ─────────────────────────────────────────
# 창 설정
# ─────────────────────────────────────────
WINDOW_MIN_WIDTH = 1024
WINDOW_MIN_HEIGHT = 700
WINDOW_DEFAULT_WIDTH = 1280
WINDOW_DEFAULT_HEIGHT = 800
SIDEBAR_WIDTH = 220

# ─────────────────────────────────────────
# 색상 테마
# ─────────────────────────────────────────
COLORS = {
    "bg": "#f0f2f5",
    "sidebar": "#1e293b",
    "sidebar_text": "#e2e8f0",
    "sidebar_hover": "#334155",
    "sidebar_active": "#3b82f6",
    "primary": "#3b82f6",
    "success": "#22c55e",
    "warning": "#f59e0b",
    "danger": "#ef4444",
    "info": "#6366f1",
    "card_bg": "#ffffff",
    "text": "#1e293b",
    "text_secondary": "#64748b",
    "border": "#e2e8f0",
    "alert_bg": "#fef2f2",
}

# ─────────────────────────────────────────
# 폰트 설정
# ─────────────────────────────────────────
FONT_FAMILY = "맑은 고딕"
FONT_SIZES = {
    "title": 18,
    "subtitle": 16,
    "heading": 14,
    "subheading": 13,
    "body_large": 12,
    "body": 11,
    "small": 10,
    "tiny": 9,
    "stat": 22,
}

# ─────────────────────────────────────────
# 구글 시트 설정
# ─────────────────────────────────────────
SHEET_NAMES = {
    "parts":       "부품마스터",
    "products":    "제품마스터",
    "bom":         "BOM",
    "history":     "입출고이력",
    "customers":   "거래처",
    "sales":       "판매이력",
    "mrp_history": "MRP이력",
}

# ─────────────────────────────────────────
# 메뉴 정의 (사이드바)
# ─────────────────────────────────────────
MENU_ITEMS = [
    ("dashboard", "📊 대시보드"),
    ("parts", "🔩 부품 관리"),
    ("products", "📦 제품 관리"),
    ("bom", "📋 BOM 관리"),
    ("receive", "📥 부품 입고"),
    ("issue", "📤 부품 출고"),
    ("produce", "🏭 제품 생산"),
    ("ship", "🚚 제품 출고(판매)"),
    ("customers", "🏢 거래처 관리"),
    ("mrp", "📋 자재소요계획"),
    ("history", "📜 입출고 이력"),
    ("report", "📊 리포트"),
    ("analytics",    "📈 분석 대시보드"),
    ("settings", "⚙️ 설정"),
    ("users",    "👤 사용자 관리"),
]

# ─────────────────────────────────────────
# 권한 관리
# ─────────────────────────────────────────
# 메뉴 ID → 한글 표시 이름
MENU_LABELS = {
    "dashboard": "📊 대시보드",
    "parts":     "🔩 부품 관리",
    "products":  "📦 제품 관리",
    "bom":       "📋 BOM 관리",
    "receive":   "📥 부품 입고",
    "issue":     "📤 부품 출고",
    "produce":   "🏭 제품 생산",
    "ship":      "🚚 제품 출고(판매)",
    "customers": "🏢 거래처 관리",
    "mrp":       "📋 자재소요계획",
    "history":   "📜 입출고 이력",
    "report":        "📊 리포트",
    "analytics":     "📈 분석 대시보드",
    "settings":      "⚙️ 설정",
    "users":     "👤 사용자 관리",
}

# 권한 부여 가능한 전체 메뉴 ID 순서
ALL_MENU_IDS = [
    "dashboard", "parts", "products", "bom",
    "receive", "issue", "produce", "ship", "customers", "mrp",
    "history", "report", "analytics", "settings", "users",
]

# 기본 관리자 계정 정보
DEFAULT_ADMIN_ID  = "admin"
DEFAULT_ADMIN_PW  = "admin1234"   # 첫 로그인 후 변경 권장
DEFAULT_ADMIN_NAME = "관리자"

# ─────────────────────────────────────────
# 테이블 컬럼 정의
# ─────────────────────────────────────────
PARTS_COLUMNS = ("품번", "업체명", "부품명", "규격", "단위", "단가", "현재재고", "안전재고", "MOQ", "상태", "비고")
PRODUCTS_COLUMNS = ("제품코드", "제품명", "규격", "현재재고", "원가", "판매가", "마진율", "비고")
BOM_COLUMNS = ("제품코드", "제품명", "부품품번", "부품명", "소요량", "단가", "소계", "비고")
HISTORY_COLUMNS = ("No", "일시", "구분", "유형", "품번/제품코드", "품명", "수량", "잔여재고", "관련제품", "비고")
CUSTOMERS_COLUMNS = ("거래처코드", "거래처명", "담당자", "연락처", "이메일", "주소", "비고")
SALES_COLUMNS = ("판매번호", "일시", "제품코드", "제품명", "거래처코드", "거래처명", "수량", "단가", "금액", "잔여재고", "비고")

# ─────────────────────────────────────────
# 쓰기 권한 정의 (메뉴 접근 권한과 별개)
# ─────────────────────────────────────────
# 규칙: 메뉴ID + "_w" 접미사
# 조회 권한 = 메뉴 접근 (MENU_ITEMS)
# 쓰기 권한 = 데이터 추가/수정/삭제/처리 버튼 표시
WRITE_MENU_IDS = [
    "parts_w",     # 부품 추가/수정/삭제
    "products_w",  # 제품 추가/수정/삭제
    "bom_w",       # BOM 추가/수정/삭제
    "receive_w",   # 부품 입고 처리
    "issue_w",     # 부품 출고 처리
    "produce_w",   # 제품 생산 처리
    "ship_w",      # 제품 출고(판매) 처리
    "customers_w", # 거래처 추가/수정/삭제
]

WRITE_MENU_LABELS = {
    "parts_w":     "🔩 부품 입력/수정",
    "products_w":  "📦 제품 입력/수정",
    "bom_w":       "📋 BOM 입력/수정",
    "receive_w":   "📥 부품 입고 처리",
    "issue_w":     "📤 부품 출고 처리",
    "produce_w":   "🏭 제품 생산 처리",
    "ship_w":      "🚚 출고(판매) 처리",
    "customers_w": "🏢 거래처 입력/수정",
}
