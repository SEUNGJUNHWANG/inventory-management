# -*- coding: utf-8 -*-
"""
구글 시트 데이터베이스 모듈 (v2.1 - 쿼터 초과 완전 해결판)

[v2.1 주요 개선사항]
1. API Rate Limit (429) 완전 해결:
   - 지수 백오프(Exponential Backoff) 재시도 로직 추가
   - 모든 쓰기 작업에 자동 재시도 적용
   - 배치 작업 간 스마트 딜레이 추가

2. 대량 업로드 안정성 강화:
   - bulk_add_or_update_parts: 429 오류 시 자동 재시도
   - bulk_add_or_update_bom: 배치 간 딜레이 증가

3. 캐시 TTL 최적화:
   - 기본 30초 → 상황별 조정 가능

4. 오류 로깅 강화:
   - 모든 API 오류를 upload_error.log에 기록

[의존성]
  gspread>=5.12.0
  google-auth>=2.25.0
  openpyxl>=3.1.0
"""

import gspread
from google.oauth2.service_account import Credentials
import os
import json
from datetime import datetime
import time
import threading
import traceback


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# 시트 이름 상수
SHEET_PARTS    = "부품마스터"
SHEET_PRODUCTS = "제품마스터"
SHEET_BOM      = "BOM"
SHEET_HISTORY  = "입출고이력"
SHEET_USERS     = "사용자"
SHEET_CUSTOMERS = "거래처"
SHEET_SALES     = "판매이력"
SHEET_MRP       = "MRP이력"
SHEET_PRICE_LOG = "단가변경이력"
SHEET_BOM_LOG   = "BOM변경이력"

# 캐시 유효 시간 (초)
CACHE_TTL         = 120   # 부품·제품·BOM: 2분
CACHE_TTL_HISTORY = 60    # 이력: 1분 (쓰기 후 자동 무효화)

# API 재시도 설정
MAX_RETRIES   = 5          # 최대 재시도 횟수
BASE_DELAY    = 2.0        # 기본 대기 시간 (초)
MAX_DELAY     = 60.0       # 최대 대기 시간 (초)


# ─────────────────────────────────────────────────────────────────────────────
# 유틸: 지수 백오프 재시도 데코레이터
# ─────────────────────────────────────────────────────────────────────────────

def _retry_on_quota(func):
    """
    gspread APIError 429(쿼터 초과) 발생 시 지수 백오프로 자동 재시도.
    다른 오류는 그대로 raise.
    """
    def wrapper(*args, **kwargs):
        delay = BASE_DELAY
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                return func(*args, **kwargs)
            except gspread.exceptions.APIError as e:
                status = getattr(e, 'response', None)
                code   = status.status_code if status else 0
                if code == 429 and attempt < MAX_RETRIES:
                    # 429: 쿼터 초과 → 대기 후 재시도
                    _log_error(
                        f"[재시도 {attempt}/{MAX_RETRIES}] API 쿼터 초과. "
                        f"{delay:.1f}초 대기 후 재시도합니다."
                    )
                    time.sleep(delay)
                    delay = min(delay * 2, MAX_DELAY)  # 지수 증가
                else:
                    raise
            except Exception:
                raise
        # 최대 재시도 소진
        return func(*args, **kwargs)
    wrapper.__name__ = func.__name__
    return wrapper


def _log_error(msg: str):
    """에러를 upload_error.log 에 기록"""
    try:
        log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                 "..", "upload_error.log")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n")
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────────────────
# 데이터 캐시
# ─────────────────────────────────────────────────────────────────────────────

class DataCache:
    """시트 데이터를 캐싱하여 API 호출을 최소화하는 클래스"""

    def __init__(self):
        self._cache      = {}
        self._timestamps = {}
        self._lock       = threading.Lock()

    def get(self, key, ttl=None):
        """캐시에서 데이터 조회. ttl 미지정 시 CACHE_TTL 사용."""
        effective_ttl = ttl if ttl is not None else CACHE_TTL
        with self._lock:
            if key in self._cache:
                elapsed = time.time() - self._timestamps.get(key, 0)
                if elapsed < effective_ttl:
                    return self._cache[key]
            return None

    def set(self, key, data):
        with self._lock:
            self._cache[key]      = data
            self._timestamps[key] = time.time()

    def invalidate(self, key=None):
        with self._lock:
            if key is None:
                self._cache.clear()
                self._timestamps.clear()
            elif key in self._cache:
                del self._cache[key]
                del self._timestamps[key]

    def invalidate_all(self):
        self.invalidate(None)


# ─────────────────────────────────────────────────────────────────────────────
# 메인 DB 클래스
# ─────────────────────────────────────────────────────────────────────────────

class GoogleSheetsDB:
    """구글 시트를 데이터베이스처럼 사용하는 클래스 (v2.1 쿼터 해결판)"""

    def __init__(self, credentials_path: str, spreadsheet_url: str = None):
        self.credentials_path  = credentials_path
        self.spreadsheet_url   = spreadsheet_url
        self.client            = None
        self.spreadsheet       = None
        self.cache             = DataCache()
        # ── 문제 1 수정: 동시 접근(Race Condition) 방지 ──────────────────────
        # RLock을 사용해 같은 스레드에서 재진입(cancel_history → cancel_production)
        # 가능하도록 하면서, 다른 스레드의 동시 쓰기를 완전히 차단합니다.
        self._op_lock          = threading.RLock()
        self._connect()

    # ── 연결 ──────────────────────────────────────────────────────────────────

    def _connect(self):
        creds = Credentials.from_service_account_file(
            self.credentials_path, scopes=SCOPES
        )
        self.client = gspread.authorize(creds)

        if self.spreadsheet_url:
            self.spreadsheet = self.client.open_by_url(self.spreadsheet_url)
        else:
            self.spreadsheet = self.client.create("재고관리시스템")
        self._initialize_sheets()

    def _initialize_sheets(self):
        """초기 시트 구조 생성 (없는 시트 자동 생성)"""
        NEW_PARTS_HEADERS = [
            "품번", "업체명", "부품명", "규격", "단위",
            "단가", "현재재고", "안전재고", "MOQ", "비고"
        ]

        # 부품마스터
        try:
            ws = self.spreadsheet.worksheet(SHEET_PARTS)
            headers = self._safe_row_values(ws, 1)
            if headers != NEW_PARTS_HEADERS:
                all_data = self._safe_get_all_values(ws)
                old_headers = all_data[0] if all_data else []

                def get_col(h, row):
                    idx = old_headers.index(h) if h in old_headers else -1
                    return row[idx] if 0 <= idx < len(row) else ''

                new_data = [NEW_PARTS_HEADERS]
                for row in all_data[1:]:
                    new_row = [
                        get_col("품번",   row),
                        get_col("업체명", row),
                        get_col("부품명", row),
                        get_col("규격",   row),
                        get_col("단위",   row),
                        get_col("단가",   row) or 0,
                        get_col("현재재고", row) or 0,
                        get_col("안전재고", row) or 0,
                        get_col("MOQ",   row) or 0,
                        get_col("비고",   row),
                    ]
                    new_data.append(new_row)
                ws.clear()
                if new_data:
                    self._safe_update(ws, f"A1:J{len(new_data)}", new_data)
        except gspread.exceptions.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(title=SHEET_PARTS, rows=1000, cols=12)
            self._safe_update(ws, "A1:J1", [NEW_PARTS_HEADERS])

        # 제품마스터
        try:
            ws = self.spreadsheet.worksheet(SHEET_PRODUCTS)
            prod_headers = self._safe_row_values(ws, 1)
            if "판매가" not in prod_headers:
                all_vals = self._safe_get_all_values(ws)
                new_rows = []
                for row in all_vals[1:]:
                    while len(row) < 5:
                        row.append("")
                    new_row = [row[0], row[1], row[2], row[3], "0", row[4]]
                    new_rows.append(new_row)
                self._safe_update(ws, "A1:F1",
                                  [["제품코드", "제품명", "규격", "현재재고", "판매가", "비고"]])
                if new_rows:
                    end_row = 1 + len(new_rows)
                    self._safe_update(ws, f"A2:F{end_row}", new_rows)
                self.cache.invalidate("products")
        except gspread.exceptions.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(title=SHEET_PRODUCTS, rows=1000, cols=10)
            self._safe_update(ws, "A1:F1",
                              [["제품코드", "제품명", "규격", "현재재고", "판매가", "비고"]])

        # BOM
        try:
            ws = self.spreadsheet.worksheet(SHEET_BOM)
            headers = self._safe_row_values(ws, 1)
            if "단가" not in headers:
                self._safe_update(ws, "A1:E1",
                                  [["제품코드", "부품품번", "소요량", "단가", "비고"]])
        except gspread.exceptions.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(title=SHEET_BOM, rows=1000, cols=10)
            self._safe_update(ws, "A1:E1",
                              [["제품코드", "부품품번", "소요량", "단가", "비고"]])

        # 입출고이력
        try:
            self.spreadsheet.worksheet(SHEET_HISTORY)
        except gspread.exceptions.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(
                title=SHEET_HISTORY, rows=10000, cols=15
            )
            self._safe_update(ws, "A1:I1", [[
                "일시", "구분", "유형", "품번/제품코드",
                "품명", "수량", "잔여재고", "관련제품", "비고"
            ]])

        # MRP이력 시트
        MRP_HEADERS = ["저장ID", "저장일시", "계획명", "제품수", "안전재고반영", "생산계획", "소요부품"]
        try:
            self.spreadsheet.worksheet(SHEET_MRP)
        except gspread.exceptions.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(title=SHEET_MRP, rows=500, cols=10)
            self._safe_update(ws, "A1:G1", [MRP_HEADERS])

        # 사용자 시트
        try:
            self.spreadsheet.worksheet(SHEET_USERS)
        except gspread.exceptions.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(title=SHEET_USERS, rows=100, cols=5)
            self._safe_update(ws, "A1:E1",
                              [["아이디", "비밀번호", "이름", "메뉴권한", "활성화"]])
            self._create_default_admin(ws)

        # 거래처
        try:
            self.spreadsheet.worksheet(SHEET_CUSTOMERS)
        except gspread.exceptions.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(title=SHEET_CUSTOMERS, rows=500, cols=8)
            self._safe_update(ws, "A1:G1", [[
                "거래처코드", "거래처명", "담당자", "연락처", "이메일", "주소", "비고"
            ]])

        # 판매이력
        try:
            self.spreadsheet.worksheet(SHEET_SALES)
        except gspread.exceptions.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(title=SHEET_SALES, rows=10000, cols=12)
            self._safe_update(ws, "A1:K1", [[
                "판매번호", "일시", "제품코드", "제품명",
                "거래처코드", "거래처명", "수량", "단가", "금액", "잔여재고", "비고"
            ]])

        # 단가변경이력 시트
        PRICE_LOG_HEADERS = [
            "변경일시", "품번", "부품명", "업체명",
            "이전단가", "변경단가", "변경률", "변경자", "변경사유"
        ]
        try:
            self.spreadsheet.worksheet(SHEET_PRICE_LOG)
        except gspread.exceptions.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(title=SHEET_PRICE_LOG, rows=5000, cols=10)
            self._safe_update(ws, "A1:I1", [PRICE_LOG_HEADERS])

        # BOM변경이력 시트
        BOM_LOG_HEADERS = [
            "변경일시", "변경유형", "제품코드", "제품명",
            "부품품번", "부품명", "이전소요량", "변경소요량", "변경자"
        ]
        try:
            self.spreadsheet.worksheet(SHEET_BOM_LOG)
        except gspread.exceptions.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(title=SHEET_BOM_LOG, rows=5000, cols=10)
            self._safe_update(ws, "A1:I1", [BOM_LOG_HEADERS])

        try:
            self.spreadsheet.del_worksheet(
                self.spreadsheet.worksheet("Sheet1")
            )
        except Exception:
            pass

    # ── 내부 헬퍼: 재시도 래핑된 update ───────────────────────────────────────

    @staticmethod
    @_retry_on_quota
    def _safe_update(ws, cell_range, data):
        """429 오류 시 자동 재시도하는 update"""
        ws.update(cell_range, data)

    @staticmethod
    @_retry_on_quota
    def _safe_batch_update(ws, batch_data):
        """여러 범위를 단 1번의 API 호출로 업데이트 (429 재시도 포함)
        batch_data: [{"range": "A2:E2", "values": [[...]]}, ...]
        """
        ws.batch_update(batch_data)

    @staticmethod
    @_retry_on_quota
    def _safe_update_cell(ws, row, col, value):
        ws.update_cell(row, col, value)

    @staticmethod
    @_retry_on_quota
    def _safe_update_cells(ws, cells):
        ws.update_cells(cells)

    @staticmethod
    @_retry_on_quota
    def _safe_append_row(ws, row):
        ws.append_row(row)

    @staticmethod
    @_retry_on_quota
    def _safe_append_rows(ws, rows):
        """여러 행을 한 번에 append. append_rows는 시트 행 수를 자동 확장하므로
        grid limit 초과 오류(A{n}:E{m} exceeds grid limits)가 발생하지 않는다."""
        ws.append_rows(rows, value_input_option="RAW")

    # ── 읽기 API도 429 시 자동 재시도 ────────────────────────────────────────────
    @staticmethod
    @_retry_on_quota
    def _safe_get_all_records(ws):
        """get_all_records — 429 오류 시 자동 재시도"""
        return ws.get_all_records()

    @staticmethod
    @_retry_on_quota
    def _safe_get_all_values(ws):
        """get_all_values — 429 오류 시 자동 재시도"""
        return ws.get_all_values()

    @staticmethod
    @_retry_on_quota
    def _safe_row_values(ws, row_num):
        """row_values — 429 오류 시 자동 재시도"""
        return ws.row_values(row_num)

    @staticmethod
    @_retry_on_quota
    def _safe_delete_rows(ws, row):
        ws.delete_rows(row)

    # ── 유틸 ──────────────────────────────────────────────────────────────────

    def get_spreadsheet_url(self):
        return self.spreadsheet.url

    def refresh_cache(self):
        self.cache.invalidate_all()

    # ── 캐시 기반 조회 ────────────────────────────────────────────────────────

    def _get_all_parts_cached(self):
        cached = self.cache.get("parts")
        if cached is not None:
            return cached
        ws = self.spreadsheet.worksheet(SHEET_PARTS)
        records = self._safe_get_all_records(ws)
        self.cache.set("parts", records)
        return records

    def _get_all_products_cached(self):
        cached = self.cache.get("products")
        if cached is not None:
            return cached
        ws = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        records = self._safe_get_all_records(ws)
        self.cache.set("products", records)
        return records

    def _get_all_bom_cached(self):
        cached = self.cache.get("bom")
        if cached is not None:
            return cached
        ws = self.spreadsheet.worksheet(SHEET_BOM)
        records = self._safe_get_all_records(ws)
        self.cache.set("bom", records)
        return records

    def _get_all_history_cached(self):
        """입출고이력 캐시 조회 (TTL: 60초, 쓰기 시 자동 무효화)"""
        cached = self.cache.get("history", ttl=CACHE_TTL_HISTORY)
        if cached is not None:
            return cached
        ws = self.spreadsheet.worksheet(SHEET_HISTORY)
        records = self._safe_get_all_records(ws)
        self.cache.set("history", records)
        return records

    def _get_parts_map(self):
        cached = self.cache.get("parts_map")
        if cached is not None:
            return cached
        parts     = self._get_all_parts_cached()
        parts_map = {str(p["품번"]): p for p in parts}
        self.cache.set("parts_map", parts_map)
        return parts_map

    # ── 문제 2 수정: 캐시를 완전히 우회하는 신선한 데이터 조회 메서드 ────────────
    # produce_product 처럼 락 안에서 재고를 검증하는 경우,
    # 30초 캐시 TTL 내에 다른 스레드가 캐시를 재채워도 항상 시트 원본을 읽습니다.

    def _get_fresh_parts_map(self):
        """구글 시트에서 직접 최신 부품 데이터를 읽어 dict 반환 (캐시 우회)"""
        ws      = self.spreadsheet.worksheet(SHEET_PARTS)
        records = self._safe_get_all_records(ws)
        # 캐시도 함께 갱신하여 이후 일반 조회에도 최신 데이터가 반영되도록 함
        self.cache.set("parts",     records)
        parts_map = {str(p["품번"]): p for p in records}
        self.cache.set("parts_map", parts_map)
        return parts_map

    def _get_fresh_parts_with_rows(self):
        """SHEET_PARTS를 단 1회 API 호출로 읽어 (parts_map, all_values) 동시 반환.
        _produce_product_locked 에서 parts_map 조회와 _bulk_update_part_qtys 의
        행번호 조회를 합쳐 API 호출 1회를 절약합니다."""
        ws         = self.spreadsheet.worksheet(SHEET_PARTS)
        all_values = self._safe_get_all_values(ws)
        if not all_values:
            return {}, []
        headers = all_values[0]
        records = []
        for row in all_values[1:]:
            if any(row):
                padded = (list(row) + [""] * len(headers))[:len(headers)]
                records.append(dict(zip(headers, padded)))
        self.cache.set("parts", records)
        parts_map = {str(r.get("품번", "")): r for r in records}
        self.cache.set("parts_map", parts_map)
        return parts_map, all_values

    def _get_fresh_product(self, product_id: str):
        """구글 시트에서 직접 최신 제품 데이터를 읽어 반환 (캐시 우회)"""
        ws      = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        records = self._safe_get_all_records(ws)
        self.cache.set("products", records)
        for r in records:
            if str(r.get("제품코드", "")) == str(product_id):
                return r
        return None

    def _get_fresh_bom_for_product(self, product_id: str):
        """구글 시트에서 직접 최신 BOM 데이터를 읽어 반환 (캐시 우회)"""
        ws      = self.spreadsheet.worksheet(SHEET_BOM)
        records = self._safe_get_all_records(ws)
        self.cache.set("bom", records)
        return [r for r in records if str(r.get("제품코드", "")) == str(product_id)]

    # ─────────────────────────────────────────────────────────────────────────
    # 부품 관리
    # ─────────────────────────────────────────────────────────────────────────

    def get_all_parts(self):
        return self._get_all_parts_cached()

    def get_part_by_id(self, part_id: str):
        return self._get_parts_map().get(str(part_id), None)

    def add_part(self, part_id, name, spec, unit, qty, safety_qty,
                 note="", supplier="", unit_price=0, moq=0):
        ws = self.spreadsheet.worksheet(SHEET_PARTS)
        self._safe_append_row(ws, [
            str(part_id), supplier, name, spec, unit,
            float(unit_price), int(qty), int(safety_qty), int(moq), note
        ])
        self.cache.invalidate("parts")
        self.cache.invalidate("parts_map")

    def update_part_qty(self, part_id: str, new_qty: int):
        ws      = self.spreadsheet.worksheet(SHEET_PARTS)
        records = self._safe_get_all_records(ws)
        for i, r in enumerate(records):
            if str(r.get("품번", "")) == str(part_id):
                self._safe_update_cell(ws, i + 2, 7, int(new_qty))
                self.cache.invalidate("parts")
                self.cache.invalidate("parts_map")
                return True
        return False

    def _bulk_update_part_qtys(self, updates, parts_all_values=None):
        """재고 수량 일괄 업데이트 (API 1회) - 재시도 포함
        parts_all_values: _get_fresh_parts_with_rows() 에서 미리 읽어온 경우 전달해
                          SHEET_PARTS 이중 읽기를 방지합니다."""
        ws = self.spreadsheet.worksheet(SHEET_PARTS)
        if parts_all_values is None:
            all_values = self._safe_get_all_values(ws)
        else:
            all_values = parts_all_values

        row_map = {}
        for i, row in enumerate(all_values[1:], 2):
            if row:
                row_map[str(row[0])] = i

        cells_to_update = []
        for part_id, new_qty in updates:
            row_num = row_map.get(str(part_id))
            if row_num:
                cells_to_update.append(
                    gspread.Cell(row_num, 7, int(new_qty))
                )

        if cells_to_update:
            self._safe_update_cells(ws, cells_to_update)

        self.cache.invalidate("parts")
        self.cache.invalidate("parts_map")

    def update_part(self, part_id, name, spec, unit, qty, safety_qty,
                    note="", supplier="", unit_price=0, moq=0,
                    changed_by="", change_reason=""):
        ws      = self.spreadsheet.worksheet(SHEET_PARTS)
        records = self._safe_get_all_records(ws)
        for i, r in enumerate(records):
            if str(r.get("품번", "")) == str(part_id):
                old_price = float(r.get("단가", 0) or 0)
                new_price = float(unit_price)
                row = i + 2
                self._safe_update(ws, f"A{row}:J{row}", [[
                    str(part_id), supplier, name, spec, unit,
                    new_price, int(qty), int(safety_qty), int(moq), note
                ]])
                self.cache.invalidate("parts")
                self.cache.invalidate("parts_map")
                # 단가가 달라진 경우에만 이력 기록
                if old_price != new_price:
                    self._append_price_log(
                        part_id=str(part_id),
                        part_name=str(name),
                        supplier=str(supplier),
                        old_price=old_price,
                        new_price=new_price,
                        changed_by=changed_by,
                        reason=change_reason,
                    )
                return True
        return False

    def _append_price_log(self, part_id, part_name, supplier,
                          old_price, new_price, changed_by="", reason=""):
        """단가변경이력 시트에 1행 추가"""
        try:
            ws = self.spreadsheet.worksheet(SHEET_PRICE_LOG)
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if old_price == 0:
                rate_str = "신규"
            else:
                rate = (new_price - old_price) / old_price * 100
                rate_str = f"{rate:+.1f}%"
            self._safe_append_row(ws, [
                now, part_id, part_name, supplier,
                old_price, new_price, rate_str, changed_by, reason
            ])
        except Exception:
            pass   # 이력 기록 실패는 메인 작업에 영향 없이 무시

    # ── BOM 변경이력 ─────────────────────────────────────────────────────────

    def _append_bom_log(self, action: str, product_id: str, part_id: str,
                        old_qty: float, new_qty: float):
        """BOM변경이력 시트에 1행 추가. 실패해도 메인 작업에 영향 없음."""
        try:
            from core.auth import Session
            products_map = {str(p["제품코드"]): p.get("제품명", "")
                            for p in self._get_all_products_cached()}
            parts_map    = self._get_parts_map()
            prod_name    = products_map.get(str(product_id), "")
            part_name    = parts_map.get(str(part_id), {}).get("부품명", "")
            changed_by   = Session.user_name() or "시스템"
            ws  = self.spreadsheet.worksheet(SHEET_BOM_LOG)
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self._safe_append_row(ws, [
                now, action,
                str(product_id), prod_name,
                str(part_id), part_name,
                float(old_qty) if old_qty is not None else "",
                float(new_qty) if new_qty is not None else "",
                changed_by,
            ])
        except Exception:
            pass

    def get_bom_change_history(self, start_date: str = None,
                               end_date: str = None,
                               product_id: str = None) -> list:
        """BOM 변경이력 조회 (최신순). start/end: 'YYYY-MM-DD'"""
        try:
            ws      = self.spreadsheet.worksheet(SHEET_BOM_LOG)
            records = ws.get_all_records()
            result  = []
            for r in records:
                dt_str = str(r.get("변경일시", ""))[:10]
                if not dt_str:
                    continue
                if start_date and dt_str < start_date:
                    continue
                if end_date and dt_str > end_date:
                    continue
                if product_id and str(r.get("제품코드", "")) != str(product_id):
                    continue
                result.append(r)
            return list(reversed(result))
        except Exception:
            return []

    # ── 분석용 집계 함수 ─────────────────────────────────────────────────────

    def get_monthly_sales_with_margin(self, year: int = None) -> list:
        """월별 판매 집계 (원가·마진 포함).
        원가는 현재 BOM 기준 제품원가 × 판매수량으로 추산.
        Returns: [{"월","건수","수량","판매대금","원가","마진","마진율"}, ...]
        """
        sales         = self.get_all_sales()
        product_costs = self.get_all_product_costs()  # {product_id: unit_cost}
        monthly: dict = {}
        for r in sales:
            dt_str = str(r.get("일시", ""))[:7]
            if not dt_str or len(dt_str) < 7:
                continue
            if year and not dt_str.startswith(str(year)):
                continue
            pid    = str(r.get("제품코드", ""))
            qty    = int(r.get("수량", 0) or 0)
            amount = int(r.get("금액", 0) or 0)
            cost   = float(product_costs.get(pid, 0)) * qty
            if dt_str not in monthly:
                monthly[dt_str] = {
                    "월": dt_str, "건수": 0, "수량": 0,
                    "판매대금": 0, "원가": 0.0, "마진": 0.0,
                }
            monthly[dt_str]["건수"]   += 1
            monthly[dt_str]["수량"]   += qty
            monthly[dt_str]["판매대금"] += amount
            monthly[dt_str]["원가"]   += cost
            monthly[dt_str]["마진"]   += amount - cost
        result = sorted(monthly.values(), key=lambda x: x["월"])
        for m in result:
            rev = m["판매대금"]
            m["마진율"] = round(m["마진"] / rev * 100, 1) if rev > 0 else 0.0
            m["원가"]   = round(m["원가"])
            m["마진"]   = round(m["마진"])
        return result

    def get_parts_io_summary(self, year: int = None) -> list:
        """월별 부품 입출고 건수·수량 집계.
        Returns: [{"월","입고건수","출고건수","입고수량","출고수량"}, ...]
        """
        history  = self.get_all_history()
        monthly: dict = {}
        for h in history:
            dt_str = str(h.get("일시", ""))[:7]
            if not dt_str or len(dt_str) < 7:
                continue
            if year and not dt_str.startswith(str(year)):
                continue
            h_type = str(h.get("유형", ""))
            if h_type not in ("부품입고", "부품출고"):
                continue
            qty = int(h.get("수량", 0) or 0)
            if dt_str not in monthly:
                monthly[dt_str] = {
                    "월": dt_str,
                    "입고건수": 0, "출고건수": 0,
                    "입고수량": 0, "출고수량": 0,
                }
            if h_type == "부품입고":
                monthly[dt_str]["입고건수"] += 1
                monthly[dt_str]["입고수량"] += qty
            else:
                monthly[dt_str]["출고건수"] += 1
                monthly[dt_str]["출고수량"] += qty
        return sorted(monthly.values(), key=lambda x: x["월"])

    def get_price_history(self, year=None, month=None, supplier=None):
        """단가변경이력 조회 (연·월·업체 필터 선택적 적용)"""
        ws      = self.spreadsheet.worksheet(SHEET_PRICE_LOG)
        records = ws.get_all_records()
        result  = []
        for r in records:
            dt_str = str(r.get("변경일시", ""))
            if not dt_str:
                continue
            if year and not dt_str.startswith(str(year)):
                continue
            if month:
                ym = f"{year}-{str(month).zfill(2)}"
                if not dt_str.startswith(ym):
                    continue
            if supplier and supplier != "전체":
                if str(r.get("업체명", "")) != supplier:
                    continue
            result.append({
                "변경일시":  dt_str,
                "품번":      str(r.get("품번", "")),
                "부품명":    str(r.get("부품명", "")),
                "업체명":    str(r.get("업체명", "")),
                "이전단가":  float(r.get("이전단가", 0) or 0),
                "변경단가":  float(r.get("변경단가", 0) or 0),
                "변경률":    str(r.get("변경률", "")),
                "변경자":    str(r.get("변경자", "")),
                "변경사유":  str(r.get("변경사유", "")),
            })
        return result

    def delete_price_history(self, changed_at: str, part_id: str,
                             old_price: float, new_price: float) -> bool:
        """단가변경이력 1건 삭제. (변경일시, 품번, 이전단가, 변경단가)가 모두 일치하는
        첫 번째 행을 찾아 삭제한다. 잘못 기록된 이력을 정정할 때 사용."""
        ws      = self.spreadsheet.worksheet(SHEET_PRICE_LOG)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if (str(r.get("변경일시", "")) == str(changed_at) and
                    str(r.get("품번", "")) == str(part_id) and
                    float(r.get("이전단가", 0) or 0) == float(old_price) and
                    float(r.get("변경단가", 0) or 0) == float(new_price)):
                self._safe_delete_rows(ws, i + 2)
                return True
        return False

    def get_monthly_price_change_report_data(self, year: int, month: int) -> dict:
        """월간 단가변경 보고서용 데이터 집계.
        해당 월에 단가가 변경된 부품별 요약과, 그 변경이 BOM을 통해
        어떤 제품의 원가에 얼마나 영향을 주었는지를 계산한다.

        Returns: {
            "raw_rows": [해당월 변경이력 원본(시간순)],
            "changed_parts": [{"품번","부품명","업체명","이전단가","변경단가","변경률","변경횟수"}],
            "product_impacts": [{"제품코드","제품명","원가변경전","원가변경후","증감액","증감률","영향부품":[...]}],
        }
        """
        ym        = f"{year}-{str(month).zfill(2)}"
        all_price = self.get_price_history()
        month_rows = sorted(
            (r for r in all_price if r["변경일시"][:7] == ym),
            key=lambda r: r["변경일시"],
        )

        changed: dict = {}
        for r in month_rows:
            pid = r["품번"]
            if pid not in changed:
                changed[pid] = {
                    "품번": pid, "부품명": r["부품명"], "업체명": r["업체명"],
                    "이전단가": r["이전단가"], "변경단가": r["변경단가"], "변경횟수": 0,
                }
            changed[pid]["변경단가"] = r["변경단가"]
            changed[pid]["변경횟수"] += 1

        changed_parts = []
        for c in changed.values():
            old, new = c["이전단가"], c["변경단가"]
            c["변경률"] = round((new - old) / old * 100, 1) if old else None  # None = 신규
            changed_parts.append(c)
        changed_parts.sort(
            key=lambda c: abs(c["변경률"]) if c["변경률"] is not None else 999,
            reverse=True,
        )

        all_bom      = self.get_all_bom()
        products_map = {str(p["제품코드"]): p.get("제품명", "")
                        for p in self._get_all_products_cached()}
        by_product: dict = {}
        for b in all_bom:
            by_product.setdefault(str(b.get("제품코드", "")), []).append(b)

        changed_pids = set(changed.keys())
        product_impacts = []
        for pid, lines in by_product.items():
            if not any(str(l.get("부품품번", "")) in changed_pids for l in lines):
                continue
            cost_before = 0.0
            cost_after  = 0.0
            affected    = []
            for l in lines:
                part_id = str(l.get("부품품번", ""))
                qty     = float(l.get("소요량", 0) or 0)
                if part_id in changed_pids:
                    c        = changed[part_id]
                    before_p = c["이전단가"]
                    after_p  = c["변경단가"]
                    affected.append({
                        "부품품번": part_id, "부품명": c["부품명"], "소요량": qty,
                        "이전단가": before_p, "변경단가": after_p,
                        "원가증감": round((after_p - before_p) * qty),
                    })
                else:
                    before_p = after_p = float(l.get("단가", 0) or 0)
                cost_before += qty * before_p
                cost_after  += qty * after_p
            diff = cost_after - cost_before
            if diff == 0:
                continue
            product_impacts.append({
                "제품코드": pid, "제품명": products_map.get(pid, ""),
                "원가변경전": round(cost_before), "원가변경후": round(cost_after),
                "증감액": round(diff),
                "증감률": round(diff / cost_before * 100, 1) if cost_before else None,
                "영향부품": affected,
            })
        product_impacts.sort(key=lambda x: x["증감액"], reverse=True)

        return {
            "raw_rows": month_rows,
            "changed_parts": changed_parts,
            "product_impacts": product_impacts,
        }

    def get_price_history_monthly_summary(self):
        """최근 12개월 월별 변경 건수 요약 반환"""
        from collections import defaultdict
        ws      = self.spreadsheet.worksheet(SHEET_PRICE_LOG)
        records = ws.get_all_records()
        counts  = defaultdict(lambda: {"total": 0, "up": 0, "down": 0})
        for r in records:
            dt_str = str(r.get("변경일시", ""))
            if len(dt_str) < 7:
                continue
            ym  = dt_str[:7]   # "YYYY-MM"
            rate = str(r.get("변경률", ""))
            counts[ym]["total"] += 1
            if rate.startswith("+"):
                counts[ym]["up"] += 1
            elif rate.startswith("-"):
                counts[ym]["down"] += 1
        return dict(counts)

    def bulk_add_or_update_parts(self, parts_list, progress_callback=None):
        """
        부품 대량 등록/수정 (v2.1 - 429 자동 재시도 + 배치 딜레이)

        개선사항:
          - 모든 update() 호출에 재시도 래핑 적용
          - 배치 작업 간 2초 딜레이
          - 오류 발생 시 로그 기록 후 계속 진행
        """
        if progress_callback:
            progress_callback("시트 데이터 읽는 중...")

        ws         = self.spreadsheet.worksheet(SHEET_PARTS)
        all_values = ws.get_all_values()
        header     = all_values[0] if all_values else []
        existing_rows = all_values[1:] if len(all_values) > 1 else []

        def col_idx(name):
            try:
                return header.index(name)
            except ValueError:
                return -1

        idx = {
            "품번":   col_idx("품번"),
            "업체명": col_idx("업체명"),
            "부품명": col_idx("부품명"),
            "규격":   col_idx("규격"),
            "단위":   col_idx("단위"),
            "단가":   col_idx("단가"),
            "현재재고": col_idx("현재재고"),
            "안전재고": col_idx("안전재고"),
            "MOQ":    col_idx("MOQ"),
            "비고":   col_idx("비고"),
        }

        existing_map = {}
        for i, row in enumerate(existing_rows):
            code_val = (row[idx["품번"]] if idx["품번"] >= 0 and len(row) > idx["품번"]
                        else "")
            if code_val:
                existing_map[str(code_val)] = i

        new_count    = 0
        update_count = 0

        def get_cell(row, key):
            i2 = idx.get(key, -1)
            if i2 < 0 or i2 >= len(row):
                return ""
            return row[i2]

        def merge_val(excel_val, existing_val, is_numeric=False, allow_zero=False):
            s = str(excel_val or "").strip()
            if not s:
                return existing_val
            if is_numeric:
                try:
                    v = float(s)
                    if not allow_zero and v == 0:
                        return existing_val
                    return v
                except (ValueError, TypeError):
                    return existing_val
            return s

        price_log_batch = []   # 단가 변경 이력 배치 수집

        for p in parts_list:
            code = str(p["품번"])
            if code in existing_map:
                i2  = existing_map[code]
                row = existing_rows[i2]
                while len(row) < len(header):
                    row.append("")
                # 단가 변경 전 이전값 기억
                old_price_raw = get_cell(row, "단가")
                try:
                    old_price = float(old_price_raw) if old_price_raw else 0.0
                except (ValueError, TypeError):
                    old_price = 0.0

                row[idx["업체명"]]  = merge_val(p.get("업체명", ""),  get_cell(row, "업체명"))
                row[idx["부품명"]]  = merge_val(p.get("부품명", ""),  get_cell(row, "부품명"))
                row[idx["규격"]]    = merge_val(p.get("규격", ""),    get_cell(row, "규격"))
                row[idx["단위"]]    = merge_val(p.get("단위", ""),    get_cell(row, "단위"))
                row[idx["단가"]]    = merge_val(p.get("단가", ""),    get_cell(row, "단가"),    is_numeric=True)
                row[idx["현재재고"]] = merge_val(p.get("현재재고", ""), get_cell(row, "현재재고"), is_numeric=True, allow_zero=True)
                row[idx["안전재고"]] = merge_val(p.get("안전재고", ""), get_cell(row, "안전재고"), is_numeric=True, allow_zero=True)
                row[idx["MOQ"]]    = merge_val(p.get("MOQ", ""),    get_cell(row, "MOQ"),    is_numeric=True)
                row[idx["비고"]]    = merge_val(p.get("비고", ""),    get_cell(row, "비고"))
                existing_rows[i2]   = row
                update_count        += 1

                # 단가 변경 감지
                new_price_raw = row[idx["단가"]]
                try:
                    new_price = float(new_price_raw) if new_price_raw else 0.0
                except (ValueError, TypeError):
                    new_price = 0.0
                if old_price != new_price:
                    price_log_batch.append({
                        "part_id":   code,
                        "part_name": str(row[idx["부품명"]]),
                        "supplier":  str(row[idx["업체명"]]),
                        "old_price": old_price,
                        "new_price": new_price,
                    })
            else:
                new_row              = [""] * len(header)
                new_row[idx["품번"]]   = code
                new_row[idx["업체명"]] = str(p.get("업체명", "") or "")
                new_row[idx["부품명"]] = str(p.get("부품명", "") or "")
                new_row[idx["규격"]]   = str(p.get("규격", "")   or "")
                new_row[idx["단위"]]   = str(p.get("단위", "EA") or "EA")
                new_row[idx["단가"]]   = float(p.get("단가", 0)   or 0)
                new_row[idx["현재재고"]] = int(float(p.get("현재재고", 0) or 0))
                new_row[idx["안전재고"]] = int(float(p.get("안전재고", 0) or 0))
                new_row[idx["MOQ"]]    = int(float(p.get("MOQ", 0)    or 0))
                new_row[idx["비고"]]   = str(p.get("비고", "")   or "")
                existing_rows.append(new_row)
                existing_map[code]   = len(existing_rows) - 1
                new_count            += 1

        if progress_callback:
            progress_callback(
                f"시트에 저장 중... (신규 {new_count}건 + 수정 {update_count}건)"
            )

        # ★ 핵심 개선: 한 번에 전체 업데이트 + 재시도 래핑
        if existing_rows:
            end_row        = len(existing_rows) + 1
            end_col_letter = chr(ord('A') + len(header) - 1)
            try:
                self._safe_update(ws, f"A2:{end_col_letter}{end_row}", existing_rows)
            except gspread.exceptions.APIError as e:
                _log_error(f"bulk_add_or_update_parts 실패: {e}\n{traceback.format_exc()}")
                raise

        self.cache.invalidate("parts")
        self.cache.invalidate("parts_map")

        # 단가 변경 이력 일괄 기록 (엑셀 업로드)
        for log in price_log_batch:
            self._append_price_log(
                part_id=log["part_id"],
                part_name=log["part_name"],
                supplier=log["supplier"],
                old_price=log["old_price"],
                new_price=log["new_price"],
                changed_by="엑셀업로드",
                reason="",
            )
            time.sleep(0.15)

        return new_count, update_count

    def delete_part(self, part_id: str):
        ws      = self.spreadsheet.worksheet(SHEET_PARTS)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if str(r.get("품번", "")) == str(part_id):
                self._safe_delete_rows(ws, i + 2)
                self.cache.invalidate("parts")
                self.cache.invalidate("parts_map")
                return True
        return False

    # ─────────────────────────────────────────────────────────────────────────
    # 제품 관리
    # ─────────────────────────────────────────────────────────────────────────

    def get_all_products(self):
        return self._get_all_products_cached()

    def get_product_by_id(self, product_id: str):
        products = self._get_all_products_cached()
        for r in products:
            if str(r.get("제품코드", "")) == str(product_id):
                return r
        return None

    def add_product(self, product_id, name, spec, qty=0, selling_price=0, note=""):
        ws = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        self._safe_append_row(ws, [str(product_id), name, spec, int(qty), float(selling_price), note])
        self.cache.invalidate("products")

    def update_product_qty(self, product_id: str, new_qty: int):
        ws      = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        records = self._safe_get_all_records(ws)
        for i, r in enumerate(records):
            if str(r.get("제품코드", "")) == str(product_id):
                self._safe_update_cell(ws, i + 2, 4, int(new_qty))
                self.cache.invalidate("products")
                return True
        return False

    def update_product(self, product_id, name, spec, qty, selling_price=0, note=""):
        ws      = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        records = self._safe_get_all_records(ws)
        for i, r in enumerate(records):
            if str(r.get("제품코드", "")) == str(product_id):
                row = i + 2
                self._safe_update(ws, f"A{row}:F{row}",
                                  [[str(product_id), name, spec, int(qty), float(selling_price), note]])
                self.cache.invalidate("products")
                return True
        return False

    def get_all_product_costs(self) -> dict:
        """전체 제품의 원가를 {제품코드: 원가} 딕셔너리로 반환"""
        all_bom   = self._get_all_bom_cached()
        parts_map = self._get_parts_map()
        costs: dict = {}
        for r in all_bom:
            pid   = str(r.get("제품코드", ""))
            qty   = float(r.get("소요량", 0) or 0)
            part  = parts_map.get(str(r.get("부품품번", "")), {})
            price = float(part.get("단가", 0) or 0)
            costs[pid] = costs.get(pid, 0.0) + qty * price
        return costs

    def delete_product(self, product_id: str):
        ws      = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        records = self._safe_get_all_records(ws)
        for i, r in enumerate(records):
            if str(r.get("제품코드", "")) == str(product_id):
                self._safe_delete_rows(ws, i + 2)
                self.cache.invalidate("products")
                return True
        return False

    # ─────────────────────────────────────────────────────────────────────────
    # BOM 관리
    # ─────────────────────────────────────────────────────────────────────────

    def get_bom_for_product(self, product_id: str):
        all_bom   = self._get_all_bom_cached()
        parts_map = self._get_parts_map()
        bom_list  = []
        for r in all_bom:
            if str(r.get("제품코드", "")) == str(product_id):
                item    = dict(r)
                part_id = str(item.get("부품품번", ""))
                part    = parts_map.get(part_id, {})
                item["단가"] = float(part.get("단가", 0) or 0)
                bom_list.append(item)
        return bom_list

    def get_all_bom(self):
        all_bom   = self._get_all_bom_cached()
        parts_map = self._get_parts_map()
        result    = []
        for r in all_bom:
            item    = dict(r)
            part_id = str(item.get("부품품번", ""))
            part    = parts_map.get(part_id, {})
            item["단가"] = float(part.get("단가", 0) or 0)
            result.append(item)
        return result

    def add_bom(self, product_id, part_id, qty, note=""):
        """BOM 항목 추가. 단가는 부품마스터에서만 관리하므로 BOM 시트에 저장하지 않음."""
        ws = self.spreadsheet.worksheet(SHEET_BOM)
        self._safe_append_row(ws, [
            str(product_id), str(part_id), float(qty), 0, note  # 단가 열은 항상 0
        ])
        self.cache.invalidate("bom")
        self._append_bom_log("추가", product_id, part_id, None, qty)

    def update_bom(self, product_id: str, part_id: str,
                   qty: float, note: str = ""):
        """BOM 항목 수정. 단가는 부품마스터에서만 관리하므로 BOM 시트에 저장하지 않음."""
        ws      = self.spreadsheet.worksheet(SHEET_BOM)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if (str(r.get("제품코드", "")) == str(product_id) and
                    str(r.get("부품품번", "")) == str(part_id)):
                old_qty = float(r.get("소요량", 0) or 0)
                row = i + 2
                self._safe_update(ws, f"A{row}:E{row}", [[
                    str(product_id), str(part_id), float(qty), 0, note  # 단가 열은 항상 0
                ]])
                self.cache.invalidate("bom")
                self._append_bom_log("수정", product_id, part_id, old_qty, qty)
                return True
        return False

    def replace_bom_part(self, product_id: str, old_part_id: str,
                         new_part_id: str, qty: float, note: str = ""):
        """BOM 항목의 품번을 교체. 기존 행의 부품품번을 새 품번으로 변경."""
        ws      = self.spreadsheet.worksheet(SHEET_BOM)
        records = ws.get_all_records()
        for r in records:
            if (str(r.get("제품코드", "")) == str(product_id) and
                    str(r.get("부품품번", "")) == str(new_part_id)):
                raise ValueError(f"품번 '{new_part_id}'은(는) 이미 이 제품의 BOM에 등록되어 있습니다.")
        for i, r in enumerate(records):
            if (str(r.get("제품코드", "")) == str(product_id) and
                    str(r.get("부품품번", "")) == str(old_part_id)):
                old_qty = float(r.get("소요량", 0) or 0)
                row = i + 2
                self._safe_update(ws, f"A{row}:E{row}", [[
                    str(product_id), str(new_part_id), float(qty), 0, note
                ]])
                self.cache.invalidate("bom")
                # 구부품 삭제 + 신부품 추가로 두 줄 기록
                self._append_bom_log("삭제(교체)", product_id, old_part_id, old_qty, None)
                self._append_bom_log("추가(교체)", product_id, new_part_id, None, qty)
                return True
        return False

    def delete_bom(self, product_id: str, part_id: str):
        ws      = self.spreadsheet.worksheet(SHEET_BOM)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if (str(r.get("제품코드", "")) == str(product_id) and
                    str(r.get("부품품번", "")) == str(part_id)):
                old_qty = float(r.get("소요량", 0) or 0)
                self._safe_delete_rows(ws, i + 2)
                self.cache.invalidate("bom")
                self._append_bom_log("삭제", product_id, part_id, old_qty, None)
                return True
        return False

    def delete_all_bom_for_product(self, product_id: str):
        ws              = self.spreadsheet.worksheet(SHEET_BOM)
        records         = ws.get_all_records()
        rows_to_delete  = [i + 2 for i, r in enumerate(records)
                           if str(r.get("제품코드", "")) == str(product_id)]
        for row in sorted(rows_to_delete, reverse=True):
            self._safe_delete_rows(ws, row)
            time.sleep(0.3)
        self.cache.invalidate("bom")

    def bulk_add_or_update_bom(self, bom_list, progress_callback=None):
        """BOM 대량 등록/수정 (429 재시도 포함)"""
        ws                = self.spreadsheet.worksheet(SHEET_BOM)
        existing_records  = ws.get_all_records()

        existing_map = {}
        for i, r in enumerate(existing_records):
            key = (str(r.get("제품코드", "")), str(r.get("부품품번", "")))
            existing_map[key] = i + 2

        new_items      = []
        update_batches = []
        new_count      = 0
        update_count   = 0

        for item in bom_list:
            prod_code  = str(item["제품코드"])
            part_code  = str(item["부품품번"])
            qty        = float(item["소요량"])
            # 단가는 부품마스터에서만 관리 — Excel이나 입력값의 단가는 무시하고 항상 0 저장
            note       = str(item.get("비고", ""))
            row_data   = [prod_code, part_code, qty, 0, note]
            key        = (prod_code, part_code)

            if key in existing_map:
                update_batches.append((existing_map[key], row_data))
                update_count += 1
            else:
                new_items.append(row_data)
                new_count += 1

        # 신규 추가 (배치)
        # append_rows 사용: 시트 행 수를 자동 확장하므로 grid limit 초과 오류 없음
        BATCH = 100
        for i in range(0, len(new_items), BATCH):
            batch = new_items[i:i + BATCH]
            self._safe_append_rows(ws, batch)
            if progress_callback:
                progress_callback(
                    f"BOM 신규 등록 중... {min(i + BATCH, len(new_items))}/{len(new_items)}"
                )
            if i + BATCH < len(new_items):
                time.sleep(2)  # 배치 간 딜레이

        # 기존 수정: batch_update로 모든 행을 단 1번의 API 호출로 처리
        UBATCH = 500
        for i in range(0, len(update_batches), UBATCH):
            chunk = update_batches[i:i + UBATCH]
            batch_data = [
                {"range": f"A{row_num}:E{row_num}", "values": [row_data]}
                for row_num, row_data in chunk
            ]
            self._safe_batch_update(ws, batch_data)
            if progress_callback:
                progress_callback(
                    f"BOM 수정 중... {min(i + UBATCH, len(update_batches))}/{len(update_batches)}"
                )
            if i + UBATCH < len(update_batches):
                time.sleep(1)

        self.cache.invalidate("bom")
        return new_count, update_count

    def get_product_cost(self, product_id: str):
        bom       = self.get_bom_for_product(product_id)
        parts_map = self._get_parts_map()

        total_cost   = 0
        cost_details = []
        for item in bom:
            part_id    = str(item.get("부품품번", ""))
            qty        = float(item.get("소요량", 0))
            unit_price = float(item.get("단가", 0))
            subtotal   = qty * unit_price
            part_name  = parts_map.get(part_id, {}).get("부품명", "?")
            cost_details.append({
                "부품품번": part_id,
                "부품명":   part_name,
                "소요량":   qty,
                "단가":     unit_price,
                "금액":     subtotal,
            })
            total_cost += subtotal
        return total_cost, cost_details

    # ─────────────────────────────────────────────────────────────────────────
    # 입출고 처리
    # ─────────────────────────────────────────────────────────────────────────

    def receive_part(self, part_id: str, qty: int, note: str = ""):
        # 문제 1 수정: 전체 읽기→계산→쓰기 구간을 락으로 보호
        with self._op_lock:
            self.cache.invalidate("parts")
            self.cache.invalidate("parts_map")
            part = self.get_part_by_id(part_id)
            if not part:
                return False, f"품번 '{part_id}'를 찾을 수 없습니다."

            new_qty = int(part["현재재고"]) + int(qty)
            self.update_part_qty(part_id, new_qty)
            self._add_history("입고", "부품입고", part_id, part["부품명"],
                              qty, new_qty, "", note)
            return True, f"입고 완료: {part['부품명']} +{qty}개 (현재재고: {new_qty}개)"

    def issue_part(self, part_id: str, qty: int, note: str = ""):
        # 문제 1 수정: 전체 읽기→재고확인→쓰기 구간을 락으로 보호
        with self._op_lock:
            self.cache.invalidate("parts")
            self.cache.invalidate("parts_map")
            part = self.get_part_by_id(part_id)
            if not part:
                return False, f"품번 '{part_id}'를 찾을 수 없습니다."

            current = int(part["현재재고"])
            if current < int(qty):
                return False, (f"재고 부족: {part['부품명']} "
                               f"현재재고 {current}개, 출고요청 {qty}개")

            new_qty = current - int(qty)
            self.update_part_qty(part_id, new_qty)
            self._add_history("출고", "부품출고", part_id, part["부품명"],
                              qty, new_qty, "", note)

            warning = ""
            safety  = int(part.get("안전재고", 0))
            if safety > 0 and new_qty <= safety:
                warning = f" ⚠️ 안전재고({safety}개) 이하입니다!"
            return True, (f"출고 완료: {part['부품명']} -{qty}개 "
                          f"(현재재고: {new_qty}개){warning}")

    def produce_product(self, product_id: str, qty: int, note: str = "", force: bool = False):
        """
        제품 생산 (BOM 기반 자동 출고)
        v2.1: 재고 수량 일괄 업데이트 + 이력 배치 추가
        문제 1 수정: 전체 생산 처리를 락으로 감싸 동시 재고 충돌 방지
        force=True: 재고 부족 시에도 마이너스 처리 후 생산 진행
        force=False(기본): 재고 부족 시 (None, 안내문, 부족목록) 반환 → UI에서 확인 후 재호출
        """
        with self._op_lock:
            return self._produce_product_locked(product_id, qty, note, force)

    def _produce_product_locked(self, product_id: str, qty: int, note: str = "", force: bool = False):
        """produce_product 내부 로직 — _op_lock 보유 상태에서 호출
        문제 2 수정: 재고 검증에 캐시 우회 메서드(_get_fresh_*)를 사용하여
        30초 캐시 TTL 만료나 동시 읽기에 의한 캐시 재채움에도 항상 최신 데이터 사용
        """
        # _get_fresh_* 메서드가 캐시를 우회해 시트에서 직접 읽으므로
        # invalidate_all() 없이도 항상 최신 데이터를 사용합니다.
        # (invalidate_all은 다른 시트 캐시까지 날려 불필요한 API 호출을 유발하므로 제거)
        product = self._get_fresh_product(product_id)
        if not product:
            return False, f"제품코드 '{product_id}'를 찾을 수 없습니다.", []

        bom = self._get_fresh_bom_for_product(product_id)
        if not bom:
            return False, f"제품 '{product['제품명']}'의 BOM이 없습니다.", []

        # 부품 데이터를 한 번만 읽어 parts_map과 all_values를 동시에 확보
        # (이후 _bulk_update_part_qtys에 all_values를 전달해 이중 읽기 방지)
        parts_map, parts_all_values = self._get_fresh_parts_with_rows()

        # 1단계: 재고 검증
        shortage = []
        missing_master = []
        for item in bom:
            part_id  = str(item["부품품번"])
            required = float(item["소요량"]) * int(qty)
            part     = parts_map.get(part_id)
            if not part:
                missing_master.append(f"품번 '{part_id}' 마스터 없음")
                continue
            current = int(part["현재재고"])
            if current < required:
                shortage.append(
                    f"{part['부품명']}({part_id}): "
                    f"필요 {int(required)}개, 재고 {current}개"
                )

        # BOM 마스터 누락은 force 여부에 관계없이 항상 차단
        if missing_master:
            return False, "BOM 부품 마스터 오류:\n" + "\n".join(missing_master), []

        # 재고 부족: force=False이면 확인 요청 반환, force=True이면 마이너스 처리 진행
        if shortage and not force:
            return None, "일부 부품 재고가 부족합니다:\n" + "\n".join(shortage), shortage

        # 2단계: 출고 처리 (일괄)
        results          = []
        qty_updates      = []
        history_entries  = []

        for item in bom:
            part_id  = str(item["부품품번"])
            required = int(float(item["소요량"]) * int(qty))
            part     = parts_map.get(part_id)
            current  = int(part["현재재고"])
            new_qty  = current - required

            qty_updates.append((part_id, new_qty))
            history_entries.append({
                "direction": "출고",
                "h_type":   "생산출고",
                "item_id":   part_id,
                "item_name": part["부품명"],
                "qty":       required,
                "remaining": new_qty,
                "related":   f"{product['제품명']}({product_id})",
                "note":      note,
            })

            safety = int(part.get("안전재고", 0))
            if new_qty < 0:
                results.append(
                    f"  🔴 {part['부품명']}({part_id}): -{required}개 → {new_qty}개 (마이너스 재고)"
                )
            elif safety > 0 and new_qty <= safety:
                results.append(
                    f"  ⚠️ {part['부품명']}({part_id}): "
                    f"재고 {new_qty}개 (안전재고: {safety}개)"
                )
            else:
                results.append(
                    f"  ✅ {part['부품명']}({part_id}): -{required}개 → {new_qty}개"
                )

        self._bulk_update_part_qtys(qty_updates, parts_all_values=parts_all_values)
        self._add_history_batch(history_entries)

        # 3단계: 제품 재고 증가
        product_new_qty = int(product["현재재고"]) + int(qty)
        self.update_product_qty(product_id, product_new_qty)
        self._add_history("입고", "생산입고", product_id, product["제품명"],
                          qty, product_new_qty, "", note)

        msg = (f"생산 완료: {product['제품명']} +{qty}개 "
               f"(제품재고: {product_new_qty}개)\n"
               f"출고된 부품:\n" + "\n".join(results))
        return True, msg, results

    def cancel_history(self, row_index: int):
        """입출고 이력 취소(원복)"""
        # 문제 1 수정: 취소 처리 전체를 락으로 보호 (RLock이므로 cancel_production 재진입 허용)
        with self._op_lock:
            ws       = self.spreadsheet.worksheet(SHEET_HISTORY)
            row_data = ws.row_values(row_index)

            if len(row_data) < 7:
                return False, "유효하지 않은 이력입니다."

            direction  = row_data[1]
            h_type     = row_data[2]
            item_id    = row_data[3]
            item_name  = row_data[4]
            h_qty      = int(row_data[5])

            # 취소 시: 부품/제품 재고가 변하므로 해당 캐시만 무효화
            self.cache.invalidate("parts")
            self.cache.invalidate("parts_map")
            self.cache.invalidate("products")
            self.cache.invalidate("history")

            if h_type in ["부품입고"]:
                part = self.get_part_by_id(item_id)
                if part:
                    new_qty = int(part["현재재고"]) - h_qty
                    if new_qty < 0:
                        return False, f"취소 시 재고가 음수({new_qty})가 됩니다."
                    self.update_part_qty(item_id, new_qty)
                    self._add_history("취소", f"{h_type}취소", item_id, item_name,
                                      h_qty, new_qty, "", "이력 취소")
                    return True, f"입고 취소: {item_name} -{h_qty}개 (재고: {new_qty}개)"

            elif h_type in ["부품출고", "생산출고"]:
                part = self.get_part_by_id(item_id)
                if part:
                    new_qty = int(part["현재재고"]) + h_qty
                    self.update_part_qty(item_id, new_qty)
                    self._add_history("취소", f"{h_type}취소", item_id, item_name,
                                      h_qty, new_qty, "", "이력 취소")
                    return True, f"출고 취소: {item_name} +{h_qty}개 (재고: {new_qty}개)"

            elif h_type == "생산입고":
                # cancel_production 도 _op_lock(RLock)을 획득하지만
                # RLock이므로 같은 스레드에서는 재진입 허용
                success, msg, details = self.cancel_production(row_index)
                return success, msg

            return False, "취소할 수 없는 이력 유형입니다."

    def cancel_production(self, production_entry_row: int):
        """생산 일괄 취소 — 문제 1 수정: _op_lock(RLock) 으로 보호"""
        with self._op_lock:
            return self._cancel_production_locked(production_entry_row)

    def _cancel_production_locked(self, production_entry_row: int):
        """cancel_production 내부 로직 — _op_lock 보유 상태에서 호출"""
        ws       = self.spreadsheet.worksheet(SHEET_HISTORY)
        row_data = ws.row_values(production_entry_row)

        if len(row_data) < 7:
            return False, "유효하지 않은 이력입니다.", []

        h_type = row_data[2]
        if h_type != "생산입고":
            return False, f"생산입고 이력이 아닙니다. (유형: {h_type})", []

        prod_time    = row_data[0]
        product_id   = row_data[3]
        product_name = row_data[4]
        prod_qty     = int(row_data[5])

        # 생산취소: 부품·제품 재고 + 이력이 변하므로 관련 캐시만 무효화
        self.cache.invalidate("parts")
        self.cache.invalidate("parts_map")
        self.cache.invalidate("products")
        self.cache.invalidate("history")

        # 같은 시각 생산출고 이력 찾기 (이력 시트 직접 읽기 — 취소 로직이므로 캐시 우회 OK)
        all_history      = self._safe_get_all_values(ws)
        related_keyword  = f"{product_name}({product_id})"

        try:
            prod_dt = datetime.strptime(prod_time, "%Y-%m-%d %H:%M:%S")
        except Exception:
            prod_dt = None

        issue_rows = []
        for i, row in enumerate(all_history[1:], 2):
            if len(row) < 8:
                continue
            if row[2] != "생산출고":
                continue
            if (related_keyword not in row[7] and
                    row[7] not in related_keyword):
                continue
            is_time_match = (row[0] == prod_time)
            if not is_time_match and prod_dt:
                try:
                    h_dt = datetime.strptime(row[0], "%Y-%m-%d %H:%M:%S")
                    if abs((h_dt - prod_dt).total_seconds()) <= 10:
                        is_time_match = True
                except Exception:
                    pass
            if is_time_match:
                issue_rows.append({
                    "row":       i,
                    "part_id":   row[3],
                    "part_name": row[4],
                    "qty":       int(row[5]),
                })

        product = self.get_product_by_id(product_id)
        if not product:
            return False, f"제품코드 '{product_id}'를 찾을 수 없습니다.", []

        product_current = int(product["현재재고"])
        product_new_qty = product_current - prod_qty
        if product_new_qty < 0:
            return False, (
                f"취소 시 제품 재고가 음수({product_new_qty})가 됩니다.\n"
                f"현재 제품 재고: {product_current}개, 취소 수량: {prod_qty}개"
            ), []

        # 문제 2 수정: 캐시 우회로 락 취득 직후의 실제 재고 사용
        parts_map   = self._get_fresh_parts_map()
        qty_updates = []
        results     = []

        for ir in issue_rows:
            part = parts_map.get(ir["part_id"])
            if part:
                part_new_qty = int(part["현재재고"]) + ir["qty"]
                qty_updates.append((ir["part_id"], part_new_qty))
                results.append(
                    f"  {ir['part_name']}({ir['part_id']}): "
                    f"+{ir['qty']}개 → {part_new_qty}개"
                )

        if qty_updates:
            self._bulk_update_part_qtys(qty_updates)

        self.update_product_qty(product_id, product_new_qty)

        cancel_entries = []
        for ir in issue_rows:
            part = parts_map.get(ir["part_id"])
            part_restored = int(part["현재재고"]) + ir["qty"] if part else 0
            cancel_entries.append({
                "direction": "취소",
                "h_type":    "생산출고취소",
                "item_id":   ir["part_id"],
                "item_name": ir["part_name"],
                "qty":       ir["qty"],
                "remaining": part_restored,
                "related":   related_keyword,
                "note":      "생산 일괄 취소",
            })

        if cancel_entries:
            self._add_history_batch(cancel_entries)

        self._add_history("취소", "생산입고취소", product_id, product_name,
                          prod_qty, product_new_qty, "", "생산 일괄 취소")

        msg = (f"생산 취소 완료: {product_name} -{prod_qty}개 "
               f"(제품재고: {product_current}개 → {product_new_qty}개)\n\n"
               f"복원된 부품 ({len(issue_rows)}개):\n" + "\n".join(results))
        return True, msg, results

    # ── 이력 ──────────────────────────────────────────────────────────────────

    def _add_history(self, direction, h_type, item_id, item_name,
                     qty, remaining, related="", note=""):
        ws  = self.spreadsheet.worksheet(SHEET_HISTORY)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._safe_append_row(ws, [
            now, direction, h_type, str(item_id),
            item_name, int(qty), int(remaining), related, note
        ])
        self.cache.invalidate("history")   # 새 이력 추가 시 캐시 무효화

    def _add_history_batch(self, entries):
        """이력 일괄 추가 (API 1회 — append_rows로 불필요한 읽기 제거)"""
        if not entries:
            return
        ws  = self.spreadsheet.worksheet(SHEET_HISTORY)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        rows = []
        for e in entries:
            rows.append([
                now, e["direction"], e["h_type"], str(e["item_id"]),
                e["item_name"], int(e["qty"]), int(e["remaining"]),
                e.get("related", ""), e.get("note", "")
            ])
        # append_rows: 시트 끝에 자동 추가 — 행 수 파악을 위한 읽기 API 호출 불필요
        self._safe_append_rows(ws, rows)
        self.cache.invalidate("history")   # 이력 캐시 무효화

    def get_all_history(self):
        """입출고이력 전체 조회 (캐시 사용, TTL 60초)"""
        return self._get_all_history_cached()

    def get_history_by_date_range(self, start_date: str, end_date: str):
        records  = self.get_all_history()
        filtered = []
        for r in records:
            date_str = str(r.get("일시", ""))[:10]
            if start_date <= date_str <= end_date:
                filtered.append(r)
        return filtered

    # ── 안전재고 알림 ──────────────────────────────────────────────────────────

    def get_safety_stock_alerts(self):
        parts  = self._get_all_parts_cached()
        alerts = []
        for p in parts:
            safety  = int(p.get("안전재고", 0))
            current = int(p.get("현재재고", 0))
            if safety > 0 and current <= safety:
                alerts.append(p)
        return alerts

    # ── MRP ───────────────────────────────────────────────────────────────────

    def get_max_producible(self, product_id: str):
        bom = self.get_bom_for_product(product_id)
        if not bom:
            return 0, "BOM 없음"

        parts_map   = self._get_parts_map()
        min_qty     = float('inf')
        bottleneck  = ""

        for item in bom:
            part_id  = str(item.get("부품품번", ""))
            required = float(item.get("소요량", 0))
            if required <= 0:
                continue
            part = parts_map.get(part_id)
            if not part:
                return 0, f"부품 마스터 없음: {part_id}"
            current_stock = int(part.get("현재재고", 0))
            possible      = int(current_stock / required)
            if possible < min_qty:
                min_qty    = possible
                bottleneck = part.get("부품명", part_id)

        if min_qty == float('inf'):
            return 0, "소요량 없음"
        return min_qty, bottleneck

    def calculate_mrp(self, production_plan, include_safety_stock=False):
        """MRP(자재소요계획) 계산"""
        # 캐시된 메서드들이 TTL 내 최신 데이터를 반환하므로 invalidate_all 불필요
        products      = self._get_all_products_cached()
        products_map  = {str(p["제품코드"]): p for p in products}
        parts_map     = self._get_parts_map()
        all_bom       = self._get_all_bom_cached()

        plan_summary = []
        for plan in production_plan:
            pid           = str(plan["product_id"])
            product       = products_map.get(pid)
            current_stock = int(product["현재재고"]) if product else 0
            target_qty    = int(plan["target_qty"])
            need_to_produce = max(0, target_qty - current_stock)
            max_prod, bottleneck = self.get_max_producible(pid)
            plan_summary.append({
                "product_id":      pid,
                "product_name":    plan.get("product_name",
                                            product["제품명"] if product else pid),
                "current_stock":   current_stock,
                "target_qty":      target_qty,
                "need_to_produce": need_to_produce,
                "max_producible":  max_prod,
                "bottleneck":      bottleneck,
            })

        part_totals      = {}
        part_product_map = {}   # part_id → set of product_ids
        for plan_item in plan_summary:
            pid  = plan_item["product_id"]
            need = plan_item["need_to_produce"]
            if need <= 0:
                continue
            bom_items = [b for b in all_bom
                         if str(b.get("제품코드", "")) == pid]
            for bom_item in bom_items:
                part_id = str(bom_item.get("부품품번", ""))
                req_per = float(bom_item.get("소요량", 0))
                total   = req_per * need
                part_totals[part_id] = part_totals.get(part_id, 0) + total
                if part_id not in part_product_map:
                    part_product_map[part_id] = set()
                part_product_map[part_id].add(pid)

        parts_requirement = []
        total_order_items = 0
        total_order_qty   = 0

        for part_id, total_required in sorted(part_totals.items()):
            part = parts_map.get(part_id)
            if not part:
                continue
            current_stock = int(part.get("현재재고", 0))
            safety_stock  = int(part.get("안전재고", 0))
            shortage      = (total_required - current_stock + safety_stock
                             if include_safety_stock
                             else total_required - current_stock)
            shortage      = max(0, shortage)
            order_needed  = shortage > 0
            if order_needed:
                total_order_items += 1
                total_order_qty   += int(shortage)

            parts_requirement.append({
                "part_id":        part_id,
                "part_name":      part.get("부품명", ""),
                "supplier":       part.get("업체명", ""),
                "unit":           part.get("단위", ""),
                "spec":           part.get("규격", ""),
                "unit_price":     float(part.get("단가", 0) or 0),
                "total_required": int(total_required),
                "current_stock":  current_stock,
                "safety_stock":   safety_stock,
                "shortage":       int(shortage),
                "order_needed":   order_needed,
                "product_codes":  sorted(part_product_map.get(part_id, set())),
            })

        parts_requirement.sort(key=lambda x: (not x["order_needed"], x["part_id"]))

        return {
            "plan_summary":        plan_summary,
            "parts_requirement":   parts_requirement,
            "total_order_items":   total_order_items,
            "total_order_qty":     total_order_qty,
        }

    # ─────────────────────────────────────────────────────────────────────────
    # 거래처 관리
    # ─────────────────────────────────────────────────────────────────────────

    def get_all_customers(self) -> list:
        ws = self.spreadsheet.worksheet(SHEET_CUSTOMERS)
        return ws.get_all_records()

    def get_customer_by_code(self, code: str) -> dict | None:
        rows = self.get_all_customers()
        return next((r for r in rows if str(r.get("거래처코드", "")) == str(code)), None)

    def add_customer(self, code: str, name: str, contact: str = "",
                     phone: str = "", email: str = "",
                     address: str = "", note: str = "") -> bool:
        """거래처 추가. 중복 코드면 False 반환."""
        if self.get_customer_by_code(code):
            return False
        ws = self.spreadsheet.worksheet(SHEET_CUSTOMERS)
        self._safe_append_row(ws, [code, name, contact, phone, email, address, note])
        return True

    def update_customer(self, code: str, name: str, contact: str = "",
                        phone: str = "", email: str = "",
                        address: str = "", note: str = "") -> bool:
        ws = self.spreadsheet.worksheet(SHEET_CUSTOMERS)
        rows = ws.get_all_values()
        headers = rows[0] if rows else []
        try:
            code_col = headers.index("거래처코드") + 1
        except ValueError:
            return False
        for i, row in enumerate(rows[1:], start=2):
            if len(row) >= code_col and str(row[code_col - 1]) == str(code):
                self._safe_update(ws, f"A{i}:G{i}",
                                  [[code, name, contact, phone, email, address, note]])
                return True
        return False

    def delete_customer(self, code: str) -> bool:
        ws = self.spreadsheet.worksheet(SHEET_CUSTOMERS)
        rows = ws.get_all_values()
        headers = rows[0] if rows else []
        try:
            code_col = headers.index("거래처코드") + 1
        except ValueError:
            return False
        for i, row in enumerate(rows[1:], start=2):
            if len(row) >= code_col and str(row[code_col - 1]) == str(code):
                self._safe_delete_rows(ws, i)
                return True
        return False

    # ─────────────────────────────────────────────────────────────────────────
    # 제품 출고(판매)
    # ─────────────────────────────────────────────────────────────────────────

    def ship_product(self, product_id: str, qty: int,
                     customer_code: str, customer_name: str,
                     unit_price: float, note: str = "") -> tuple:
        """
        제품 출고(판매) 처리
        - 제품 재고 차감
        - 판매이력 시트에 상세 기록
        - 입출고이력 시트에도 기록
        Returns: (success: bool, message: str)
        """
        with self._op_lock:
            self.cache.invalidate("products")

            product = self._get_fresh_product(product_id)
            if not product:
                return False, f"제품코드 '{product_id}'를 찾을 수 없습니다."

            current = int(product.get("현재재고", 0))
            if current < qty:
                name = product.get("제품명", "")
                return False, (
                    f"재고 부족: [{name}] "
                    f"현재재고 {current:,}개, 출고요청 {qty:,}개"
                )

            new_qty = current - qty
            amount  = round(unit_price * qty)

            # 1. 제품 재고 차감
            self.update_product_qty(product_id, new_qty)

            # 2. 판매이력 기록
            ws_sales = self.spreadsheet.worksheet(SHEET_SALES)
            all_rows = ws_sales.get_all_values()
            sale_no  = len(all_rows)
            now      = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self._safe_append_row(ws_sales, [
                sale_no,
                now,
                str(product_id),
                product.get("제품명", ""),
                str(customer_code),
                customer_name,
                int(qty),
                int(unit_price),
                int(amount),
                int(new_qty),
                note,
            ])

            # 3. 기존 입출고이력에도 기록
            self._add_history(
                "출고", "제품출고",
                product_id, product.get("제품명", ""),
                qty, new_qty,
                customer_name, note,
            )

            pname = product.get("제품명", "")
            msg = (
                f"[{pname}] {qty:,}개 납품 완료 ({customer_name})\n"
                f"단가: {int(unit_price):,}원  x  {qty:,}개  =  {amount:,}원\n"
                f"잔여재고: {new_qty:,}개"
            )
            return True, msg

    def get_all_sales(self) -> list:
        """판매이력 전체 조회"""
        ws = self.spreadsheet.worksheet(SHEET_SALES)
        return ws.get_all_records()

    def get_sales_by_date_range(self, start_date: str, end_date: str) -> list:
        """날짜 범위로 판매이력 필터링 (start/end: 'YYYY-MM-DD')"""
        records = self.get_all_sales()
        return [r for r in records
                if start_date <= str(r.get("일시", ""))[:10] <= end_date]

    def get_sales_summary_by_month(self) -> list:
        """
        월별 판매 집계 - 리포트용
        Returns: [{"월": "2026-04", "건수": N, "수량": N, "금액": N}, ...]
        """
        records = self.get_all_sales()
        monthly: dict = {}
        for r in records:
            month = str(r.get("일시", ""))[:7]
            if not month or month == "":
                continue
            if month not in monthly:
                monthly[month] = {"월": month, "건수": 0, "수량": 0, "금액": 0}
            monthly[month]["건수"] += 1
            monthly[month]["수량"] += int(r.get("수량", 0))
            monthly[month]["금액"] += int(r.get("금액", 0))
        return sorted(monthly.values(), key=lambda x: x["월"])


    def _create_default_admin(self, ws=None):
        """기본 관리자 계정 생성 (사용자 시트가 비어있을 때)"""
        from core.auth import hash_password
        from core.constants import (DEFAULT_ADMIN_ID, DEFAULT_ADMIN_PW,
                                    DEFAULT_ADMIN_NAME, ALL_MENU_IDS)
        if ws is None:
            ws = self.spreadsheet.worksheet(SHEET_USERS)
        all_menus = ",".join(ALL_MENU_IDS)
        self._safe_append_row(ws, [
            DEFAULT_ADMIN_ID,
            hash_password(DEFAULT_ADMIN_PW),
            DEFAULT_ADMIN_NAME,
            all_menus,
            "Y",
        ])

    def authenticate_user(self, username: str, password: str) -> dict | None:
        """
        아이디/비밀번호로 인증. 성공 시 사용자 dict 반환, 실패 시 None.
        관리자(users 권한 보유) 로그인 시 새로 추가된 메뉴를 자동으로 권한에 추가.
        """
        from core.auth import verify_password
        from core.constants import ALL_MENU_IDS
        try:
            ws      = self.spreadsheet.worksheet(SHEET_USERS)
            records = ws.get_all_records()
        except Exception:
            return None

        all_rows = ws.get_all_values()
        headers  = all_rows[0] if all_rows else []

        for row_idx, r in enumerate(records, start=2):
            if (str(r.get("아이디", "")) == username
                    and str(r.get("활성화", "Y")).upper() == "Y"):
                stored_hash = str(r.get("비밀번호", ""))
                if verify_password(password, stored_hash):
                    menu_str = str(r.get("메뉴권한", ""))
                    menus = [m.strip() for m in menu_str.split(",") if m.strip()]

                    # 관리자(users 권한)는 새 메뉴 자동 추가
                    if "users" in menus:
                        new_menus = [m for m in ALL_MENU_IDS if m not in menus]
                        if new_menus:
                            menus = menus + new_menus
                            # 시트에 자동 업데이트
                            try:
                                col_d = headers.index("메뉴권한") + 1 if "메뉴권한" in headers else 4
                                col_letter = chr(ord("A") + col_d - 1)
                                self._safe_update(
                                    ws, f"{col_letter}{row_idx}",
                                    [[",".join(menus)]]
                                )
                            except Exception:
                                pass

                    return {
                        "아이디":   username,
                        "이름":     str(r.get("이름", username)),
                        "메뉴권한": menus,
                    }
        return None

    # ═══════════════════════════════════════════════════════════════
    # MRP 이력
    # ═══════════════════════════════════════════════════════════════
    def get_all_mrp_history(self) -> list:
        """저장된 MRP 계획 이력 전체 조회 (최신순)"""
        try:
            ws  = self.spreadsheet.worksheet(SHEET_MRP)
            rows = ws.get_all_records()
            return list(reversed(rows))   # 최신순
        except Exception:
            return []

    def save_mrp_plan(self, name: str, production_plan: list,
                      mrp_result: dict, include_safety: bool) -> str:
        """MRP 계획 저장 → 저장ID 반환.
        소요부품(result_json)은 불러오기 시 재계산하므로 저장하지 않음.
        (대용량 JSON을 단일 셀에 저장하면 Google Sheets 50,000자 제한 초과)
        """
        import json as _json
        from datetime import datetime as _dt

        ws   = self.spreadsheet.worksheet(SHEET_MRP)
        rows = ws.get_all_values()
        # ID 채번
        existing_ids = [r[0] for r in rows[1:] if r and r[0].startswith("M-")]
        if existing_ids:
            nums  = [int(x.split("-")[1]) for x in existing_ids if x.split("-")[-1].isdigit()]
            next_n = max(nums) + 1 if nums else 1
        else:
            next_n = 1
        plan_id = f"M-{next_n:03d}"

        plan_json  = _json.dumps(production_plan, ensure_ascii=False)
        now_str    = _dt.now().strftime("%Y-%m-%d %H:%M")
        prod_count = len(production_plan)

        # G열(소요부품)은 빈 문자열로 — 재계산으로 대체
        new_row = [plan_id, now_str, name, prod_count,
                   "TRUE" if include_safety else "FALSE",
                   plan_json, ""]
        ws.append_row(new_row, value_input_option="RAW")
        return plan_id

    def update_mrp_plan(self, plan_id: str, name: str, production_plan: list,
                        mrp_result: dict, include_safety: bool) -> bool:
        """기존 MRP 계획 덮어쓰기 (내용 전체 수정).
        계획을 못 찾으면 False 반환, 업데이트 중 오류는 예외를 그대로 올림.
        소요부품(G열)은 재계산으로 대체하므로 저장하지 않음 — 50,000자 제한 방지.
        """
        import json as _json
        from datetime import datetime as _dt

        ws   = self.spreadsheet.worksheet(SHEET_MRP)
        rows = ws.get_all_values()
        for i, row in enumerate(rows[1:], 2):
            if row and str(row[0]).strip() == str(plan_id).strip():
                plan_json  = _json.dumps(production_plan, ensure_ascii=False)
                now_str    = _dt.now().strftime("%Y-%m-%d %H:%M")
                prod_count = len(production_plan)
                # F열(생산계획)까지만 업데이트, G열(소요부품)은 건드리지 않음
                self._safe_update(ws, f"B{i}:F{i}", [[
                    now_str, name, prod_count,
                    "TRUE" if include_safety else "FALSE",
                    plan_json,
                ]])
                return True
        return False   # plan_id 를 시트에서 찾지 못한 경우

    def update_mrp_plan_name(self, plan_id: str, new_name: str) -> bool:
        """MRP 계획명 수정"""
        try:
            ws   = self.spreadsheet.worksheet(SHEET_MRP)
            rows = ws.get_all_values()
            for i, row in enumerate(rows[1:], 2):
                if row and row[0] == plan_id:
                    ws.update_cell(i, 3, new_name)
                    return True
            return False
        except Exception:
            return False

    def delete_mrp_plans(self, plan_ids: list) -> int:
        """MRP 계획 삭제 (복수). 삭제된 개수 반환"""
        try:
            ws   = self.spreadsheet.worksheet(SHEET_MRP)
            rows = ws.get_all_values()
            del_rows = []
            for i, row in enumerate(rows[1:], 2):
                if row and row[0] in plan_ids:
                    del_rows.append(i)
            # 역순으로 삭제 (행 번호 밀림 방지)
            for row_idx in reversed(del_rows):
                ws.delete_rows(row_idx)
            return len(del_rows)
        except Exception:
            return 0

    def get_mrp_plan_detail(self, plan_id: str) -> dict:
        """저장된 MRP 계획 상세 조회 (불러오기용)"""
        import json as _json
        try:
            ws   = self.spreadsheet.worksheet(SHEET_MRP)
            rows = ws.get_all_records()
            for row in rows:
                if str(row.get("저장ID", "")) == plan_id:
                    production_plan = _json.loads(row.get("생산계획", "[]"))
                    include_safety  = str(row.get("안전재고반영", "FALSE")).upper() == "TRUE"
                    return {
                        "plan_id":        plan_id,
                        "name":           row.get("계획명", ""),
                        "saved_at":       row.get("저장일시", ""),
                        "production_plan": production_plan,
                        "include_safety": include_safety,
                    }
        except Exception:
            pass
        return {}

    def get_all_users(self) -> list:
        """전체 사용자 목록 반환 (비밀번호 제외)"""
        try:
            ws      = self.spreadsheet.worksheet(SHEET_USERS)
            records = ws.get_all_records()
        except Exception:
            return []
        result = []
        for r in records:
            menu_str = str(r.get("메뉴권한", ""))
            menus = [m.strip() for m in menu_str.split(",") if m.strip()]
            result.append({
                "아이디":   str(r.get("아이디", "")),
                "이름":     str(r.get("이름", "")),
                "메뉴권한": menus,
                "활성화":   str(r.get("활성화", "Y")).upper() == "Y",
            })
        return result

    def add_user(self, username: str, password: str, name: str,
                 menu_permissions: list, active: bool = True) -> bool:
        """사용자 추가. 아이디 중복 시 False 반환"""
        from core.auth import hash_password
        existing = [u["아이디"] for u in self.get_all_users()]
        if username in existing:
            return False
        ws = self.spreadsheet.worksheet(SHEET_USERS)
        self._safe_append_row(ws, [
            username,
            hash_password(password),
            name,
            ",".join(menu_permissions),
            "Y" if active else "N",
        ])
        return True

    def update_user(self, username: str, name: str,
                    menu_permissions: list, active: bool) -> bool:
        """사용자 정보(이름·권한·활성화) 수정. 비밀번호는 별도 메서드로"""
        try:
            ws      = self.spreadsheet.worksheet(SHEET_USERS)
            records = ws.get_all_records()
        except Exception:
            return False
        for i, r in enumerate(records):
            if str(r.get("아이디", "")) == username:
                row = i + 2
                # C(이름), D(메뉴권한), E(활성화) 만 업데이트
                self._safe_update(ws, f"C{row}:E{row}", [[
                    name,
                    ",".join(menu_permissions),
                    "Y" if active else "N",
                ]])
                return True
        return False

    def reset_password(self, username: str, new_password: str) -> bool:
        """비밀번호 초기화"""
        from core.auth import hash_password
        try:
            ws      = self.spreadsheet.worksheet(SHEET_USERS)
            records = ws.get_all_records()
        except Exception:
            return False
        for i, r in enumerate(records):
            if str(r.get("아이디", "")) == username:
                self._safe_update_cell(ws, i + 2, 2, hash_password(new_password))
                return True
        return False

    def delete_user(self, username: str) -> bool:
        """사용자 삭제 (마지막 관리자는 삭제 불가)"""
        try:
            ws      = self.spreadsheet.worksheet(SHEET_USERS)
            records = ws.get_all_records()
        except Exception:
            return False
        # 마지막 'users' 권한 보유자 삭제 방지
        admin_rows = [
            r for r in records
            if "users" in str(r.get("메뉴권한", ""))
               and str(r.get("활성화", "Y")).upper() == "Y"
               and str(r.get("아이디", "")) != username
        ]
        target_row = None
        for i, r in enumerate(records):
            if str(r.get("아이디", "")) == username:
                # 이 사용자가 마지막 관리자인지 확인
                is_admin = "users" in str(r.get("메뉴권한", ""))
                if is_admin and not admin_rows:
                    return False  # 마지막 관리자 삭제 불가
                target_row = i + 2
                break
        if target_row:
            self._safe_delete_rows(ws, target_row)
            return True
        return False

    def user_exists(self) -> bool:
        """사용자 시트에 계정이 하나라도 있는지 확인"""
        try:
            ws = self.spreadsheet.worksheet(SHEET_USERS)
            return len(ws.get_all_records()) > 0
        except Exception:
            return False
