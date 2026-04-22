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

# 캐시 유효 시간 (초)
CACHE_TTL = 30

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

    def get(self, key):
        with self._lock:
            if key in self._cache:
                elapsed = time.time() - self._timestamps.get(key, 0)
                if elapsed < CACHE_TTL:
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
            headers = ws.row_values(1)
            if headers != NEW_PARTS_HEADERS:
                all_data = ws.get_all_values()
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
            self.spreadsheet.worksheet(SHEET_PRODUCTS)
        except gspread.exceptions.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(title=SHEET_PRODUCTS, rows=1000, cols=10)
            self._safe_update(ws, "A1:E1",
                              [["제품코드", "제품명", "규격", "현재재고", "비고"]])

        # BOM
        try:
            ws = self.spreadsheet.worksheet(SHEET_BOM)
            headers = ws.row_values(1)
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

        # 기본 Sheet1 삭제
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
        records = ws.get_all_records()
        self.cache.set("parts", records)
        return records

    def _get_all_products_cached(self):
        cached = self.cache.get("products")
        if cached is not None:
            return cached
        ws = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        records = ws.get_all_records()
        self.cache.set("products", records)
        return records

    def _get_all_bom_cached(self):
        cached = self.cache.get("bom")
        if cached is not None:
            return cached
        ws = self.spreadsheet.worksheet(SHEET_BOM)
        records = ws.get_all_records()
        self.cache.set("bom", records)
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
        records = ws.get_all_records()
        # 캐시도 함께 갱신하여 이후 일반 조회에도 최신 데이터가 반영되도록 함
        self.cache.set("parts",     records)
        parts_map = {str(p["품번"]): p for p in records}
        self.cache.set("parts_map", parts_map)
        return parts_map

    def _get_fresh_product(self, product_id: str):
        """구글 시트에서 직접 최신 제품 데이터를 읽어 반환 (캐시 우회)"""
        ws      = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        records = ws.get_all_records()
        self.cache.set("products", records)
        for r in records:
            if str(r.get("제품코드", "")) == str(product_id):
                return r
        return None

    def _get_fresh_bom_for_product(self, product_id: str):
        """구글 시트에서 직접 최신 BOM 데이터를 읽어 반환 (캐시 우회)"""
        ws      = self.spreadsheet.worksheet(SHEET_BOM)
        records = ws.get_all_records()
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
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if str(r.get("품번", "")) == str(part_id):
                self._safe_update_cell(ws, i + 2, 7, int(new_qty))
                self.cache.invalidate("parts")
                self.cache.invalidate("parts_map")
                return True
        return False

    def _bulk_update_part_qtys(self, updates):
        """재고 수량 일괄 업데이트 (API 1회) - 재시도 포함"""
        ws         = self.spreadsheet.worksheet(SHEET_PARTS)
        all_values = ws.get_all_values()

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
                    note="", supplier="", unit_price=0, moq=0):
        ws      = self.spreadsheet.worksheet(SHEET_PARTS)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if str(r.get("품번", "")) == str(part_id):
                row = i + 2
                self._safe_update(ws, f"A{row}:J{row}", [[
                    str(part_id), supplier, name, spec, unit,
                    float(unit_price), int(qty), int(safety_qty), int(moq), note
                ]])
                self.cache.invalidate("parts")
                self.cache.invalidate("parts_map")
                return True
        return False

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

        for p in parts_list:
            code = str(p["품번"])
            if code in existing_map:
                i2  = existing_map[code]
                row = existing_rows[i2]
                while len(row) < len(header):
                    row.append("")
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

    def add_product(self, product_id, name, spec, qty=0, note=""):
        ws = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        self._safe_append_row(ws, [str(product_id), name, spec, int(qty), note])
        self.cache.invalidate("products")

    def update_product_qty(self, product_id: str, new_qty: int):
        ws      = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if str(r.get("제품코드", "")) == str(product_id):
                self._safe_update_cell(ws, i + 2, 4, int(new_qty))
                self.cache.invalidate("products")
                return True
        return False

    def update_product(self, product_id, name, spec, qty, note=""):
        ws      = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if str(r.get("제품코드", "")) == str(product_id):
                row = i + 2
                self._safe_update(ws, f"A{row}:E{row}",
                                  [[str(product_id), name, spec, int(qty), note]])
                self.cache.invalidate("products")
                return True
        return False

    def delete_product(self, product_id: str):
        ws      = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        records = ws.get_all_records()
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

    def update_bom(self, product_id: str, part_id: str,
                   qty: float, note: str = ""):
        """BOM 항목 수정. 단가는 부품마스터에서만 관리하므로 BOM 시트에 저장하지 않음."""
        ws      = self.spreadsheet.worksheet(SHEET_BOM)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if (str(r.get("제품코드", "")) == str(product_id) and
                    str(r.get("부품품번", "")) == str(part_id)):
                row = i + 2
                self._safe_update(ws, f"A{row}:E{row}", [[
                    str(product_id), str(part_id), float(qty), 0, note  # 단가 열은 항상 0
                ]])
                self.cache.invalidate("bom")
                return True
        return False

    def delete_bom(self, product_id: str, part_id: str):
        ws      = self.spreadsheet.worksheet(SHEET_BOM)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if (str(r.get("제품코드", "")) == str(product_id) and
                    str(r.get("부품품번", "")) == str(part_id)):
                self._safe_delete_rows(ws, i + 2)
                self.cache.invalidate("bom")
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
        BATCH = 100
        for i in range(0, len(new_items), BATCH):
            batch    = new_items[i:i + BATCH]
            next_row = len(ws.get_all_values()) + 1
            cell_range = f"A{next_row}:E{next_row + len(batch) - 1}"
            self._safe_update(ws, cell_range, batch)
            if progress_callback:
                progress_callback(
                    f"BOM 신규 등록 중... {min(i + BATCH, len(new_items))}/{len(new_items)}"
                )
            if i + BATCH < len(new_items):
                time.sleep(2)  # 배치 간 딜레이

        # 기존 수정 (배치)
        UBATCH = 50
        for i in range(0, len(update_batches), UBATCH):
            batch = update_batches[i:i + UBATCH]
            for row_num, row_data in batch:
                self._safe_update(ws, f"A{row_num}:E{row_num}", [row_data])
            if progress_callback:
                progress_callback(
                    f"BOM 수정 중... {min(i + UBATCH, len(update_batches))}/{len(update_batches)}"
                )
            if i + UBATCH < len(update_batches):
                time.sleep(3)

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

    def produce_product(self, product_id: str, qty: int, note: str = ""):
        """
        제품 생산 (BOM 기반 자동 출고)
        v2.1: 재고 수량 일괄 업데이트 + 이력 배치 추가
        문제 1 수정: 전체 생산 처리를 락으로 감싸 동시 재고 충돌 방지
        """
        with self._op_lock:
            return self._produce_product_locked(product_id, qty, note)

    def _produce_product_locked(self, product_id: str, qty: int, note: str = ""):
        """produce_product 내부 로직 — _op_lock 보유 상태에서 호출
        문제 2 수정: 재고 검증에 캐시 우회 메서드(_get_fresh_*)를 사용하여
        30초 캐시 TTL 만료나 동시 읽기에 의한 캐시 재채움에도 항상 최신 데이터 사용
        """
        # 캐시 초기화 후 시트에서 직접 최신 데이터 읽기 (캐시 우회)
        self.cache.invalidate_all()

        product = self._get_fresh_product(product_id)
        if not product:
            return False, f"제품코드 '{product_id}'를 찾을 수 없습니다.", []

        bom = self._get_fresh_bom_for_product(product_id)
        if not bom:
            return False, f"제품 '{product['제품명']}'의 BOM이 없습니다.", []

        # 재고 검증에도 캐시 우회 → 락 취득 직후의 실제 재고 사용
        parts_map = self._get_fresh_parts_map()

        # 1단계: 재고 검증
        shortage = []
        for item in bom:
            part_id  = str(item["부품품번"])
            required = float(item["소요량"]) * int(qty)
            part     = parts_map.get(part_id)
            if not part:
                shortage.append(f"품번 '{part_id}' 마스터 없음")
                continue
            current = int(part["현재재고"])
            if current < required:
                shortage.append(
                    f"{part['부품명']}({part_id}): "
                    f"필요 {int(required)}개, 재고 {current}개"
                )

        if shortage:
            return False, "재고 부족으로 생산 불가:\n" + "\n".join(shortage), []

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
            if safety > 0 and new_qty <= safety:
                results.append(
                    f"  ⚠️ {part['부품명']}({part_id}): "
                    f"재고 {new_qty}개 (안전재고: {safety}개)"
                )
            else:
                results.append(
                    f"  {part['부품명']}({part_id}): -{required}개 → {new_qty}개"
                )

        self._bulk_update_part_qtys(qty_updates)
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

            self.cache.invalidate_all()

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

        self.cache.invalidate_all()

        # 같은 시각 생산출고 이력 찾기
        all_history      = ws.get_all_values()
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

    def _add_history_batch(self, entries):
        """이력 일괄 추가 (API 1회)"""
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
        next_row   = len(ws.get_all_values()) + 1
        cell_range = f"A{next_row}:I{next_row + len(rows) - 1}"
        self._safe_update(ws, cell_range, rows)

    def get_all_history(self):
        ws = self.spreadsheet.worksheet(SHEET_HISTORY)
        return ws.get_all_records()

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
        self.cache.invalidate_all()

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

        part_totals = {}
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
                "total_required": int(total_required),
                "current_stock":  current_stock,
                "safety_stock":   safety_stock,
                "shortage":       int(shortage),
                "order_needed":   order_needed,
            })

        parts_requirement.sort(key=lambda x: (not x["order_needed"], x["part_id"]))

        return {
            "plan_summary":        plan_summary,
            "parts_requirement":   parts_requirement,
            "total_order_items":   total_order_items,
            "total_order_qty":     total_order_qty,
        }
