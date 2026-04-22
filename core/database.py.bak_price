"""
구글 시트 데이터베이스 모듈 (v2 - 캐싱 최적화)
- 구글 시트를 DB처럼 사용하여 부품/제품/BOM/이력 데이터를 관리합니다.
- 데이터 캐싱으로 API 호출을 최소화하여 Quota 제한을 방지합니다.
- gspread + google-auth 라이브러리를 사용합니다.
"""

import gspread
from google.oauth2.service_account import Credentials
import os
import json
from datetime import datetime
import time
import threading


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# 시트 이름 상수
SHEET_PARTS = "부품마스터"
SHEET_PRODUCTS = "제품마스터"
SHEET_BOM = "BOM"
SHEET_HISTORY = "입출고이력"

# 캐시 유효 시간 (초)
CACHE_TTL = 30


class DataCache:
    """시트 데이터를 캐싱하여 API 호출을 최소화하는 클래스"""

    def __init__(self):
        self._cache = {}
        self._timestamps = {}
        self._lock = threading.Lock()

    def get(self, key):
        """캐시에서 데이터 조회. 유효기간 내이면 캐시 반환, 아니면 None"""
        with self._lock:
            if key in self._cache:
                elapsed = time.time() - self._timestamps.get(key, 0)
                if elapsed < CACHE_TTL:
                    return self._cache[key]
            return None

    def set(self, key, data):
        """캐시에 데이터 저장"""
        with self._lock:
            self._cache[key] = data
            self._timestamps[key] = time.time()

    def invalidate(self, key=None):
        """캐시 무효화. key=None이면 전체 무효화"""
        with self._lock:
            if key is None:
                self._cache.clear()
                self._timestamps.clear()
            elif key in self._cache:
                del self._cache[key]
                del self._timestamps[key]

    def invalidate_all(self):
        """전체 캐시 무효화"""
        self.invalidate(None)


class GoogleSheetsDB:
    """구글 시트를 데이터베이스처럼 사용하는 클래스 (캐싱 최적화)"""

    def __init__(self, credentials_path: str, spreadsheet_url: str = None):
        self.credentials_path = credentials_path
        self.spreadsheet_url = spreadsheet_url
        self.client = None
        self.spreadsheet = None
        self.cache = DataCache()
        self._connect()

    def _connect(self):
        """구글 시트에 연결"""
        creds = Credentials.from_service_account_file(
            self.credentials_path, scopes=SCOPES
        )
        self.client = gspread.authorize(creds)

        if self.spreadsheet_url:
            self.spreadsheet = self.client.open_by_url(self.spreadsheet_url)
        else:
            self.spreadsheet = self.client.create("재고관리시스템")
        # 항상 시트 초기화 실행 (없는 시트는 자동 생성)
        self._initialize_sheets()

    def _initialize_sheets(self):
        """초기 시트 구조 생성"""
        # 부품마스터 시트
        try:
            ws = self.spreadsheet.worksheet(SHEET_PARTS)
            headers = ws.row_values(1)
            if "업체명" not in headers:
                all_data = ws.get_all_values()
                new_data = [["품번", "부품명", "규격", "단위", "업체명", "현재재고", "안전재고", "비고"]]
                for row in all_data[1:]:
                    while len(row) < 7:
                        row.append('')
                    new_row = row[:4] + [''] + row[4:7]
                    new_data.append(new_row)
                if new_data:
                    ws.update(f"A1:H{len(new_data)}", new_data)
        except gspread.exceptions.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(title=SHEET_PARTS, rows=1000, cols=10)
            ws.update(
                "A1:H1",
                [["품번", "부품명", "규격", "단위", "업체명", "현재재고", "안전재고", "비고"]],
            )

        # 제품마스터 시트
        try:
            ws = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        except gspread.exceptions.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(
                title=SHEET_PRODUCTS, rows=1000, cols=10
            )
            ws.update("A1:E1", [["제품코드", "제품명", "규격", "현재재고", "비고"]])

        # BOM 시트 (단가 컬럼 포함)
        try:
            ws = self.spreadsheet.worksheet(SHEET_BOM)
            headers = ws.row_values(1)
            if "단가" not in headers:
                if len(headers) >= 4 and headers[3] == "비고":
                    ws.update("A1:E1", [["제품코드", "부품품번", "소요량", "단가", "비고"]])
                    records = ws.get_all_values()
                    if len(records) > 1:
                        for i, row in enumerate(records[1:], 2):
                            if len(row) >= 4:
                                old_note = row[3] if len(row) > 3 else ""
                                ws.update(f"D{i}:E{i}", [[0, old_note]])
                                time.sleep(0.3)
                elif len(headers) >= 3:
                    ws.update("A1:E1", [["제품코드", "부품품번", "소요량", "단가", "비고"]])
        except gspread.exceptions.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(title=SHEET_BOM, rows=1000, cols=10)
            ws.update("A1:E1", [["제품코드", "부품품번", "소요량", "단가", "비고"]])

        # 입출고이력 시트
        try:
            ws = self.spreadsheet.worksheet(SHEET_HISTORY)
        except gspread.exceptions.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(
                title=SHEET_HISTORY, rows=10000, cols=15
            )
            ws.update(
                "A1:I1",
                [
                    [
                        "일시",
                        "구분",
                        "유형",
                        "품번/제품코드",
                        "품명",
                        "수량",
                        "잔여재고",
                        "관련제품",
                        "비고",
                    ]
                ],
            )

        # 기본 Sheet1 삭제 시도
        try:
            default_ws = self.spreadsheet.worksheet("Sheet1")
            self.spreadsheet.del_worksheet(default_ws)
        except:
            pass

    def get_spreadsheet_url(self):
        """스프레드시트 URL 반환"""
        return self.spreadsheet.url

    def refresh_cache(self):
        """전체 캐시 강제 갱신"""
        self.cache.invalidate_all()

    # ─────────────────────────────────────────
    # 캐시 기반 데이터 조회 (API 호출 최소화)
    # ─────────────────────────────────────────
    def _get_all_parts_cached(self):
        """부품 전체 조회 (캐시 사용)"""
        cached = self.cache.get("parts")
        if cached is not None:
            return cached
        ws = self.spreadsheet.worksheet(SHEET_PARTS)
        records = ws.get_all_records()
        self.cache.set("parts", records)
        return records

    def _get_all_products_cached(self):
        """제품 전체 조회 (캐시 사용)"""
        cached = self.cache.get("products")
        if cached is not None:
            return cached
        ws = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        records = ws.get_all_records()
        self.cache.set("products", records)
        return records

    def _get_all_bom_cached(self):
        """BOM 전체 조회 (캐시 사용)"""
        cached = self.cache.get("bom")
        if cached is not None:
            return cached
        ws = self.spreadsheet.worksheet(SHEET_BOM)
        records = ws.get_all_records()
        self.cache.set("bom", records)
        return records

    def _get_parts_map(self):
        """부품 맵 (품번 → 부품 dict) 캐시 반환"""
        cached = self.cache.get("parts_map")
        if cached is not None:
            return cached
        parts = self._get_all_parts_cached()
        parts_map = {str(p["품번"]): p for p in parts}
        self.cache.set("parts_map", parts_map)
        return parts_map

    # ─────────────────────────────────────────
    # 부품 관련 메서드
    # ─────────────────────────────────────────
    def get_all_parts(self):
        """모든 부품 목록 조회"""
        return self._get_all_parts_cached()

    def get_part_by_id(self, part_id: str):
        """품번으로 부품 조회 (캐시 사용)"""
        parts_map = self._get_parts_map()
        return parts_map.get(str(part_id), None)

    def add_part(self, part_id, name, spec, unit, qty, safety_qty, note="", supplier=""):
        """부품 추가"""
        ws = self.spreadsheet.worksheet(SHEET_PARTS)
        ws.append_row([str(part_id), name, spec, unit, supplier, int(qty), int(safety_qty), note])
        self.cache.invalidate("parts")
        self.cache.invalidate("parts_map")

    def update_part_qty(self, part_id: str, new_qty: int):
        """부품 재고 수량 업데이트"""
        ws = self.spreadsheet.worksheet(SHEET_PARTS)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if str(r.get("품번", "")) == str(part_id):
                ws.update_cell(i + 2, 6, int(new_qty))
                self.cache.invalidate("parts")
                self.cache.invalidate("parts_map")
                return True
        return False

    def _bulk_update_part_qtys(self, updates):
        """
        부품 재고 수량 일괄 업데이트 (API 호출 최소화)
        updates: [(part_id, new_qty), ...]
        """
        ws = self.spreadsheet.worksheet(SHEET_PARTS)
        all_values = ws.get_all_values()

        # 품번 → 행번호 맵 구성
        row_map = {}
        for i, row in enumerate(all_values[1:], 2):
            if row:
                row_map[str(row[0])] = i

        # 배치 업데이트 구성
        cells_to_update = []
        for part_id, new_qty in updates:
            row_num = row_map.get(str(part_id))
            if row_num:
                cells_to_update.append(gspread.Cell(row_num, 6, int(new_qty)))

        if cells_to_update:
            ws.update_cells(cells_to_update)

        self.cache.invalidate("parts")
        self.cache.invalidate("parts_map")

    def update_part(self, part_id, name, spec, unit, qty, safety_qty, note="", supplier=""):
        """부품 정보 전체 업데이트"""
        ws = self.spreadsheet.worksheet(SHEET_PARTS)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if str(r.get("품번", "")) == str(part_id):
                row = i + 2
                ws.update(f"A{row}:H{row}", [[str(part_id), name, spec, unit, supplier, int(qty), int(safety_qty), note]])
                self.cache.invalidate("parts")
                self.cache.invalidate("parts_map")
                return True
        return False

    def bulk_add_or_update_parts(self, parts_list, progress_callback=None):
        """부품 대량 등록/업데이트 (배치 처리로 API 호출 최소화)"""
        ws = self.spreadsheet.worksheet(SHEET_PARTS)
        existing_records = ws.get_all_records()
        existing_map = {}
        for i, r in enumerate(existing_records):
            existing_map[str(r.get("품번", ""))] = i + 2

        new_parts = []
        update_batches = []
        new_count = 0
        update_count = 0

        for p in parts_list:
            code = str(p["품번"])
            supplier = p.get("업체명", "")
            row_data = [code, p["부품명"], p["규격"], p["단위"], supplier,
                        int(p["현재재고"]), int(p["안전재고"]), p["비고"]]
            if code in existing_map:
                update_batches.append((existing_map[code], row_data))
                update_count += 1
            else:
                new_parts.append(row_data)
                new_count += 1

        # 1) 신규 부품 일괄 추가
        if new_parts:
            batch_size = 100
            for i in range(0, len(new_parts), batch_size):
                batch = new_parts[i:i + batch_size]
                next_row = len(ws.get_all_values()) + 1
                cell_range = f"A{next_row}:H{next_row + len(batch) - 1}"
                ws.update(cell_range, batch)
                if progress_callback:
                    progress_callback(f"신규 등록 중... {min(i + batch_size, len(new_parts))}/{len(new_parts)}")
                if i + batch_size < len(new_parts):
                    time.sleep(2)

        # 2) 기존 부품 업데이트
        if update_batches:
            batch_size = 50
            for i in range(0, len(update_batches), batch_size):
                batch = update_batches[i:i + batch_size]
                for row_num, row_data in batch:
                    ws.update(f"A{row_num}:H{row_num}", [row_data])
                if progress_callback:
                    progress_callback(f"업데이트 중... {min(i + batch_size, len(update_batches))}/{len(update_batches)}")
                if i + batch_size < len(update_batches):
                    time.sleep(3)

        self.cache.invalidate("parts")
        self.cache.invalidate("parts_map")
        return new_count, update_count

    def delete_part(self, part_id: str):
        """부품 삭제"""
        ws = self.spreadsheet.worksheet(SHEET_PARTS)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if str(r.get("품번", "")) == str(part_id):
                ws.delete_rows(i + 2)
                self.cache.invalidate("parts")
                self.cache.invalidate("parts_map")
                return True
        return False

    # ─────────────────────────────────────────
    # 제품 관련 메서드
    # ─────────────────────────────────────────
    def get_all_products(self):
        """모든 제품 목록 조회"""
        return self._get_all_products_cached()

    def get_product_by_id(self, product_id: str):
        """제품코드로 제품 조회 (캐시 사용)"""
        products = self._get_all_products_cached()
        for r in products:
            if str(r.get("제품코드", "")) == str(product_id):
                return r
        return None

    def add_product(self, product_id, name, spec, qty=0, note=""):
        """제품 추가"""
        ws = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        ws.append_row([str(product_id), name, spec, int(qty), note])
        self.cache.invalidate("products")

    def update_product_qty(self, product_id: str, new_qty: int):
        """제품 재고 수량 업데이트"""
        ws = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if str(r.get("제품코드", "")) == str(product_id):
                ws.update_cell(i + 2, 4, int(new_qty))
                self.cache.invalidate("products")
                return True
        return False

    def update_product(self, product_id, name, spec, qty, note=""):
        """제품 정보 전체 업데이트"""
        ws = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if str(r.get("제품코드", "")) == str(product_id):
                row = i + 2
                ws.update(f"A{row}:E{row}", [[str(product_id), name, spec, int(qty), note]])
                self.cache.invalidate("products")
                return True
        return False

    def delete_product(self, product_id: str):
        """제품 삭제"""
        ws = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if str(r.get("제품코드", "")) == str(product_id):
                ws.delete_rows(i + 2)
                self.cache.invalidate("products")
                return True
        return False

    # ─────────────────────────────────────────
    # BOM 관련 메서드
    # ─────────────────────────────────────────
    def get_bom_for_product(self, product_id: str):
        """특정 제품의 BOM(소요량) 조회 (캐시 사용)"""
        all_bom = self._get_all_bom_cached()
        bom_list = []
        for r in all_bom:
            if str(r.get("제품코드", "")) == str(product_id):
                bom_list.append(r)
        return bom_list

    def get_all_bom(self):
        """전체 BOM 조회"""
        return self._get_all_bom_cached()

    def add_bom(self, product_id, part_id, qty, note="", unit_price=0):
        """BOM 항목 추가 (단가 포함)"""
        ws = self.spreadsheet.worksheet(SHEET_BOM)
        ws.append_row([str(product_id), str(part_id), float(qty), float(unit_price), note])
        self.cache.invalidate("bom")

    def update_bom(self, product_id: str, part_id: str, qty: float, note: str = "", unit_price: float = 0):
        """BOM 항목 수정"""
        ws = self.spreadsheet.worksheet(SHEET_BOM)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if str(r.get("제품코드", "")) == str(product_id) and str(r.get("부품품번", "")) == str(part_id):
                row = i + 2
                ws.update(f"A{row}:E{row}", [[str(product_id), str(part_id), float(qty), float(unit_price), note]])
                self.cache.invalidate("bom")
                return True
        return False

    def delete_bom(self, product_id: str, part_id: str):
        """BOM 항목 삭제"""
        ws = self.spreadsheet.worksheet(SHEET_BOM)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if str(r.get("제품코드", "")) == str(product_id) and str(
                r.get("부품품번", "")
            ) == str(part_id):
                ws.delete_rows(i + 2)
                self.cache.invalidate("bom")
                return True
        return False

    def delete_all_bom_for_product(self, product_id: str):
        """특정 제품의 모든 BOM 삭제"""
        ws = self.spreadsheet.worksheet(SHEET_BOM)
        records = ws.get_all_records()
        rows_to_delete = []
        for i, r in enumerate(records):
            if str(r.get("제품코드", "")) == str(product_id):
                rows_to_delete.append(i + 2)
        for row in sorted(rows_to_delete, reverse=True):
            ws.delete_rows(row)
            time.sleep(0.5)
        self.cache.invalidate("bom")

    def bulk_add_or_update_bom(self, bom_list, progress_callback=None):
        """
        BOM 대량 등록/업데이트 (배치 처리)
        bom_list: [{"제품코드": ..., "부품품번": ..., "소요량": ..., "단가": ..., "비고": ...}, ...]
        """
        ws = self.spreadsheet.worksheet(SHEET_BOM)
        existing_records = ws.get_all_records()

        # 기존 BOM 맵: (제품코드, 부품품번) → 행번호
        existing_map = {}
        for i, r in enumerate(existing_records):
            key = (str(r.get("제품코드", "")), str(r.get("부품품번", "")))
            existing_map[key] = i + 2

        new_items = []
        update_batches = []
        new_count = 0
        update_count = 0

        for item in bom_list:
            prod_code = str(item["제품코드"])
            part_code = str(item["부품품번"])
            qty = float(item["소요량"])
            unit_price = float(item.get("단가", 0))
            note = str(item.get("비고", ""))
            row_data = [prod_code, part_code, qty, unit_price, note]

            key = (prod_code, part_code)
            if key in existing_map:
                update_batches.append((existing_map[key], row_data))
                update_count += 1
            else:
                new_items.append(row_data)
                new_count += 1

        # 1) 신규 BOM 일괄 추가
        if new_items:
            batch_size = 100
            for i in range(0, len(new_items), batch_size):
                batch = new_items[i:i + batch_size]
                next_row = len(ws.get_all_values()) + 1
                cell_range = f"A{next_row}:E{next_row + len(batch) - 1}"
                ws.update(cell_range, batch)
                if progress_callback:
                    progress_callback(f"BOM 신규 등록 중... {min(i + batch_size, len(new_items))}/{len(new_items)}")
                if i + batch_size < len(new_items):
                    time.sleep(2)

        # 2) 기존 BOM 업데이트
        if update_batches:
            batch_size = 50
            for i in range(0, len(update_batches), batch_size):
                batch = update_batches[i:i + batch_size]
                for row_num, row_data in batch:
                    ws.update(f"A{row_num}:E{row_num}", [row_data])
                if progress_callback:
                    progress_callback(f"BOM 업데이트 중... {min(i + batch_size, len(update_batches))}/{len(update_batches)}")
                if i + batch_size < len(update_batches):
                    time.sleep(3)

        self.cache.invalidate("bom")
        return new_count, update_count

    def get_product_cost(self, product_id: str):
        """
        제품별 원가 계산 (BOM 기반, 캐시 사용)
        반환: (총원가, [(부품품번, 부품명, 소요량, 단가, 소계), ...])
        """
        bom = self.get_bom_for_product(product_id)
        parts_map = self._get_parts_map()

        total_cost = 0
        cost_details = []

        for item in bom:
            part_id = str(item.get("부품품번", ""))
            qty = float(item.get("소요량", 0))
            unit_price = float(item.get("단가", 0))
            subtotal = qty * unit_price

            part_name = parts_map.get(part_id, {}).get("부품명", "?")

            cost_details.append({
                "부품품번": part_id,
                "부품명": part_name,
                "소요량": qty,
                "단가": unit_price,
                "소계": subtotal,
            })
            total_cost += subtotal

        return total_cost, cost_details

    # ─────────────────────────────────────────
    # 입출고 처리 메서드 (캐시 최적화)
    # ─────────────────────────────────────────
    def receive_part(self, part_id: str, qty: int, note: str = ""):
        """부품 입고 처리"""
        # 캐시 무효화 후 최신 데이터로 처리
        self.cache.invalidate("parts")
        self.cache.invalidate("parts_map")
        part = self.get_part_by_id(part_id)
        if not part:
            return False, f"품번 '{part_id}'를 찾을 수 없습니다."

        new_qty = int(part["현재재고"]) + int(qty)
        self.update_part_qty(part_id, new_qty)

        self._add_history("입고", "부품입고", part_id, part["부품명"], qty, new_qty, "", note)
        return True, f"입고 완료: {part['부품명']} +{qty}개 (현재재고: {new_qty}개)"

    def issue_part(self, part_id: str, qty: int, note: str = ""):
        """부품 개별 출고 처리"""
        # 캐시 무효화 후 최신 데이터로 처리
        self.cache.invalidate("parts")
        self.cache.invalidate("parts_map")
        part = self.get_part_by_id(part_id)
        if not part:
            return False, f"품번 '{part_id}'를 찾을 수 없습니다."

        current = int(part["현재재고"])
        if current < int(qty):
            return False, f"재고 부족: {part['부품명']} 현재재고 {current}개, 출고요청 {qty}개"

        new_qty = current - int(qty)
        self.update_part_qty(part_id, new_qty)

        self._add_history("출고", "개별출고", part_id, part["부품명"], qty, new_qty, "", note)

        warning = ""
        safety = int(part.get("안전재고", 0))
        if safety > 0 and new_qty <= safety:
            warning = f" ⚠️ 안전재고({safety}개) 이하입니다!"

        return True, f"출고 완료: {part['부품명']} -{qty}개 (현재재고: {new_qty}개){warning}"

    def produce_product(self, product_id: str, qty: int, note: str = ""):
        """
        제품 생산 처리 (BOM 기반 부품 자동 출고) - 캐시 최적화 버전
        - 전체 부품 데이터를 1회만 조회하여 API 호출 최소화
        - 부품 재고 업데이트를 일괄 처리
        """
        # 캐시 무효화 후 최신 데이터 1회 조회
        self.cache.invalidate_all()

        product = self.get_product_by_id(product_id)
        if not product:
            return False, f"제품코드 '{product_id}'를 찾을 수 없습니다.", []

        bom = self.get_bom_for_product(product_id)
        if not bom:
            return False, f"제품 '{product['제품명']}'의 BOM(소요량) 정보가 없습니다.", []

        # 부품 전체를 1회 조회하여 맵 구성 (API 1회)
        parts_map = self._get_parts_map()

        # 1단계: 재고 충분한지 사전 확인 (추가 API 호출 없음)
        warnings = []
        shortage = []
        for item in bom:
            part_id = str(item["부품품번"])
            required = float(item["소요량"]) * int(qty)
            part = parts_map.get(part_id)
            if not part:
                shortage.append(f"품번 '{part_id}' 미등록")
                continue
            current = int(part["현재재고"])
            if current < required:
                shortage.append(
                    f"{part['부품명']}({part_id}): 필요 {int(required)}개, 현재 {current}개"
                )

        if shortage:
            return False, "재고 부족으로 생산 불가:\n" + "\n".join(shortage), []

        # 2단계: 부품 출고 처리 (일괄 업데이트)
        results = []
        qty_updates = []  # (part_id, new_qty) 리스트
        history_entries = []  # 이력 엔트리 리스트

        for item in bom:
            part_id = str(item["부품품번"])
            required = int(float(item["소요량"]) * int(qty))
            part = parts_map.get(part_id)
            current = int(part["현재재고"])
            new_qty = current - required

            qty_updates.append((part_id, new_qty))
            history_entries.append({
                "direction": "출고",
                "h_type": "생산출고",
                "item_id": part_id,
                "item_name": part["부품명"],
                "qty": required,
                "remaining": new_qty,
                "related": f"{product['제품명']}({product_id})",
                "note": note,
            })

            safety = int(part.get("안전재고", 0))
            if safety > 0 and new_qty <= safety:
                warnings.append(
                    f"⚠️ {part['부품명']}({part_id}): 현재 {new_qty}개 (안전재고: {safety}개)"
                )

            results.append(f"  {part['부품명']}({part_id}): -{required}개 → 잔여 {new_qty}개")

        # 일괄 재고 업데이트 (API 1회)
        self._bulk_update_part_qtys(qty_updates)

        # 이력 일괄 추가
        self._add_history_batch(history_entries)

        # 3단계: 제품 재고 증가
        product_new_qty = int(product["현재재고"]) + int(qty)
        self.update_product_qty(product_id, product_new_qty)

        self._add_history(
            "입고", "생산입고", product_id, product["제품명"], qty, product_new_qty, "", note
        )

        msg = f"생산 완료: {product['제품명']} +{qty}개 (제품재고: {product_new_qty}개)\n"
        msg += "출고된 부품:\n" + "\n".join(results)

        return True, msg, warnings

    def cancel_history(self, row_index: int):
        """
        이력 취소 처리 (입출고 원복)
        row_index: 이력 시트의 행 번호 (2부터 시작)
        """
        ws = self.spreadsheet.worksheet(SHEET_HISTORY)
        row_data = ws.row_values(row_index)

        if len(row_data) < 7:
            return False, "유효하지 않은 이력입니다."

        direction = row_data[1]
        h_type = row_data[2]
        item_id = row_data[3]
        item_name = row_data[4]
        h_qty = int(row_data[5])

        # 캐시 무효화 후 최신 데이터로 처리
        self.cache.invalidate_all()

        if h_type in ["부품입고"]:
            part = self.get_part_by_id(item_id)
            if part:
                new_qty = int(part["현재재고"]) - h_qty
                if new_qty < 0:
                    return False, f"취소 시 재고가 음수({new_qty})가 됩니다."
                self.update_part_qty(item_id, new_qty)
                self._add_history("취소", f"{h_type}취소", item_id, item_name, h_qty, new_qty, "", "이력 취소")
                return True, f"입고 취소: {item_name} -{h_qty}개 (현재재고: {new_qty}개)"

        elif h_type in ["개별출고", "생산출고"]:
            part = self.get_part_by_id(item_id)
            if part:
                new_qty = int(part["현재재고"]) + h_qty
                self.update_part_qty(item_id, new_qty)
                self._add_history("취소", f"{h_type}취소", item_id, item_name, h_qty, new_qty, "", "이력 취소")
                return True, f"출고 취소: {item_name} +{h_qty}개 (현재재고: {new_qty}개)"

        elif h_type == "생산입고":
            # 생산입고 취소는 항상 부품 재고까지 함께 원복하는 cancel_production으로 처리
            success, msg, details = self.cancel_production(row_index)
            return success, msg

        return False, "취소할 수 없는 이력 유형입니다."

    def cancel_production(self, production_entry_row: int):
        """
        생산 일괄 취소 처리
        - '생산입고' 이력을 기준으로 같은 시각에 발생한 관련 '생산출고' 이력을 모두 찾아
          제품 재고 원복 + 소요 부품 재고 일괄 원복을 한 번에 처리합니다.
        - production_entry_row: '생산입고' 이력의 행 번호 (2부터 시작)
        """
        ws = self.spreadsheet.worksheet(SHEET_HISTORY)
        row_data = ws.row_values(production_entry_row)

        if len(row_data) < 7:
            return False, "유효하지 않은 이력입니다.", []

        h_type = row_data[2]
        if h_type != "생산입고":
            return False, "생산입고 이력만 일괄 취소할 수 있습니다.\n선택한 이력의 유형: " + h_type, []

        prod_time = row_data[0]       # 생산 일시
        product_id = row_data[3]      # 제품코드
        product_name = row_data[4]    # 제품명
        prod_qty = int(row_data[5])   # 생산 수량

        # 캐시 무효화 후 최신 데이터로 처리
        self.cache.invalidate_all()

        # 1단계: 같은 시각(±10초) + 같은 관련제품의 '생산출고' 이력 모두 찾기
        all_history = ws.get_all_values()
        related_keyword = f"{product_name}({product_id})"

        # 시간 비교를 위한 datetime 변환
        try:
            prod_dt = datetime.strptime(prod_time, "%Y-%m-%d %H:%M:%S")
        except:
            prod_dt = None

        issue_rows = []  # (행번호, 품번, 품명, 출고수량)
        for i, row in enumerate(all_history[1:], 2):  # 헤더 제외
            if len(row) < 8:
                continue
            
            h_time_str = row[0]
            h_type = row[2]
            h_related = row[7]

            # 1. 유형 확인
            if h_type != "생산출고":
                continue
            
            # 2. 관련제품 키워드 확인 (정확히 일치하거나 포함되는지)
            if related_keyword not in h_related and h_related not in related_keyword:
                continue

            # 3. 시간 확인 (정확히 일치하거나 ±10초 이내)
            is_time_match = (h_time_str == prod_time)
            if not is_time_match and prod_dt:
                try:
                    h_dt = datetime.strptime(h_time_str, "%Y-%m-%d %H:%M:%S")
                    if abs((h_dt - prod_dt).total_seconds()) <= 10:
                        is_time_match = True
                except:
                    pass
            
            if is_time_match:
                issue_rows.append({
                    "row": i,
                    "part_id": row[3],
                    "part_name": row[4],
                    "qty": int(row[5]),
                })

        # 2단계: 제품 재고 원복 (생산입고 취소)
        product = self.get_product_by_id(product_id)
        if not product:
            return False, f"제품코드 '{product_id}'를 찾을 수 없습니다.", []

        product_current = int(product["현재재고"])
        product_new_qty = product_current - prod_qty
        if product_new_qty < 0:
            return False, (f"취소 시 제품 재고가 음수({product_new_qty})가 됩니다.\n"
                           f"현재 제품 재고: {product_current}개, 취소 수량: {prod_qty}개"), []

        # 3단계: 부품 재고 원복 준비 (일괄 처리)
        parts_map = self._get_parts_map()
        qty_updates = []  # (part_id, new_qty)
        results = []

        for ir in issue_rows:
            part = parts_map.get(ir["part_id"])
            if part:
                part_current = int(part["현재재고"])
                part_new_qty = part_current + ir["qty"]
                qty_updates.append((ir["part_id"], part_new_qty))
                results.append(
                    f"  {ir['part_name']}({ir['part_id']}): +{ir['qty']}개 → 재고 {part_new_qty}개"
                )
            else:
                results.append(
                    f"  {ir['part_name']}({ir['part_id']}): 부품 미등록 (원복 불가)"
                )

        # 4단계: 실제 업데이트 실행
        # 4-1) 부품 재고 일괄 원복 (API 1회)
        if qty_updates:
            self._bulk_update_part_qtys(qty_updates)

        # 4-2) 제품 재고 원복
        self.update_product_qty(product_id, product_new_qty)

        # 5단계: 취소 이력 일괄 기록
        cancel_entries = []

        # 부품 출고 취소 이력
        for ir in issue_rows:
            part = parts_map.get(ir["part_id"])
            part_restored = int(part["현재재고"]) + ir["qty"] if part else 0
            cancel_entries.append({
                "direction": "취소",
                "h_type": "생산출고취소",
                "item_id": ir["part_id"],
                "item_name": ir["part_name"],
                "qty": ir["qty"],
                "remaining": part_restored,
                "related": related_keyword,
                "note": "생산 일괄 취소",
            })

        # 이력 일괄 추가 (API 1회)
        if cancel_entries:
            self._add_history_batch(cancel_entries)

        # 제품 생산입고 취소 이력 (별도 1건)
        self._add_history(
            "취소", "생산입고취소", product_id, product_name,
            prod_qty, product_new_qty, "", "생산 일괄 취소"
        )

        msg = (f"생산 일괄 취소 완료: {product_name} -{prod_qty}개 "
               f"(제품재고: {product_current}개 → {product_new_qty}개)\n\n"
               f"원복된 부품 ({len(issue_rows)}종):\n" + "\n".join(results))

        return True, msg, results

    def _add_history(self, direction, h_type, item_id, item_name, qty, remaining, related="", note=""):
        """이력 추가"""
        ws = self.spreadsheet.worksheet(SHEET_HISTORY)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ws.append_row([now, direction, h_type, str(item_id), item_name, int(qty), int(remaining), related, note])

    def _add_history_batch(self, entries):
        """이력 일괄 추가 (API 호출 최소화)"""
        if not entries:
            return
        ws = self.spreadsheet.worksheet(SHEET_HISTORY)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        rows = []
        for e in entries:
            rows.append([
                now, e["direction"], e["h_type"], str(e["item_id"]),
                e["item_name"], int(e["qty"]), int(e["remaining"]),
                e.get("related", ""), e.get("note", "")
            ])

        # 일괄 추가
        next_row = len(ws.get_all_values()) + 1
        cell_range = f"A{next_row}:I{next_row + len(rows) - 1}"
        ws.update(cell_range, rows)

    # ─────────────────────────────────────────
    # 이력 조회 메서드
    # ─────────────────────────────────────────
    def get_all_history(self):
        """전체 이력 조회"""
        ws = self.spreadsheet.worksheet(SHEET_HISTORY)
        return ws.get_all_records()

    def get_history_by_date_range(self, start_date: str, end_date: str):
        """기간별 이력 조회 (YYYY-MM-DD 형식)"""
        records = self.get_all_history()
        filtered = []
        for r in records:
            date_str = str(r.get("일시", ""))[:10]
            if start_date <= date_str <= end_date:
                filtered.append(r)
        return filtered

    # ─────────────────────────────────────────
    # 안전재고 경고 조회
    # ─────────────────────────────────────────
    def get_safety_stock_alerts(self):
        """안전재고 이하 부품 목록 조회 (캐시 사용)"""
        parts = self._get_all_parts_cached()
        alerts = []
        for p in parts:
            safety = int(p.get("안전재고", 0))
            current = int(p.get("현재재고", 0))
            if safety > 0 and current <= safety:
                alerts.append(p)
        return alerts

    # ─────────────────────────────────────────
    # MRP (자재소요계획) 관련 메서드
    # ─────────────────────────────────────────
    def get_max_producible(self, product_id: str):
        """
        현재 부품 재고 기준 최대 생산 가능 수량 계산
        반환: (최대수량, 병목부품명) 또는 (0, 사유)
        """
        bom = self.get_bom_for_product(product_id)
        if not bom:
            return 0, "BOM 미등록"

        parts_map = self._get_parts_map()
        min_qty = float('inf')
        bottleneck = ""

        for item in bom:
            part_id = str(item.get("부품품번", ""))
            required = float(item.get("소요량", 0))
            if required <= 0:
                continue

            part = parts_map.get(part_id)
            if not part:
                return 0, f"부품 미등록: {part_id}"

            current_stock = int(part.get("현재재고", 0))
            possible = int(current_stock / required)

            if possible < min_qty:
                min_qty = possible
                bottleneck = part.get("부품명", part_id)

        if min_qty == float('inf'):
            return 0, "소요량 없음"

        return min_qty, bottleneck

    def calculate_mrp(self, production_plan, include_safety_stock=False):
        """
        MRP(자재소요계획) 계산

        production_plan: [{"product_id": str, "product_name": str, "target_qty": int}, ...]
        include_safety_stock: 안전재고 반영 여부

        반환: {
            "plan_summary": [{"product_id", "product_name", "current_stock", "target_qty", "need_to_produce", "max_producible", "bottleneck"}, ...],
            "parts_requirement": [{"part_id", "part_name", "supplier", "unit", "total_required", "current_stock", "safety_stock", "shortage", "order_needed"}, ...],
            "total_order_items": int,
            "total_order_qty": int,
        }
        """
        # 캐시 갱신
        self.cache.invalidate_all()

        products = self._get_all_products_cached()
        products_map = {str(p["제품코드"]): p for p in products}
        parts_map = self._get_parts_map()
        all_bom = self._get_all_bom_cached()

        # 1. 생산 계획 요약 계산
        plan_summary = []
        for plan in production_plan:
            pid = str(plan["product_id"])
            product = products_map.get(pid)
            current_stock = int(product["현재재고"]) if product else 0
            target_qty = int(plan["target_qty"])
            need_to_produce = max(0, target_qty - current_stock)

            max_prod, bottleneck = self.get_max_producible(pid)

            plan_summary.append({
                "product_id": pid,
                "product_name": plan.get("product_name", product["제품명"] if product else pid),
                "current_stock": current_stock,
                "target_qty": target_qty,
                "need_to_produce": need_to_produce,
                "max_producible": max_prod,
                "bottleneck": bottleneck,
            })

        # 2. 부품별 총 소요량 집계
        part_totals = {}  # part_id → 총소요량
        for plan_item in plan_summary:
            pid = plan_item["product_id"]
            need = plan_item["need_to_produce"]
            if need <= 0:
                continue

            bom_items = [b for b in all_bom if str(b.get("제품코드", "")) == pid]
            for bom_item in bom_items:
                part_id = str(bom_item.get("부품품번", ""))
                required_per_unit = float(bom_item.get("소요량", 0))
                total_for_this = required_per_unit * need

                if part_id in part_totals:
                    part_totals[part_id] += total_for_this
                else:
                    part_totals[part_id] = total_for_this

        # 3. 부품별 발주 필요 수량 계산
        parts_requirement = []
        total_order_items = 0
        total_order_qty = 0

        for part_id, total_required in sorted(part_totals.items()):
            part = parts_map.get(part_id)
            if not part:
                continue

            current_stock = int(part.get("현재재고", 0))
            safety_stock = int(part.get("안전재고", 0))
            part_name = part.get("부품명", "")
            supplier = part.get("업체명", "")
            unit = part.get("단위", "")

            total_required_int = int(total_required) if total_required == int(total_required) else total_required

            if include_safety_stock:
                # 발주 필요 = 총소요량 - 현재재고 + 안전재고
                shortage = total_required - current_stock + safety_stock
            else:
                # 발주 필요 = 총소요량 - 현재재고
                shortage = total_required - current_stock

            shortage = max(0, int(shortage) if shortage == int(shortage) else shortage)
            order_needed = shortage > 0

            if order_needed:
                total_order_items += 1
                total_order_qty += int(shortage)

            parts_requirement.append({
                "part_id": part_id,
                "part_name": part_name,
                "supplier": supplier,
                "unit": unit,
                "total_required": total_required_int,
                "current_stock": current_stock,
                "safety_stock": safety_stock,
                "shortage": shortage,
                "order_needed": order_needed,
            })

        # 발주 필요한 부품을 상단에 정렬
        parts_requirement.sort(key=lambda x: (not x["order_needed"], x["part_id"]))

        return {
            "plan_summary": plan_summary,
            "parts_requirement": parts_requirement,
            "total_order_items": total_order_items,
            "total_order_qty": total_order_qty,
        }
