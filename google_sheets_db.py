"""
구글 시트 데이터베이스 모듈
- 구글 시트를 DB처럼 사용하여 부품/제품/BOM/이력 데이터를 관리합니다.
- gspread + google-auth 라이브러리를 사용합니다.
"""

import gspread
from google.oauth2.service_account import Credentials
import os
import json
from datetime import datetime


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# 시트 이름 상수
SHEET_PARTS = "부품마스터"
SHEET_PRODUCTS = "제품마스터"
SHEET_BOM = "BOM"
SHEET_HISTORY = "입출고이력"


class GoogleSheetsDB:
    """구글 시트를 데이터베이스처럼 사용하는 클래스"""

    def __init__(self, credentials_path: str, spreadsheet_url: str = None):
        self.credentials_path = credentials_path
        self.spreadsheet_url = spreadsheet_url
        self.client = None
        self.spreadsheet = None
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
            # 새 스프레드시트 생성
            self.spreadsheet = self.client.create("재고관리시스템")
        # 항상 시트 초기화 실행 (없는 시트는 자동 생성)
        self._initialize_sheets()

    def _initialize_sheets(self):
        """초기 시트 구조 생성"""
        # 부품마스터 시트
        try:
            ws = self.spreadsheet.worksheet(SHEET_PARTS)
        except gspread.exceptions.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(title=SHEET_PARTS, rows=1000, cols=10)
            ws.update(
                "A1:G1",
                [["품번", "부품명", "규격", "단위", "현재재고", "안전재고", "비고"]],
            )

        # 제품마스터 시트
        try:
            ws = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        except gspread.exceptions.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(
                title=SHEET_PRODUCTS, rows=1000, cols=10
            )
            ws.update("A1:E1", [["제품코드", "제품명", "규격", "현재재고", "비고"]])

        # BOM 시트
        try:
            ws = self.spreadsheet.worksheet(SHEET_BOM)
        except gspread.exceptions.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(title=SHEET_BOM, rows=1000, cols=10)
            ws.update("A1:D1", [["제품코드", "부품품번", "소요량", "비고"]])

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

    # ─────────────────────────────────────────
    # 부품 관련 메서드
    # ─────────────────────────────────────────
    def get_all_parts(self):
        """모든 부품 목록 조회"""
        ws = self.spreadsheet.worksheet(SHEET_PARTS)
        records = ws.get_all_records()
        return records

    def get_part_by_id(self, part_id: str):
        """품번으로 부품 조회"""
        ws = self.spreadsheet.worksheet(SHEET_PARTS)
        records = ws.get_all_records()
        for r in records:
            if str(r.get("품번", "")) == str(part_id):
                return r
        return None

    def add_part(self, part_id, name, spec, unit, qty, safety_qty, note=""):
        """부품 추가"""
        ws = self.spreadsheet.worksheet(SHEET_PARTS)
        ws.append_row([str(part_id), name, spec, unit, int(qty), int(safety_qty), note])

    def update_part_qty(self, part_id: str, new_qty: int):
        """부품 재고 수량 업데이트"""
        ws = self.spreadsheet.worksheet(SHEET_PARTS)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if str(r.get("품번", "")) == str(part_id):
                ws.update_cell(i + 2, 5, int(new_qty))  # E열 = 현재재고
                return True
        return False

    def update_part(self, part_id, name, spec, unit, qty, safety_qty, note=""):
        """부품 정보 전체 업데이트"""
        ws = self.spreadsheet.worksheet(SHEET_PARTS)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if str(r.get("품번", "")) == str(part_id):
                row = i + 2
                ws.update(f"A{row}:G{row}", [[str(part_id), name, spec, unit, int(qty), int(safety_qty), note]])
                return True
        return False

    def bulk_add_or_update_parts(self, parts_list, progress_callback=None):
        """부품 대량 등록/업데이트 (배치 처리로 API 호출 최소화)"""
        import time
        ws = self.spreadsheet.worksheet(SHEET_PARTS)
        existing_records = ws.get_all_records()
        existing_map = {}
        for i, r in enumerate(existing_records):
            existing_map[str(r.get("품번", ""))] = i + 2  # 행 번호 (1-indexed header + offset)

        new_parts = []
        update_batches = []  # (row_num, data) 쌍
        new_count = 0
        update_count = 0

        for p in parts_list:
            code = str(p["품번"])
            row_data = [code, p["부품명"], p["규격"], p["단위"],
                        int(p["현재재고"]), int(p["안전재고"]), p["비고"]]
            if code in existing_map:
                update_batches.append((existing_map[code], row_data))
                update_count += 1
            else:
                new_parts.append(row_data)
                new_count += 1

        # 1) 신규 부품 일괄 추가 (배치)
        if new_parts:
            batch_size = 100
            for i in range(0, len(new_parts), batch_size):
                batch = new_parts[i:i + batch_size]
                # 현재 마지막 행 이후에 추가
                next_row = len(ws.get_all_values()) + 1
                cell_range = f"A{next_row}:G{next_row + len(batch) - 1}"
                ws.update(cell_range, batch)
                if progress_callback:
                    progress_callback(f"신규 등록 중... {min(i + batch_size, len(new_parts))}/{len(new_parts)}")
                if i + batch_size < len(new_parts):
                    time.sleep(2)

        # 2) 기존 부품 업데이트 (배치)
        if update_batches:
            batch_size = 50
            for i in range(0, len(update_batches), batch_size):
                batch = update_batches[i:i + batch_size]
                for row_num, row_data in batch:
                    ws.update(f"A{row_num}:G{row_num}", [row_data])
                if progress_callback:
                    progress_callback(f"업데이트 중... {min(i + batch_size, len(update_batches))}/{len(update_batches)}")
                # API 제한 방지: 배치마다 대기
                if i + batch_size < len(update_batches):
                    time.sleep(3)

        return new_count, update_count

    def delete_part(self, part_id: str):
        """부품 삭제"""
        ws = self.spreadsheet.worksheet(SHEET_PARTS)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if str(r.get("품번", "")) == str(part_id):
                ws.delete_rows(i + 2)
                return True
        return False

    # ─────────────────────────────────────────
    # 제품 관련 메서드
    # ─────────────────────────────────────────
    def get_all_products(self):
        """모든 제품 목록 조회"""
        ws = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        records = ws.get_all_records()
        return records

    def get_product_by_id(self, product_id: str):
        """제품코드로 제품 조회"""
        ws = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        records = ws.get_all_records()
        for r in records:
            if str(r.get("제품코드", "")) == str(product_id):
                return r
        return None

    def add_product(self, product_id, name, spec, qty=0, note=""):
        """제품 추가"""
        ws = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        ws.append_row([str(product_id), name, spec, int(qty), note])

    def update_product_qty(self, product_id: str, new_qty: int):
        """제품 재고 수량 업데이트"""
        ws = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if str(r.get("제품코드", "")) == str(product_id):
                ws.update_cell(i + 2, 4, int(new_qty))  # D열 = 현재재고
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
                return True
        return False

    def delete_product(self, product_id: str):
        """제품 삭제"""
        ws = self.spreadsheet.worksheet(SHEET_PRODUCTS)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if str(r.get("제품코드", "")) == str(product_id):
                ws.delete_rows(i + 2)
                return True
        return False

    # ─────────────────────────────────────────
    # BOM 관련 메서드
    # ─────────────────────────────────────────
    def get_bom_for_product(self, product_id: str):
        """특정 제품의 BOM(소요량) 조회"""
        ws = self.spreadsheet.worksheet(SHEET_BOM)
        records = ws.get_all_records()
        bom_list = []
        for r in records:
            if str(r.get("제품코드", "")) == str(product_id):
                bom_list.append(r)
        return bom_list

    def get_all_bom(self):
        """전체 BOM 조회"""
        ws = self.spreadsheet.worksheet(SHEET_BOM)
        return ws.get_all_records()

    def add_bom(self, product_id, part_id, qty, note=""):
        """BOM 항목 추가"""
        ws = self.spreadsheet.worksheet(SHEET_BOM)
        ws.append_row([str(product_id), str(part_id), float(qty), note])

    def delete_bom(self, product_id: str, part_id: str):
        """BOM 항목 삭제"""
        ws = self.spreadsheet.worksheet(SHEET_BOM)
        records = ws.get_all_records()
        for i, r in enumerate(records):
            if str(r.get("제품코드", "")) == str(product_id) and str(
                r.get("부품품번", "")
            ) == str(part_id):
                ws.delete_rows(i + 2)
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
        # 역순으로 삭제 (인덱스 밀림 방지)
        for row in sorted(rows_to_delete, reverse=True):
            ws.delete_rows(row)

    # ─────────────────────────────────────────
    # 입출고 처리 메서드
    # ─────────────────────────────────────────
    def receive_part(self, part_id: str, qty: int, note: str = ""):
        """부품 입고 처리"""
        part = self.get_part_by_id(part_id)
        if not part:
            return False, f"품번 '{part_id}'를 찾을 수 없습니다."

        new_qty = int(part["현재재고"]) + int(qty)
        self.update_part_qty(part_id, new_qty)

        # 이력 기록
        self._add_history("입고", "부품입고", part_id, part["부품명"], qty, new_qty, "", note)
        return True, f"입고 완료: {part['부품명']} +{qty}개 (현재재고: {new_qty}개)"

    def issue_part(self, part_id: str, qty: int, note: str = ""):
        """부품 개별 출고 처리"""
        part = self.get_part_by_id(part_id)
        if not part:
            return False, f"품번 '{part_id}'를 찾을 수 없습니다."

        current = int(part["현재재고"])
        if current < int(qty):
            return False, f"재고 부족: {part['부품명']} 현재재고 {current}개, 출고요청 {qty}개"

        new_qty = current - int(qty)
        self.update_part_qty(part_id, new_qty)

        # 이력 기록
        self._add_history("출고", "개별출고", part_id, part["부품명"], qty, new_qty, "", note)

        # 안전재고 경고 확인
        warning = ""
        safety = int(part.get("안전재고", 0))
        if safety > 0 and new_qty <= safety:
            warning = f" ⚠️ 안전재고({safety}개) 이하입니다!"

        return True, f"출고 완료: {part['부품명']} -{qty}개 (현재재고: {new_qty}개){warning}"

    def produce_product(self, product_id: str, qty: int, note: str = ""):
        """
        제품 생산 처리 (BOM 기반 부품 자동 출고)
        - 제품 재고 증가
        - BOM에 따라 부품 재고 자동 감소
        """
        product = self.get_product_by_id(product_id)
        if not product:
            return False, f"제품코드 '{product_id}'를 찾을 수 없습니다.", []

        bom = self.get_bom_for_product(product_id)
        if not bom:
            return False, f"제품 '{product['제품명']}'의 BOM(소요량) 정보가 없습니다.", []

        # 1단계: 재고 충분한지 사전 확인
        warnings = []
        shortage = []
        for item in bom:
            part_id = str(item["부품품번"])
            required = float(item["소요량"]) * int(qty)
            part = self.get_part_by_id(part_id)
            if not part:
                shortage.append(f"품번 '{part_id}' 미등록")
                continue
            current = int(part["현재재고"])
            if current < required:
                shortage.append(
                    f"{part['부품명']}({part_id}): 필요 {required}개, 현재 {current}개"
                )

        if shortage:
            return False, "재고 부족으로 생산 불가:\n" + "\n".join(shortage), []

        # 2단계: 부품 출고 처리
        results = []
        for item in bom:
            part_id = str(item["부품품번"])
            required = int(float(item["소요량"]) * int(qty))
            part = self.get_part_by_id(part_id)
            current = int(part["현재재고"])
            new_qty = current - required
            self.update_part_qty(part_id, new_qty)

            # 이력 기록
            self._add_history(
                "출고",
                "생산출고",
                part_id,
                part["부품명"],
                required,
                new_qty,
                f"{product['제품명']}({product_id})",
                note,
            )

            # 안전재고 경고
            safety = int(part.get("안전재고", 0))
            if safety > 0 and new_qty <= safety:
                warnings.append(
                    f"⚠️ {part['부품명']}({part_id}): 현재 {new_qty}개 (안전재고: {safety}개)"
                )

            results.append(f"  {part['부품명']}({part_id}): -{required}개 → 잔여 {new_qty}개")

        # 3단계: 제품 재고 증가
        product_new_qty = int(product["현재재고"]) + int(qty)
        self.update_product_qty(product_id, product_new_qty)

        # 제품 입고 이력
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

        # 일시, 구분, 유형, 품번/제품코드, 품명, 수량, 잔여재고, 관련제품, 비고
        direction = row_data[1]  # 입고/출고
        h_type = row_data[2]  # 부품입고/개별출고/생산출고/생산입고
        item_id = row_data[3]
        item_name = row_data[4]
        h_qty = int(row_data[5])

        if h_type in ["부품입고"]:
            # 입고 취소 → 재고 감소
            part = self.get_part_by_id(item_id)
            if part:
                new_qty = int(part["현재재고"]) - h_qty
                if new_qty < 0:
                    return False, f"취소 시 재고가 음수({new_qty})가 됩니다."
                self.update_part_qty(item_id, new_qty)
                self._add_history("취소", f"{h_type}취소", item_id, item_name, h_qty, new_qty, "", "이력 취소")
                return True, f"입고 취소: {item_name} -{h_qty}개 (현재재고: {new_qty}개)"

        elif h_type in ["개별출고", "생산출고"]:
            # 출고 취소 → 재고 증가
            part = self.get_part_by_id(item_id)
            if part:
                new_qty = int(part["현재재고"]) + h_qty
                self.update_part_qty(item_id, new_qty)
                self._add_history("취소", f"{h_type}취소", item_id, item_name, h_qty, new_qty, "", "이력 취소")
                return True, f"출고 취소: {item_name} +{h_qty}개 (현재재고: {new_qty}개)"

        elif h_type == "생산입고":
            # 제품 생산입고 취소 → 제품 재고 감소
            product = self.get_product_by_id(item_id)
            if product:
                new_qty = int(product["현재재고"]) - h_qty
                if new_qty < 0:
                    return False, f"취소 시 제품 재고가 음수({new_qty})가 됩니다."
                self.update_product_qty(item_id, new_qty)
                self._add_history("취소", "생산입고취소", item_id, item_name, h_qty, new_qty, "", "이력 취소")
                return True, f"생산 취소: {item_name} -{h_qty}개 (제품재고: {new_qty}개)"

        return False, "취소할 수 없는 이력 유형입니다."

    def _add_history(self, direction, h_type, item_id, item_name, qty, remaining, related="", note=""):
        """이력 추가"""
        ws = self.spreadsheet.worksheet(SHEET_HISTORY)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ws.append_row([now, direction, h_type, str(item_id), item_name, int(qty), int(remaining), related, note])

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
        """안전재고 이하 부품 목록 조회"""
        parts = self.get_all_parts()
        alerts = []
        for p in parts:
            safety = int(p.get("안전재고", 0))
            current = int(p.get("현재재고", 0))
            if safety > 0 and current <= safety:
                alerts.append(p)
        return alerts
