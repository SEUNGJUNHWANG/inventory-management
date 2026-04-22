"""
재고관리 시스템 - BOM 관리 페이지
- 제품코드별 중복 없는 목록 표시
- 더블클릭 시 해당 제품의 부품 명세 상세 창
- BOM 대량 등록 (엑셀 양식)
- BOM 개별 추가/수정/삭제
- 제품별 원가 계산
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import os
from core.constants import COLORS, FONT_FAMILY, FONT_SIZES, BOM_COLUMNS


class BomPage:
    def __init__(self, app):
        self.app = app
        self.product_tree = None
        self.products_map = {}
        self.parts_map = {}
        self.parts_data = {}  # 품번 → 전체 부품 정보
        self.bom_data = []    # 전체 BOM 데이터 캐시

    def render(self):
        scroll_frame = self.app._create_scrollable_frame()

        # ── 헤더 ──
        header = tk.Frame(scroll_frame, bg=COLORS["bg"])
        header.pack(fill=tk.X, padx=5, pady=(0, 10))
        tk.Label(header, text="📋 BOM (소요량) 관리", bg=COLORS["bg"],
                 fg=COLORS["text"], font=(FONT_FAMILY, FONT_SIZES["title"], "bold")).pack(side=tk.LEFT)

        # 버튼 영역
        btn_frame = tk.Frame(header, bg=COLORS["bg"])
        btn_frame.pack(side=tk.RIGHT)

        tk.Button(btn_frame, text="📝 BOM 양식 다운로드",
                  font=(FONT_FAMILY, FONT_SIZES["small"]),
                  bg="#059669", fg="white", padx=10, pady=4,
                  cursor="hand2", command=self._download_bom_template).pack(side=tk.LEFT, padx=3)

        tk.Button(btn_frame, text="📄 엑셀 대량 등록",
                  font=(FONT_FAMILY, FONT_SIZES["small"]),
                  bg="#7c3aed", fg="white", padx=10, pady=4,
                  cursor="hand2", command=self._bulk_upload_bom).pack(side=tk.LEFT, padx=3)

        tk.Button(btn_frame, text="+ BOM 추가",
                  font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                  bg=COLORS["primary"], fg="white", padx=10, pady=4,
                  cursor="hand2", command=self._add_bom_dialog).pack(side=tk.LEFT, padx=3)

        tk.Button(btn_frame, text="💰 전체 원가 요약",
                  font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                  bg="#f59e0b", fg="white", padx=10, pady=4,
                  cursor="hand2", command=self._show_all_products_cost).pack(side=tk.LEFT, padx=3)

        # 설명
        tk.Label(scroll_frame,
                 text="제품을 더블클릭하면 해당 제품의 부품 명세를 확인할 수 있습니다. 우클릭으로 제품 BOM 전체 삭제가 가능합니다.",
                 bg=COLORS["bg"], fg=COLORS["text_secondary"],
                 font=(FONT_FAMILY, FONT_SIZES["small"])).pack(anchor="w", padx=5, pady=(0, 10))

        # ── 제품별 BOM 목록 테이블 ──
        card = tk.Frame(scroll_frame, bg=COLORS["card_bg"],
                        highlightbackground=COLORS["border"], highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True, padx=5)

        tree_frame = tk.Frame(card, bg=COLORS["card_bg"])
        tree_frame.pack(fill=tk.BOTH, expand=True)

        y_scroll = ttk.Scrollbar(tree_frame, orient="vertical")
        y_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        product_cols = ("No", "제품코드", "제품명", "규격", "BOM 부품 수", "총 원가", "현재재고")
        self.product_tree = ttk.Treeview(tree_frame, columns=product_cols, show="headings", height=22,
                                          yscrollcommand=y_scroll.set)
        y_scroll.config(command=self.product_tree.yview)

        col_widths = {"No": 50, "제품코드": 130, "제품명": 220, "규격": 150,
                      "BOM 부품 수": 110, "총 원가": 130, "현재재고": 100}
        col_anchors = {"No": "center", "제품코드": "center", "제품명": "w", "규격": "w",
                       "BOM 부품 수": "center", "총 원가": "e", "현재재고": "center"}

        for col in product_cols:
            self.product_tree.heading(col, text=col)
            self.product_tree.column(col, width=col_widths.get(col, 120),
                                      anchor=col_anchors.get(col, "center"))

        self.product_tree.pack(fill=tk.BOTH, expand=True)

        # 반응형: 제품명 컬럼 너비 자동 조정
        def _on_bom_tree_resize(event):
            total = self.product_tree.winfo_width()
            fixed = 50 + 130 + 150 + 110 + 130 + 100 + 20  # No+제품코드+규격+BOM부품수+원가+재고+스크롤바
            remaining = max(150, total - fixed)
            self.product_tree.column("제품명", width=remaining)
        self.product_tree.bind("<Configure>", _on_bom_tree_resize)

        # 더블클릭 → 상세 명세 창
        self.product_tree.bind("<Double-1>", self._on_product_double_click)

        # 우클릭 메뉴
        self.product_menu = tk.Menu(self.app.root, tearoff=0)
        self.product_menu.add_command(label="📋 부품 명세 보기", command=self._open_detail_from_menu)
        self.product_menu.add_command(label="💰 원가 분석", command=self._show_cost_from_menu)
        self.product_menu.add_separator()
        self.product_menu.add_command(label="🗑️ 이 제품 BOM 전체 삭제", command=self._delete_product_bom)
        self.product_tree.bind("<Button-3>", self._right_click_product)

        self._load_data()

    def _load_data(self):
        """BOM 데이터 로드 → 제품별 그룹핑"""
        def load():
            try:
                bom = self.app.db.get_all_bom()
                products = self.app.db.get_all_products()
                parts = self.app.db.get_all_parts()
                self.app.root.after(0, lambda: render(bom, products, parts))
            except Exception as e:
                err_msg = str(e)
                self.app.root.after(0, lambda: messagebox.showerror("오류", err_msg))

        def render(bom, products, parts):
            self.products_map = {str(p["제품코드"]): p for p in products}
            self.parts_map = {str(p["품번"]): p["부품명"] for p in parts}
            self.parts_data = {str(p["품번"]): p for p in parts}
            self.bom_data = bom
            self._render_product_table()

        threading.Thread(target=load, daemon=True).start()

    def _render_product_table(self):
        """제품코드별 중복 없는 목록 렌더링"""
        self.product_tree.delete(*self.product_tree.get_children())

        # BOM을 제품코드별로 그룹핑
        product_bom = {}  # 제품코드 → [bom_items]
        for b in self.bom_data:
            prod_id = str(b.get("제품코드", ""))
            if prod_id not in product_bom:
                product_bom[prod_id] = []
            product_bom[prod_id].append(b)

        # BOM이 등록된 제품만 표시 (제품코드 순 정렬)
        for idx, (prod_id, bom_items) in enumerate(sorted(product_bom.items()), 1):
            prod_info = self.products_map.get(prod_id, {})
            prod_name = prod_info.get("제품명", "?")
            prod_spec = prod_info.get("규격", "")
            prod_stock = prod_info.get("현재재고", 0)
            part_count = len(bom_items)

            # 총 원가 계산
            total_cost = 0
            for b in bom_items:
                qty = float(b.get("소요량", 0))
                price = float(b.get("단가", 0))
                total_cost += qty * price

            self.product_tree.insert("", "end", values=(
                idx,
                prod_id,
                prod_name,
                prod_spec,
                f"{part_count}종",
                f"{total_cost:,.0f}원" if total_cost > 0 else "미설정",
                prod_stock,
            ))

    # ─────────────────────────────────────────
    # 더블클릭 → 제품 부품 명세 상세 창
    # ─────────────────────────────────────────
    def _on_product_double_click(self, event):
        """제품 더블클릭 시 부품 명세 상세 창 열기"""
        selected = self.product_tree.selection()
        if not selected:
            return
        values = self.product_tree.item(selected[0])["values"]
        prod_code = str(values[1])
        self._open_bom_detail(prod_code)

    def _open_detail_from_menu(self):
        """우클릭 메뉴에서 부품 명세 보기"""
        selected = self.product_tree.selection()
        if not selected:
            return
        values = self.product_tree.item(selected[0])["values"]
        prod_code = str(values[1])
        self._open_bom_detail(prod_code)

    def _open_bom_detail(self, prod_code):
        """특정 제품의 부품 명세 상세 창"""
        prod_info = self.products_map.get(prod_code, {})
        prod_name = prod_info.get("제품명", "?")
        prod_spec = prod_info.get("규격", "")

        # 해당 제품의 BOM 필터링
        bom_items = [b for b in self.bom_data if str(b.get("제품코드", "")) == prod_code]

        dialog = tk.Toplevel(self.app.root)
        dialog.title(f"부품 명세 - {prod_code} / {prod_name}")
        dialog.transient(self.app.root)
        dialog.grab_set()

        # 화면 크기의 70%로 설정
        sw = dialog.winfo_screenwidth()
        sh = dialog.winfo_screenheight()
        dw = max(int(sw * 0.7), 900)
        dh = max(int(sh * 0.65), 550)
        x = (sw - dw) // 2
        y = (sh - dh) // 2
        dialog.geometry(f"{dw}x{dh}+{x}+{y}")
        dialog.minsize(800, 500)

        # ── 헤더 영역 ──
        header = tk.Frame(dialog, bg=COLORS["primary"], padx=15, pady=12)
        header.pack(fill=tk.X)

        tk.Label(header, text=f"📋 {prod_name}", bg=COLORS["primary"], fg="white",
                 font=(FONT_FAMILY, FONT_SIZES["heading"], "bold")).pack(anchor="w")

        info_text = f"제품코드: {prod_code}"
        if prod_spec:
            info_text += f"  |  규격: {prod_spec}"
        info_text += f"  |  BOM 부품: {len(bom_items)}종"

        tk.Label(header, text=info_text, bg=COLORS["primary"], fg="#e0e7ff",
                 font=(FONT_FAMILY, FONT_SIZES["small"])).pack(anchor="w", pady=(3, 0))

        # 총 원가
        total_cost = 0
        for b in bom_items:
            qty = float(b.get("소요량", 0))
            price = float(b.get("단가", 0))
            total_cost += qty * price

        if total_cost > 0:
            tk.Label(header, text=f"💰 총 원가: {total_cost:,.0f}원", bg=COLORS["primary"], fg="#fbbf24",
                     font=(FONT_FAMILY, FONT_SIZES["body"], "bold")).pack(anchor="w", pady=(3, 0))

        # ── 버튼 영역 ──
        action_frame = tk.Frame(dialog, bg=COLORS["bg"], padx=10, pady=8)
        action_frame.pack(fill=tk.X)

        tk.Button(action_frame, text="+ BOM 항목 추가",
                  font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                  bg=COLORS["primary"], fg="white", padx=10, pady=3,
                  cursor="hand2",
                  command=lambda: self._add_bom_to_product(prod_code, dialog)).pack(side=tk.LEFT, padx=3)

        tk.Button(action_frame, text="💰 원가 분석",
                  font=(FONT_FAMILY, FONT_SIZES["small"]),
                  bg="#f59e0b", fg="white", padx=10, pady=3,
                  cursor="hand2",
                  command=lambda: self._show_cost_dialog(prod_code)).pack(side=tk.LEFT, padx=3)

        # ── 부품 명세 테이블 ──
        tree_frame = tk.Frame(dialog, bg=COLORS["card_bg"])
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        detail_cols = ("No", "품번", "부품명", "규격", "단위", "업체명", "소요량", "현재재고", "단가", "소계", "비고")
        detail_tree = ttk.Treeview(tree_frame, columns=detail_cols, show="headings", height=18)

        y_sb = ttk.Scrollbar(tree_frame, orient="vertical", command=detail_tree.yview)
        y_sb.pack(side=tk.RIGHT, fill=tk.Y)
        detail_tree.configure(yscrollcommand=y_sb.set)

        col_cfg = {
            "No": (45, "center"), "품번": (120, "center"), "부품명": (160, "w"),
            "규격": (120, "w"), "단위": (55, "center"), "업체명": (120, "w"),
            "소요량": (70, "center"), "현재재고": (85, "center"),
            "단가": (90, "e"), "소계": (90, "e"), "비고": (120, "w"),
        }

        for col in detail_cols:
            w, anchor = col_cfg.get(col, (100, "center"))
            detail_tree.heading(col, text=col)
            detail_tree.column(col, width=w, anchor=anchor)

        # 재고 부족 행 스타일
        detail_tree.tag_configure("shortage", foreground="#ef4444", background="#fef2f2")
        detail_tree.tag_configure("normal", foreground="#1e293b")

        for i, b in enumerate(bom_items, 1):
            part_id = str(b.get("부품품번", ""))
            part_info = self.parts_data.get(part_id, {})
            part_name = part_info.get("부품명", "?")
            part_spec = part_info.get("규격", "")
            part_unit = part_info.get("단위", "EA")
            part_supplier = part_info.get("업체명", "")
            current_stock = int(part_info.get("현재재고", 0))
            qty = float(b.get("소요량", 0))
            unit_price = float(b.get("단가", 0))
            subtotal = qty * unit_price
            note = b.get("비고", "")

            # 재고 부족 여부 판단
            tag = "shortage" if current_stock < qty else "normal"

            detail_tree.insert("", "end", values=(
                i,
                part_id,
                part_name,
                part_spec,
                part_unit,
                part_supplier,
                qty,
                current_stock,
                f"{unit_price:,.0f}" if unit_price > 0 else "-",
                f"{subtotal:,.0f}" if subtotal > 0 else "-",
                note,
            ), tags=(tag,))

        detail_tree.pack(fill=tk.BOTH, expand=True)

        # 우클릭 메뉴 (상세 테이블용)
        detail_menu = tk.Menu(dialog, tearoff=0)
        detail_menu.add_command(label="✏️ 수정",
                                command=lambda: self._edit_bom_in_detail(detail_tree, prod_code, dialog))
        detail_menu.add_command(label="🗑️ 삭제",
                                command=lambda: self._delete_bom_in_detail(detail_tree, prod_code, dialog))

        def on_detail_right_click(event):
            item = detail_tree.identify_row(event.y)
            if item:
                detail_tree.selection_set(item)
                detail_menu.post(event.x_root, event.y_root)

        detail_tree.bind("<Button-3>", on_detail_right_click)
        detail_tree.bind("<Double-1>",
                         lambda e: self._edit_bom_in_detail(detail_tree, prod_code, dialog))

        # ── 하단 닫기 버튼 ──
        btn_frame = tk.Frame(dialog, bg=COLORS["bg"], padx=10, pady=8)
        btn_frame.pack(fill=tk.X)

        tk.Button(btn_frame, text="닫기", font=(FONT_FAMILY, FONT_SIZES["small"]),
                  padx=15, pady=5, command=dialog.destroy).pack(side=tk.RIGHT, padx=5)

    def _add_bom_to_product(self, prod_code, parent_dialog):
        """상세 창에서 특정 제품에 BOM 항목 추가"""
        dialog = tk.Toplevel(parent_dialog)
        dialog.title(f"BOM 추가 - {prod_code}")
        dialog.geometry("450x250")
        dialog.resizable(False, False)
        dialog.transient(parent_dialog)
        dialog.grab_set()

        fields = {}
        labels = [("부품품번", ""), ("소요량", "1"), ("비고", "")]

        for i, (label, default) in enumerate(labels):
            tk.Label(dialog, text=label + ":", font=(FONT_FAMILY, FONT_SIZES["small"])).grid(
                row=i, column=0, padx=10, pady=5, sticky="e")
            entry = tk.Entry(dialog, font=(FONT_FAMILY, FONT_SIZES["small"]), width=30)
            entry.grid(row=i, column=1, padx=10, pady=5)
            entry.insert(0, default)
            fields[label] = entry

        def save():
            try:
                part_code = fields["부품품번"].get().strip()
                qty = float(fields["소요량"].get())
                note = fields["비고"].get().strip()

                if not part_code:
                    messagebox.showwarning("입력 오류", "부품품번은 필수입니다.")
                    return

                # 부품 등록 여부 확인
                if part_code not in self.parts_data:
                    messagebox.showwarning("미등록 부품",
                                           f"부품품번 '{part_code}'이(가) 부품 관리에 등록되어 있지 않습니다.\n"
                                           "먼저 부품 관리에서 등록해 주세요.")
                    return

                # 단가는 부품마스터에서만 관리 — BOM에 별도 저장하지 않음
                self.app.db.add_bom(prod_code, part_code, qty, note)
                messagebox.showinfo("성공", "BOM 항목이 추가되었습니다.")
                dialog.destroy()
                parent_dialog.destroy()
                self._load_data()
                # 다시 상세 창 열기
                self.app.root.after(500, lambda: self._open_bom_detail(prod_code))
            except Exception as e:
                messagebox.showerror("오류", str(e))

        tk.Button(dialog, text="저장", font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                  bg=COLORS["primary"], fg="white", padx=20, pady=5,
                  command=save).grid(row=len(labels), column=0, columnspan=2, pady=15)

    def _edit_bom_in_detail(self, detail_tree, prod_code, parent_dialog):
        """상세 테이블에서 BOM 항목 수정"""
        selected = detail_tree.selection()
        if not selected:
            messagebox.showinfo("안내", "수정할 항목을 선택해 주세요.")
            return

        values = detail_tree.item(selected[0])["values"]
        part_code = str(values[1])
        part_name = str(values[2])
        current_qty = str(values[6])  # 업체명 추가로 인덱스 +1
        current_price = str(values[8]).replace(",", "").replace("-", "0")
        current_note = str(values[10]) if len(values) > 10 else ""

        dialog = tk.Toplevel(parent_dialog)
        dialog.title(f"BOM 수정 - {part_name}")
        dialog.geometry("500x360")
        dialog.resizable(False, False)
        dialog.transient(parent_dialog)
        dialog.grab_set()

        # 부품 정보 (읽기 전용)
        info_frame = tk.Frame(dialog, bg=COLORS["bg"], padx=15, pady=10)
        info_frame.pack(fill=tk.X)

        prod_name = self.products_map.get(prod_code, {}).get("제품명", "?")
        tk.Label(info_frame, text=f"제품: {prod_code} - {prod_name}",
                 bg=COLORS["bg"], fg=COLORS["text"],
                 font=(FONT_FAMILY, FONT_SIZES["body"], "bold")).pack(anchor="w")
        tk.Label(info_frame, text=f"부품: {part_code} - {part_name}",
                 bg=COLORS["bg"], fg=COLORS["text_secondary"],
                 font=(FONT_FAMILY, FONT_SIZES["small"])).pack(anchor="w", pady=(2, 0))

        # 수정 필드
        edit_frame = tk.Frame(dialog, padx=15, pady=10)
        edit_frame.pack(fill=tk.X)

        tk.Label(edit_frame, text="소요량:", font=(FONT_FAMILY, FONT_SIZES["body"])).grid(
            row=0, column=0, padx=5, pady=8, sticky="e")
        qty_entry = tk.Entry(edit_frame, font=(FONT_FAMILY, FONT_SIZES["body"]), width=20)
        qty_entry.grid(row=0, column=1, padx=5, pady=8)
        qty_entry.insert(0, current_qty)

        # 단가 조회 (참고용 표시만, 수정은 부품관리에서)
        part_info_for_price = self.app.db.get_part_by_id(part_code)
        auto_price_val = float(part_info_for_price.get("단가", 0) or 0) if part_info_for_price else 0

        price_info_frame = tk.Frame(edit_frame)
        price_info_frame.grid(row=1, column=0, columnspan=2, sticky="w", padx=5, pady=4)
        tk.Label(price_info_frame,
                 text=f"단가: {auto_price_val:,.0f}원  (부품마스터 기준 / 변경은 부품관리 메뉴에서)",
                 font=(FONT_FAMILY, FONT_SIZES["small"]), fg="#6b7280").pack(side=tk.LEFT)

        tk.Label(edit_frame, text="비고:", font=(FONT_FAMILY, FONT_SIZES["body"])).grid(
            row=2, column=0, padx=5, pady=8, sticky="e")
        note_entry = tk.Entry(edit_frame, font=(FONT_FAMILY, FONT_SIZES["body"]), width=20)
        note_entry.grid(row=2, column=1, padx=5, pady=8)
        note_entry.insert(0, current_note)

        # 소계 미리보기
        subtotal_label = tk.Label(edit_frame, text="", fg=COLORS["primary"],
                                   font=(FONT_FAMILY, FONT_SIZES["body"], "bold"))
        subtotal_label.grid(row=3, column=0, columnspan=2, pady=5)

        def update_subtotal(*args):
            try:
                q = float(qty_entry.get() or 0)
                subtotal_label.config(text=f"소계(참고): {q * auto_price_val:,.0f}원")
            except:
                subtotal_label.config(text="")

        qty_entry.bind("<KeyRelease>", update_subtotal)
        update_subtotal()

        def save():
            try:
                new_qty = float(qty_entry.get())
                new_note = note_entry.get().strip()

                if new_qty <= 0:
                    messagebox.showwarning("입력 오류", "소요량은 0보다 커야 합니다.")
                    return

                # 단가는 부품마스터에서만 관리 — BOM에 별도 저장하지 않음
                self.app.db.update_bom(prod_code, part_code, new_qty, new_note)
                messagebox.showinfo("성공", "BOM 항목이 수정되었습니다.")
                dialog.destroy()
                parent_dialog.destroy()
                self._load_data()
                self.app.root.after(500, lambda: self._open_bom_detail(prod_code))
            except Exception as e:
                messagebox.showerror("오류", str(e))

        btn_frame = tk.Frame(dialog, padx=15, pady=10)
        btn_frame.pack(fill=tk.X)

        tk.Button(btn_frame, text="취소", font=(FONT_FAMILY, FONT_SIZES["small"]),
                  padx=15, pady=5, command=dialog.destroy).pack(side=tk.RIGHT, padx=5)
        tk.Button(btn_frame, text="💾 저장", font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                  bg=COLORS["primary"], fg="white", padx=20, pady=5,
                  command=save).pack(side=tk.RIGHT, padx=5)

    def _delete_bom_in_detail(self, detail_tree, prod_code, parent_dialog):
        """상세 테이블에서 BOM 항목 삭제"""
        selected = detail_tree.selection()
        if not selected:
            return
        values = detail_tree.item(selected[0])["values"]
        part_code = str(values[1])
        part_name = str(values[2])

        if messagebox.askyesno("삭제 확인", f"BOM 항목을 삭제하시겠습니까?\n부품: {part_code} - {part_name}"):
            try:
                self.app.db.delete_bom(prod_code, part_code)
                messagebox.showinfo("성공", "BOM 항목이 삭제되었습니다.")
                parent_dialog.destroy()
                self._load_data()
                self.app.root.after(500, lambda: self._open_bom_detail(prod_code))
            except Exception as e:
                messagebox.showerror("오류", str(e))

    # ─────────────────────────────────────────
    # 우클릭 메뉴 (제품 목록)
    # ─────────────────────────────────────────
    def _right_click_product(self, event):
        item = self.product_tree.identify_row(event.y)
        if item:
            self.product_tree.selection_set(item)
            self.product_menu.post(event.x_root, event.y_root)

    def _delete_product_bom(self):
        """선택된 제품의 BOM 전체 삭제"""
        selected = self.product_tree.selection()
        if not selected:
            return
        values = self.product_tree.item(selected[0])["values"]
        prod_code = str(values[1])
        prod_name = str(values[2])

        if messagebox.askyesno("삭제 확인",
                               f"'{prod_name}' ({prod_code}) 제품의 BOM을 전체 삭제하시겠습니까?\n"
                               "이 작업은 되돌릴 수 없습니다."):
            try:
                self.app.db.delete_all_bom_for_product(prod_code)
                messagebox.showinfo("성공", f"'{prod_name}' 제품의 BOM이 전체 삭제되었습니다.")
                self._load_data()
            except Exception as e:
                messagebox.showerror("오류", str(e))

    def _show_cost_from_menu(self):
        """우클릭 메뉴에서 원가 분석"""
        selected = self.product_tree.selection()
        if not selected:
            return
        values = self.product_tree.item(selected[0])["values"]
        prod_code = str(values[1])
        self._show_cost_dialog(prod_code)

    # ─────────────────────────────────────────
    # BOM 양식 다운로드
    # ─────────────────────────────────────────
    def _download_bom_template(self):
        """BOM 대량 등록 양식 다운로드"""
        save_path = filedialog.asksaveasfilename(
            title="BOM 양식 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
            initialfile="BOM_대량등록_양식.xlsx"
        )
        if not save_path:
            return

        def generate():
            try:
                from utils.excel_utils import create_bom_template
                products = self.app.db.get_all_products()
                parts = self.app.db.get_all_parts()
                create_bom_template(save_path, products=products, parts=parts)
                self.app.root.after(0, lambda: messagebox.showinfo(
                    "완료",
                    f"BOM 양식이 저장되었습니다.\n\n{save_path}\n\n"
                    "양식에는 '제품목록(참고)'과 '부품목록(참고)' 시트가 포함되어 있어\n"
                    "제품코드와 부품품번을 쉽게 확인할 수 있습니다."
                ))
            except Exception as e:
                err_msg = str(e)
                self.app.root.after(0, lambda: messagebox.showerror("오류", err_msg))

        threading.Thread(target=generate, daemon=True).start()

    # ─────────────────────────────────────────
    # BOM 엑셀 대량 등록
    # ─────────────────────────────────────────
    def _bulk_upload_bom(self):
        """엑셀 파일로 BOM 대량 등록"""
        file_path = filedialog.askopenfilename(
            title="BOM 엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx *.xls")]
        )
        if not file_path:
            return

        def parse():
            try:
                from utils.excel_utils import parse_bom_excel
                items = parse_bom_excel(file_path)
                if not items:
                    self.app.root.after(0, lambda: messagebox.showwarning("경고", "파일에서 BOM 데이터를 찾을 수 없습니다."))
                    return
                # 부품/제품 등록 여부 검증을 위해 마스터 데이터 조회
                products = self.app.db.get_all_products()
                parts = self.app.db.get_all_parts()
                registered_products = {str(p["제품코드"]) for p in products}
                registered_parts = {str(p["품번"]) for p in parts}
                self.app.root.after(0, lambda: self._show_bom_preview(
                    items, file_path, registered_products, registered_parts))
            except Exception as e:
                err_msg = str(e)
                self.app.root.after(0, lambda: messagebox.showerror("파싱 오류", err_msg))

        threading.Thread(target=parse, daemon=True).start()

    def _show_bom_preview(self, items, file_path, registered_products=None, registered_parts=None):
        """BOM 대량 등록 미리보기 대화상자 (미등록 부품/제품 검증 포함)"""
        if registered_products is None:
            registered_products = set()
        if registered_parts is None:
            registered_parts = set()

        dialog = tk.Toplevel(self.app.root)
        dialog.title(f"BOM 대량 등록 미리보기 - {len(items)}건")
        dialog.transient(self.app.root)
        dialog.grab_set()

        # 화면 크기의 75%로 설정
        sw = dialog.winfo_screenwidth()
        sh = dialog.winfo_screenheight()
        dw = max(int(sw * 0.75), 900)
        dh = max(int(sh * 0.75), 600)
        x = (sw - dw) // 2
        y = (sh - dh) // 2
        dialog.geometry(f"{dw}x{dh}+{x}+{y}")
        dialog.minsize(900, 600)

        # ── 미등록 항목 검증 ──
        invalid_items = []
        valid_items = []
        unregistered_parts = set()
        unregistered_products = set()

        for item in items:
            prod_code = str(item["제품코드"])
            part_code = str(item["부품품번"])
            is_valid = True
            reasons = []

            if prod_code not in registered_products:
                is_valid = False
                reasons.append(f"미등록 제품코드: {prod_code}")
                unregistered_products.add(prod_code)

            if part_code not in registered_parts:
                is_valid = False
                reasons.append(f"미등록 부품품번: {part_code}")
                unregistered_parts.add(part_code)

            if is_valid:
                valid_items.append(item)
            else:
                invalid_items.append((item, reasons))

        has_errors = len(invalid_items) > 0

        # 상단 정보
        info_frame = tk.Frame(dialog, bg=COLORS["bg"], padx=10, pady=10)
        info_frame.pack(fill=tk.X)

        tk.Label(info_frame, text=f"📄 파일: {os.path.basename(file_path)}",
                 bg=COLORS["bg"], font=(FONT_FAMILY, FONT_SIZES["small"])).pack(anchor="w")

        # 제품별 집계
        product_counts = {}
        for item in items:
            pc = item["제품코드"]
            product_counts[pc] = product_counts.get(pc, 0) + 1

        tk.Label(info_frame,
                 text=f"총 {len(items)}건 (제품 {len(product_counts)}종, 부품 {len(set(i['부품품번'] for i in items))}종)",
                 bg=COLORS["bg"], fg=COLORS["primary"],
                 font=(FONT_FAMILY, FONT_SIZES["body"], "bold")).pack(anchor="w", pady=(5, 0))

        # 미등록 경고 메시지
        if has_errors:
            error_frame = tk.Frame(dialog, bg="#fef2f2", padx=10, pady=8)
            error_frame.pack(fill=tk.X, padx=10, pady=(5, 0))

            error_msg = f"⛔ 미등록 항목 {len(invalid_items)}건이 발견되었습니다. 먼저 등록 후 다시 시도해 주세요."
            tk.Label(error_frame, text=error_msg,
                     bg="#fef2f2", fg=COLORS["danger"],
                     font=(FONT_FAMILY, FONT_SIZES["body"], "bold")).pack(anchor="w")

            detail_parts = []
            if unregistered_products:
                detail_parts.append(f"미등록 제품코드 {len(unregistered_products)}종: {', '.join(sorted(unregistered_products))}")
            if unregistered_parts:
                detail_parts.append(f"미등록 부품품번 {len(unregistered_parts)}종: {', '.join(sorted(unregistered_parts))}")

            for dp in detail_parts:
                tk.Label(error_frame, text=f"  → {dp}",
                         bg="#fef2f2", fg=COLORS["danger"],
                         font=(FONT_FAMILY, FONT_SIZES["small"])).pack(anchor="w")

            tk.Label(error_frame,
                     text="💡 부품 관리 또는 제품 관리에서 먼저 등록한 후 다시 업로드하세요.",
                     bg="#fef2f2", fg=COLORS["text_secondary"],
                     font=(FONT_FAMILY, FONT_SIZES["small"])).pack(anchor="w", pady=(3, 0))

        # 미리보기 테이블
        tree_frame = tk.Frame(dialog)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 단가는 부품마스터에서만 관리하므로 미리보기 컬럼에서 제외
        cols = ("No", "상태", "제품코드", "부품품번", "소요량", "비고")
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=20)

        y_sb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        y_sb.pack(side=tk.RIGHT, fill=tk.Y)
        tree.configure(yscrollcommand=y_sb.set)

        for col in cols:
            tree.heading(col, text=col)
            if col == "No":
                tree.column(col, width=50, anchor="center")
            elif col == "상태":
                tree.column(col, width=120, anchor="center")
            elif col == "비고":
                tree.column(col, width=250, anchor="w")
            else:
                tree.column(col, width=140, anchor="center")

        # 에러 행 스타일
        tree.tag_configure("error", foreground="#ef4444", background="#fef2f2")
        tree.tag_configure("ok", foreground="#1e293b")

        for i, item in enumerate(items, 1):
            prod_code = str(item["제품코드"])
            part_code = str(item["부품품번"])

            # 상태 판별
            errors = []
            if prod_code not in registered_products:
                errors.append("제품 미등록")
            if part_code not in registered_parts:
                errors.append("부품 미등록")

            if errors:
                status = "❌ " + ", ".join(errors)
                tag = "error"
            else:
                status = "✅ 정상"
                tag = "ok"

            tree.insert("", "end", values=(
                i, status, item["제품코드"], item["부품품번"],
                item["소요량"], item.get("비고", "")
            ), tags=(tag,))

        tree.pack(fill=tk.BOTH, expand=True)

        # 하단 버튼
        btn_frame = tk.Frame(dialog, bg=COLORS["bg"], padx=10, pady=10)
        btn_frame.pack(fill=tk.X)

        self._upload_progress_label = tk.Label(btn_frame, text="", bg=COLORS["bg"],
                                                fg=COLORS["text_secondary"],
                                                font=(FONT_FAMILY, FONT_SIZES["small"]))
        self._upload_progress_label.pack(side=tk.LEFT)

        tk.Button(btn_frame, text="취소", font=(FONT_FAMILY, FONT_SIZES["small"]),
                  padx=15, pady=5, command=dialog.destroy).pack(side=tk.RIGHT, padx=5)

        if has_errors:
            # 미등록 항목이 있으면 업로드 버튼 비활성화
            upload_btn = tk.Button(btn_frame, text="⛔ 미등록 항목 해결 필요",
                                   font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                                   bg="#9ca3af", fg="white", padx=20, pady=5,
                                   state="disabled")
            upload_btn.pack(side=tk.RIGHT, padx=5)
        else:
            upload_btn = tk.Button(btn_frame, text=f"✅ {len(valid_items)}건 업데이트 실행",
                                   font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                                   bg=COLORS["primary"], fg="white", padx=20, pady=5,
                                   cursor="hand2",
                                   command=lambda: self._execute_bom_upload(valid_items, dialog, upload_btn))
            upload_btn.pack(side=tk.RIGHT, padx=5)

    def _execute_bom_upload(self, items, dialog, btn):
        """BOM 대량 업로드 실행"""
        btn.config(state="disabled", text="업로드 중...")

        def upload():
            try:
                def progress_cb(msg):
                    self.app.root.after(0, lambda: self._upload_progress_label.config(text=msg))

                new_count, update_count = self.app.db.bulk_add_or_update_bom(items, progress_callback=progress_cb)

                def on_done():
                    messagebox.showinfo("완료",
                                        f"BOM 대량 등록이 완료되었습니다.\n\n"
                                        f"신규 등록: {new_count}건\n"
                                        f"업데이트: {update_count}건")
                    dialog.destroy()
                    self._load_data()

                self.app.root.after(0, on_done)
            except Exception as e:
                err_msg = str(e)
                def on_error():
                    messagebox.showerror("업로드 실패", f"오류가 발생했습니다.\n\n{err_msg}")
                    btn.config(state="normal", text="✅ 업데이트 실행")
                self.app.root.after(0, on_error)

        threading.Thread(target=upload, daemon=True).start()

    # ─────────────────────────────────────────
    # BOM 개별 추가
    # ─────────────────────────────────────────
    def _add_bom_dialog(self):
        dialog = tk.Toplevel(self.app.root)
        dialog.title("BOM 추가")
        dialog.geometry("450x320")
        dialog.resizable(False, False)
        dialog.transient(self.app.root)
        dialog.grab_set()

        fields = {}
        labels = [("제품코드", ""), ("부품품번", ""), ("소요량", "1"), ("비고", "")]

        for i, (label, default) in enumerate(labels):
            tk.Label(dialog, text=label + ":", font=(FONT_FAMILY, FONT_SIZES["small"])).grid(
                row=i, column=0, padx=10, pady=5, sticky="e")
            entry = tk.Entry(dialog, font=(FONT_FAMILY, FONT_SIZES["small"]), width=30)
            entry.grid(row=i, column=1, padx=10, pady=5)
            entry.insert(0, default)
            fields[label] = entry

        def save():
            try:
                prod_code = fields["제품코드"].get().strip()
                part_code = fields["부품품번"].get().strip()
                qty = float(fields["소요량"].get())
                note = fields["비고"].get().strip()

                if not prod_code or not part_code:
                    messagebox.showwarning("입력 오류", "제품코드와 부품품번은 필수입니다.")
                    return

                # 제품/부품 등록 여부 확인
                if prod_code not in self.products_map:
                    messagebox.showwarning("미등록 제품",
                                           f"제품코드 '{prod_code}'이(가) 제품 관리에 등록되어 있지 않습니다.\n"
                                           "먼저 제품 관리에서 등록해 주세요.")
                    return

                if part_code not in self.parts_data:
                    messagebox.showwarning("미등록 부품",
                                           f"부품품번 '{part_code}'이(가) 부품 관리에 등록되어 있지 않습니다.\n"
                                           "먼저 부품 관리에서 등록해 주세요.")
                    return

                # 단가는 부품마스터에서만 관리 — BOM에 별도 저장하지 않음
                self.app.db.add_bom(prod_code, part_code, qty, note)
                messagebox.showinfo("성공", "BOM 항목이 추가되었습니다.")
                dialog.destroy()
                self._load_data()
            except Exception as e:
                messagebox.showerror("오류", str(e))

        tk.Button(dialog, text="저장", font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                  bg=COLORS["primary"], fg="white", padx=20, pady=5,
                  command=save).grid(row=len(labels), column=0, columnspan=2, pady=15)

    # ─────────────────────────────────────────
    # 제품별 원가 분석
    # ─────────────────────────────────────────
    def _show_cost_dialog(self, product_code):
        """특정 제품 원가 분석 대화상자"""
        prod_info = self.products_map.get(product_code, {})
        product_name = prod_info.get("제품명", "?")

        def load():
            try:
                total_cost, details = self.app.db.get_product_cost(product_code)
                self.app.root.after(0, lambda: show(total_cost, details))
            except Exception as e:
                err_msg = str(e)
                self.app.root.after(0, lambda: messagebox.showerror("오류", err_msg))

        def show(total_cost, details):
            dialog = tk.Toplevel(self.app.root)
            dialog.title(f"💰 원가 분석 - {product_name}")
            dialog.geometry("700x500")
            dialog.transient(self.app.root)
            dialog.grab_set()

            # 헤더
            header = tk.Frame(dialog, bg=COLORS["primary"], padx=15, pady=15)
            header.pack(fill=tk.X)

            tk.Label(header, text=f"제품: {product_code} - {product_name}",
                     bg=COLORS["primary"], fg="white",
                     font=(FONT_FAMILY, FONT_SIZES["heading"], "bold")).pack(anchor="w")
            tk.Label(header, text=f"총 원가: {total_cost:,.0f}원",
                     bg=COLORS["primary"], fg="white",
                     font=(FONT_FAMILY, FONT_SIZES["stat"], "bold")).pack(anchor="w", pady=(5, 0))

            # 상세 테이블
            tree_frame = tk.Frame(dialog)
            tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

            cols = ("부품품번", "부품명", "소요량", "단가", "소계", "비율")
            tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=15)

            for col in cols:
                tree.heading(col, text=col)
                if col in ("단가", "소계"):
                    tree.column(col, width=100, anchor="e")
                elif col == "비율":
                    tree.column(col, width=70, anchor="center")
                elif col == "부품명":
                    tree.column(col, width=180, anchor="w")
                else:
                    tree.column(col, width=100, anchor="center")

            for d in details:
                ratio = (d["소계"] / total_cost * 100) if total_cost > 0 else 0
                tree.insert("", "end", values=(
                    d["부품품번"], d["부품명"], d["소요량"],
                    f"{d['단가']:,.0f}", f"{d['소계']:,.0f}",
                    f"{ratio:.1f}%"
                ))

            tree.pack(fill=tk.BOTH, expand=True)

            tk.Button(dialog, text="닫기", font=(FONT_FAMILY, FONT_SIZES["small"]),
                      padx=15, pady=5, command=dialog.destroy).pack(pady=10)

        threading.Thread(target=load, daemon=True).start()

    def _show_all_products_cost(self):
        """전체 제품 원가 요약"""
        def load():
            try:
                products = self.app.db.get_all_products()
                cost_summary = []
                for p in products:
                    pid = str(p["제품코드"])
                    total_cost, _ = self.app.db.get_product_cost(pid)
                    cost_summary.append({
                        "제품코드": pid,
                        "제품명": p["제품명"],
                        "원가": total_cost,
                    })
                self.app.root.after(0, lambda: show(cost_summary))
            except Exception as e:
                err_msg = str(e)
                self.app.root.after(0, lambda: messagebox.showerror("오류", err_msg))

        def show(summary):
            dialog = tk.Toplevel(self.app.root)
            dialog.title("💰 전체 제품 원가 요약")
            dialog.geometry("600x450")
            dialog.transient(self.app.root)
            dialog.grab_set()

            tk.Label(dialog, text="💰 전체 제품 원가 요약",
                     font=(FONT_FAMILY, FONT_SIZES["heading"], "bold"),
                     padx=15, pady=10).pack(anchor="w")

            tk.Label(dialog, text="※ 특정 제품의 상세 원가를 보려면 제품을 더블클릭한 후 '원가 분석' 버튼을 누르세요.",
                     fg=COLORS["text_secondary"],
                     font=(FONT_FAMILY, FONT_SIZES["small"]),
                     padx=15).pack(anchor="w")

            tree_frame = tk.Frame(dialog)
            tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

            cols = ("제품코드", "제품명", "원가")
            tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=15)
            for col in cols:
                tree.heading(col, text=col)
                if col == "원가":
                    tree.column(col, width=120, anchor="e")
                elif col == "제품명":
                    tree.column(col, width=250, anchor="w")
                else:
                    tree.column(col, width=120, anchor="center")

            for s in summary:
                tree.insert("", "end", values=(
                    s["제품코드"], s["제품명"],
                    f"{s['원가']:,.0f}원" if s["원가"] > 0 else "미설정"
                ))

            tree.pack(fill=tk.BOTH, expand=True)

            tk.Button(dialog, text="닫기", font=(FONT_FAMILY, FONT_SIZES["small"]),
                      padx=15, pady=5, command=dialog.destroy).pack(pady=10)

        threading.Thread(target=load, daemon=True).start()
