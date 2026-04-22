"""
재고관리 시스템 - MRP(자재소요계획) 페이지
- 생산 계획 입력 (자동완성 검색)
- 소요 부품 계산 및 발주 리스트 출력
- 안전재고 반영 옵션
- 최대 생산 가능 수량 표시
- 엑셀 내보내기
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import os
from datetime import datetime
from core.constants import COLORS, FONT_FAMILY, FONT_SIZES


class MrpPage:
    """MRP(자재소요계획) 페이지"""

    def __init__(self, app):
        self.app = app
        self.production_plan = []  # [{"product_id", "product_name", "target_qty"}, ...]
        self.mrp_result = None
        self.products_cache = []

    def render(self):
        """페이지 렌더링"""
        # 스크롤 가능한 프레임
        self.scroll_frame = self.app._create_scrollable_frame()

        # ── 타이틀 ──
        title_frame = tk.Frame(self.scroll_frame, bg=COLORS["bg"])
        title_frame.pack(fill=tk.X, padx=5, pady=(5, 10))
        tk.Label(title_frame, text="📋 자재소요계획 (MRP)",
                 bg=COLORS["bg"], fg=COLORS["text"],
                 font=(FONT_FAMILY, FONT_SIZES["title"], "bold")).pack(side=tk.LEFT)

        # ── 1. 생산 계획 입력 카드 ──
        self._build_plan_input_card()

        # ── 2. 생산 계획 목록 카드 ──
        self._build_plan_list_card()

        # ── 3. 소요량 계산 버튼 영역 ──
        self._build_action_bar()

        # ── 4. 결과 영역 (계산 후 표시) ──
        self.result_frame = tk.Frame(self.scroll_frame, bg=COLORS["bg"])
        self.result_frame.pack(fill=tk.BOTH, expand=True, padx=5)

        # 제품 목록 캐시 로드
        self._load_products_cache()

    # ═══════════════════════════════════════════
    # 1. 생산 계획 입력 카드
    # ═══════════════════════════════════════════
    def _build_plan_input_card(self):
        card = tk.Frame(self.scroll_frame, bg=COLORS["card_bg"],
                        highlightbackground=COLORS["border"], highlightthickness=1)
        card.pack(fill=tk.X, padx=5, pady=(0, 5))

        tk.Label(card, text="생산 계획 입력",
                 bg=COLORS["card_bg"], fg=COLORS["text"],
                 font=(FONT_FAMILY, FONT_SIZES["heading"], "bold")).pack(
            anchor="w", padx=15, pady=(12, 5))

        tk.Label(card, text="제품을 검색하여 선택하고, 목표 수량을 입력한 뒤 추가하세요.",
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"],
                 font=(FONT_FAMILY, FONT_SIZES["small"])).pack(anchor="w", padx=15, pady=(0, 8))

        form = tk.Frame(card, bg=COLORS["card_bg"])
        form.pack(fill=tk.X, padx=15, pady=(0, 12))

        # 제품 검색
        tk.Label(form, text="제품 검색:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).grid(
            row=0, column=0, sticky="e", padx=(0, 8), pady=5)

        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(form, textvariable=self.search_var,
                                     font=(FONT_FAMILY, FONT_SIZES["body"]), width=35)
        self.search_entry.grid(row=0, column=1, padx=(0, 5), pady=5, sticky="w")
        self.search_entry.bind("<KeyRelease>", self._on_search_change)
        self.search_entry.bind("<FocusIn>", self._on_search_focus)

        # 제품 정보 라벨
        self.product_info_label = tk.Label(form, text="", bg=COLORS["card_bg"],
                                           fg=COLORS["info"],
                                           font=(FONT_FAMILY, FONT_SIZES["small"]))
        self.product_info_label.grid(row=0, column=2, padx=10, sticky="w")

        # 자동완성 리스트박스 (오버레이)
        self.autocomplete_frame = tk.Frame(form, bg="white",
                                           highlightbackground=COLORS["border"],
                                           highlightthickness=1)
        self.autocomplete_listbox = tk.Listbox(self.autocomplete_frame,
                                                font=(FONT_FAMILY, FONT_SIZES["small"]),
                                                height=6, width=50,
                                                selectbackground=COLORS["primary"],
                                                selectforeground="white",
                                                activestyle="none",
                                                cursor="hand2")
        self.autocomplete_listbox.pack(fill=tk.BOTH, expand=True)
        self.autocomplete_listbox.bind("<<ListboxSelect>>", self._on_autocomplete_select)
        self.autocomplete_listbox.bind("<Double-Button-1>", self._on_autocomplete_select)

        # 목표 수량
        tk.Label(form, text="목표 수량:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).grid(
            row=1, column=0, sticky="e", padx=(0, 8), pady=5)

        qty_frame = tk.Frame(form, bg=COLORS["card_bg"])
        qty_frame.grid(row=1, column=1, sticky="w", pady=5)

        self.qty_entry = tk.Entry(qty_frame, font=(FONT_FAMILY, FONT_SIZES["body"]), width=12)
        self.qty_entry.pack(side=tk.LEFT, padx=(0, 10))
        self.qty_entry.bind("<Return>", lambda e: self._add_to_plan())

        # 최대 생산 가능 수량 라벨
        self.max_prod_label = tk.Label(qty_frame, text="",
                                        bg=COLORS["card_bg"], fg=COLORS["text_secondary"],
                                        font=(FONT_FAMILY, FONT_SIZES["small"]))
        self.max_prod_label.pack(side=tk.LEFT)

        # 추가 버튼
        tk.Button(form, text="➕ 계획에 추가",
                  font=(FONT_FAMILY, FONT_SIZES["body"], "bold"),
                  bg=COLORS["primary"], fg="white",
                  padx=15, pady=4, cursor="hand2",
                  command=self._add_to_plan).grid(
            row=1, column=2, padx=10, pady=5, sticky="w")

        # 선택된 제품 ID 저장용
        self.selected_product_id = None
        self.selected_product_name = None

    # ═══════════════════════════════════════════
    # 2. 생산 계획 목록 카드
    # ═══════════════════════════════════════════
    def _build_plan_list_card(self):
        card = tk.Frame(self.scroll_frame, bg=COLORS["card_bg"],
                        highlightbackground=COLORS["border"], highlightthickness=1)
        card.pack(fill=tk.X, padx=5, pady=(0, 5))

        header = tk.Frame(card, bg=COLORS["card_bg"])
        header.pack(fill=tk.X, padx=15, pady=(12, 5))

        tk.Label(header, text="생산 계획 목록",
                 bg=COLORS["card_bg"], fg=COLORS["text"],
                 font=(FONT_FAMILY, FONT_SIZES["heading"], "bold")).pack(side=tk.LEFT)

        self.plan_count_label = tk.Label(header, text="(0개 제품)",
                                          bg=COLORS["card_bg"], fg=COLORS["text_secondary"],
                                          font=(FONT_FAMILY, FONT_SIZES["small"]))
        self.plan_count_label.pack(side=tk.LEFT, padx=10)

        tk.Button(header, text="🗑 전체 삭제",
                  font=(FONT_FAMILY, FONT_SIZES["small"]),
                  bg=COLORS["danger"], fg="white",
                  padx=8, pady=2, cursor="hand2",
                  command=self._clear_plan).pack(side=tk.RIGHT)

        # 트리뷰
        tree_frame = tk.Frame(card, bg=COLORS["card_bg"])
        tree_frame.pack(fill=tk.X, padx=15, pady=(0, 12))

        columns = ("no", "product_id", "product_name", "current_stock",
                    "target_qty", "need_produce", "max_producible", "bottleneck")
        self.plan_tree = ttk.Treeview(tree_frame, columns=columns,
                                       show="headings", height=5)

        col_config = [
            ("no", "No", 40, "center"),
            ("product_id", "제품코드", 100, "center"),
            ("product_name", "제품명", 150, "w"),
            ("current_stock", "현재재고", 80, "center"),
            ("target_qty", "목표수량", 80, "center"),
            ("need_produce", "추가생산필요", 100, "center"),
            ("max_producible", "최대생산가능", 100, "center"),
            ("bottleneck", "병목부품", 150, "w"),
        ]

        for col_id, heading, width, anchor in col_config:
            self.plan_tree.heading(col_id, text=heading)
            self.plan_tree.column(col_id, width=width, anchor=anchor)

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.plan_tree.yview)
        self.plan_tree.configure(yscrollcommand=scrollbar.set)
        self.plan_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 반응형: 제품명/병목부품 컬럼 자동 조정
        def _on_plan_resize(event):
            total = self.plan_tree.winfo_width()
            fixed = 40 + 100 + 80 + 80 + 100 + 100 + 20  # No+코드+현재재고+목표+추가생산+최대+스크롤바
            remaining = max(200, total - fixed)
            self.plan_tree.column("product_name", width=int(remaining * 0.55))
            self.plan_tree.column("bottleneck", width=int(remaining * 0.45))
        self.plan_tree.bind("<Configure>", _on_plan_resize)

        # 우클릭 메뉴
        self.plan_menu = tk.Menu(self.plan_tree, tearoff=0)
        self.plan_menu.add_command(label="선택 항목 삭제", command=self._delete_selected_plan)
        self.plan_tree.bind("<Button-3>", self._show_plan_menu)

    # ═══════════════════════════════════════════
    # 3. 액션 바
    # ═══════════════════════════════════════════
    def _build_action_bar(self):
        action_frame = tk.Frame(self.scroll_frame, bg=COLORS["bg"])
        action_frame.pack(fill=tk.X, padx=5, pady=5)

        left = tk.Frame(action_frame, bg=COLORS["bg"])
        left.pack(side=tk.LEFT)

        # 안전재고 반영 체크박스
        self.safety_stock_var = tk.BooleanVar(value=False)
        self.safety_check = tk.Checkbutton(
            left, text="안전재고 반영 (발주 수량에 안전재고 포함)",
            variable=self.safety_stock_var,
            bg=COLORS["bg"], fg=COLORS["text"],
            font=(FONT_FAMILY, FONT_SIZES["small"]),
            activebackground=COLORS["bg"],
            selectcolor="white")
        self.safety_check.pack(side=tk.LEFT, padx=(5, 20))

        right = tk.Frame(action_frame, bg=COLORS["bg"])
        right.pack(side=tk.RIGHT)

        # 소요량 계산 버튼
        self.calc_btn = tk.Button(right, text="🔍 소요량 계산",
                                   font=(FONT_FAMILY, FONT_SIZES["body_large"], "bold"),
                                   bg=COLORS["success"], fg="white",
                                   padx=20, pady=8, cursor="hand2",
                                   command=self._calculate_mrp)
        self.calc_btn.pack(side=tk.LEFT, padx=5)

        # 엑셀 내보내기 버튼
        self.export_btn = tk.Button(right, text="📥 엑셀 내보내기",
                                     font=(FONT_FAMILY, FONT_SIZES["body_large"], "bold"),
                                     bg=COLORS["info"], fg="white",
                                     padx=20, pady=8, cursor="hand2",
                                     state="disabled",
                                     command=self._export_excel)
        self.export_btn.pack(side=tk.LEFT, padx=5)

    # ═══════════════════════════════════════════
    # 자동완성 검색 로직
    # ═══════════════════════════════════════════
    def _load_products_cache(self):
        """제품 목록 캐시 로드 (백그라운드)"""
        def load():
            try:
                products = self.app.db.get_all_products()
                self.products_cache = products
            except:
                self.products_cache = []
        threading.Thread(target=load, daemon=True).start()

    def _on_search_focus(self, event=None):
        """검색창 포커스 시 전체 목록 표시"""
        text = self.search_var.get().strip()
        if not text:
            self._show_autocomplete(self.products_cache[:10])

    def _on_search_change(self, event=None):
        """검색어 변경 시 자동완성 업데이트"""
        # ESC 키로 닫기
        if event and event.keysym == "Escape":
            self._hide_autocomplete()
            return

        # 아래/위 화살표로 리스트 탐색
        if event and event.keysym == "Down":
            self.autocomplete_listbox.focus_set()
            if self.autocomplete_listbox.size() > 0:
                self.autocomplete_listbox.selection_set(0)
            return

        text = self.search_var.get().strip().lower()
        if not text:
            self._show_autocomplete(self.products_cache[:10])
            return

        # 제품코드 또는 제품명으로 필터링
        filtered = []
        for p in self.products_cache:
            code = str(p.get("제품코드", "")).lower()
            name = str(p.get("제품명", "")).lower()
            if text in code or text in name:
                filtered.append(p)

        self._show_autocomplete(filtered[:10])

    def _show_autocomplete(self, products):
        """자동완성 리스트 표시"""
        self.autocomplete_listbox.delete(0, tk.END)

        if not products:
            self._hide_autocomplete()
            return

        for p in products:
            code = str(p.get("제품코드", ""))
            name = str(p.get("제품명", ""))
            stock = int(p.get("현재재고", 0))
            self.autocomplete_listbox.insert(tk.END, f"{code}  |  {name}  (재고: {stock}개)")

        # 리스트 위치 설정 (검색창 바로 아래)
        self.autocomplete_frame.place(
            in_=self.search_entry,
            x=0, y=self.search_entry.winfo_height(),
            width=self.search_entry.winfo_width() + 150)

    def _hide_autocomplete(self):
        """자동완성 리스트 숨기기"""
        self.autocomplete_frame.place_forget()

    def _on_autocomplete_select(self, event=None):
        """자동완성 항목 선택"""
        selection = self.autocomplete_listbox.curselection()
        if not selection:
            return

        idx = selection[0]
        text = self.search_var.get().strip().lower()

        # 필터링된 목록에서 선택
        if text:
            filtered = [p for p in self.products_cache
                        if text in str(p.get("제품코드", "")).lower()
                        or text in str(p.get("제품명", "")).lower()]
        else:
            filtered = self.products_cache[:10]

        if idx < len(filtered):
            product = filtered[idx]
            self.selected_product_id = str(product["제품코드"])
            self.selected_product_name = str(product["제품명"])

            self.search_var.set(f"{self.selected_product_id} - {self.selected_product_name}")
            self._hide_autocomplete()

            # 제품 정보 표시
            stock = int(product.get("현재재고", 0))
            self.product_info_label.configure(
                text=f"현재재고: {stock}개")

            # 문제 3 수정: 제품 선택 시 get_max_producible() 개별 API 호출 제거
            # 최대 생산 가능 수량은 '소요량 계산' 버튼 클릭 시 일괄 계산됩니다.
            self.max_prod_label.configure(
                text="※ 최대 생산 가능 수량은 '소요량 계산' 후 확인하세요.",
                fg=COLORS["text_secondary"])

            # 수량 입력으로 포커스 이동
            self.qty_entry.focus_set()

    # ═══════════════════════════════════════════
    # 생산 계획 관리
    # ═══════════════════════════════════════════
    def _add_to_plan(self):
        """생산 계획에 제품 추가"""
        if not self.selected_product_id:
            messagebox.showwarning("알림", "제품을 먼저 검색하여 선택해 주세요.")
            self.search_entry.focus_set()
            return

        qty_text = self.qty_entry.get().strip()
        if not qty_text:
            messagebox.showwarning("알림", "목표 수량을 입력해 주세요.")
            self.qty_entry.focus_set()
            return

        try:
            target_qty = int(qty_text)
            if target_qty <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("알림", "목표 수량은 1 이상의 정수를 입력해 주세요.")
            self.qty_entry.focus_set()
            return

        # 중복 체크
        for plan in self.production_plan:
            if plan["product_id"] == self.selected_product_id:
                if messagebox.askyesno("중복 제품",
                                        f"{self.selected_product_name}이(가) 이미 목록에 있습니다.\n"
                                        f"기존 수량을 {target_qty}개로 변경하시겠습니까?"):
                    plan["target_qty"] = target_qty
                    self._refresh_plan_tree()
                return

        self.production_plan.append({
            "product_id": self.selected_product_id,
            "product_name": self.selected_product_name,
            "target_qty": target_qty,
        })

        self._refresh_plan_tree()

        # 입력 초기화
        self.search_var.set("")
        self.qty_entry.delete(0, tk.END)
        self.selected_product_id = None
        self.selected_product_name = None
        self.product_info_label.configure(text="")
        self.max_prod_label.configure(text="")
        self.search_entry.focus_set()

    def _refresh_plan_tree(self):
        """생산 계획 트리뷰 새로고침
        문제 3 수정: 계획 추가/삭제 시마다 get_max_producible()을 제품 수만큼
        반복 호출하던 구조를 제거합니다.
        → 최대생산가능/병목부품은 '소요량 계산' 버튼 클릭 후에만 표시됩니다.
          (API 호출 횟수: 제품 N개 × 2회 → 0회 로 감소)
        """
        self.plan_tree.delete(*self.plan_tree.get_children())

        products_map = {str(p["제품코드"]): p for p in self.products_cache}

        for i, plan in enumerate(self.production_plan, 1):
            pid           = plan["product_id"]
            product       = products_map.get(pid)
            current_stock = int(product["현재재고"]) if product else 0
            need          = max(0, plan["target_qty"] - current_stock)

            tag = "need" if need > 0 else "ok"
            self.plan_tree.insert("", tk.END, values=(
                i, pid, plan["product_name"],
                f"{current_stock}개", f"{plan['target_qty']}개",
                f"{need}개" if need > 0 else "충분",
                "계산 전",   # 최대생산가능 — 소요량 계산 후 갱신
                "계산 전"    # 병목부품     — 소요량 계산 후 갱신
            ), tags=(tag,))

        self.plan_tree.tag_configure("need", foreground=COLORS["danger"])
        self.plan_tree.tag_configure("ok",   foreground=COLORS["success"])

        self.plan_count_label.configure(text=f"({len(self.production_plan)}개 제품)")

    def _delete_selected_plan(self):
        """선택된 계획 항목 삭제"""
        selected = self.plan_tree.selection()
        if not selected:
            return
        values = self.plan_tree.item(selected[0])["values"]
        product_id = str(values[1])

        self.production_plan = [p for p in self.production_plan if p["product_id"] != product_id]
        self._refresh_plan_tree()

    def _show_plan_menu(self, event):
        """우클릭 메뉴"""
        item = self.plan_tree.identify_row(event.y)
        if item:
            self.plan_tree.selection_set(item)
            self.plan_menu.post(event.x_root, event.y_root)

    def _clear_plan(self):
        """전체 계획 삭제"""
        if not self.production_plan:
            return
        if messagebox.askyesno("확인", "생산 계획 목록을 전체 삭제하시겠습니까?"):
            self.production_plan = []
            self._refresh_plan_tree()
            # 결과 영역도 초기화
            for w in self.result_frame.winfo_children():
                w.destroy()
            self.export_btn.configure(state="disabled")

    # ═══════════════════════════════════════════
    # MRP 계산
    # ═══════════════════════════════════════════
    def _calculate_mrp(self):
        """소요량 계산 실행"""
        if not self.production_plan:
            messagebox.showwarning("알림", "생산 계획을 먼저 입력해 주세요.")
            return

        self.calc_btn.configure(state="disabled", text="계산 중...")

        include_safety = self.safety_stock_var.get()

        def process():
            try:
                result = self.app.db.calculate_mrp(
                    self.production_plan,
                    include_safety_stock=include_safety
                )
                self.app.root.after(0, lambda: self._show_mrp_result(result))
            except Exception as e:
                self.app.root.after(0, lambda: self._show_mrp_error(str(e)))

        threading.Thread(target=process, daemon=True).start()

    def _show_mrp_error(self, error_msg):
        """계산 오류 표시"""
        self.calc_btn.configure(state="normal", text="🔍 소요량 계산")
        messagebox.showerror("계산 오류", f"MRP 계산 중 오류가 발생했습니다:\n{error_msg}")

    def _show_mrp_result(self, result):
        """MRP 계산 결과 표시"""
        self.calc_btn.configure(state="normal", text="🔍 소요량 계산")
        self.mrp_result = result
        self.export_btn.configure(state="normal")

        # 결과 영역 초기화
        for w in self.result_frame.winfo_children():
            w.destroy()

        # ── 생산 계획 요약 업데이트 ──
        self._refresh_plan_tree_with_result(result["plan_summary"])

        # ── 요약 통계 카드 ──
        stats_frame = tk.Frame(self.result_frame, bg=COLORS["bg"])
        stats_frame.pack(fill=tk.X, pady=(5, 5))

        total_parts = len(result["parts_requirement"])
        order_items = result["total_order_items"]
        order_qty = result["total_order_qty"]
        sufficient = total_parts - order_items

        stats = [
            ("총 소요 부품", f"{total_parts}종", COLORS["info"]),
            ("재고 충분", f"{sufficient}종", COLORS["success"]),
            ("발주 필요", f"{order_items}종", COLORS["danger"] if order_items > 0 else COLORS["success"]),
            ("총 발주 수량", f"{order_qty:,}개", COLORS["warning"] if order_qty > 0 else COLORS["success"]),
        ]

        for i, (label, value, color) in enumerate(stats):
            stat_card = tk.Frame(stats_frame, bg="white",
                                  highlightbackground=color, highlightthickness=2)
            stat_card.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=3)
            tk.Label(stat_card, text=label, bg="white", fg=COLORS["text_secondary"],
                     font=(FONT_FAMILY, FONT_SIZES["small"])).pack(pady=(8, 2))
            tk.Label(stat_card, text=value, bg="white", fg=color,
                     font=(FONT_FAMILY, FONT_SIZES["stat"], "bold")).pack(pady=(0, 8))

        # ── 발주 리스트 카드 ──
        self._build_requirement_table(result)

    def _refresh_plan_tree_with_result(self, plan_summary):
        """MRP 계산 결과로 생산 계획 트리 업데이트
        문제 3 수정: 최대생산가능/병목부품은 여기서 한 번만 채워집니다.
        calculate_mrp() 안에서 이미 일괄 계산된 결과를 받아 표시하므로
        추가 API 호출이 전혀 발생하지 않습니다.
        """
        self.plan_tree.delete(*self.plan_tree.get_children())

        for i, item in enumerate(plan_summary, 1):
            need = item["need_to_produce"]
            tag  = "need" if need > 0 else "ok"
            self.plan_tree.insert("", tk.END, values=(
                i, item["product_id"], item["product_name"],
                f"{item['current_stock']}개", f"{item['target_qty']}개",
                f"{need}개" if need > 0 else "충분",
                f"{item['max_producible']}개",
                item["bottleneck"] if need > 0 else "-"
            ), tags=(tag,))

        self.plan_tree.tag_configure("need", foreground=COLORS["danger"])
        self.plan_tree.tag_configure("ok",   foreground=COLORS["success"])

    def _build_requirement_table(self, result):
        """발주 리스트 테이블 구성"""
        card = tk.Frame(self.result_frame, bg=COLORS["card_bg"],
                        highlightbackground=COLORS["border"], highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True, padx=0, pady=(5, 10))

        header = tk.Frame(card, bg=COLORS["card_bg"])
        header.pack(fill=tk.X, padx=15, pady=(12, 5))

        tk.Label(header, text="부품 소요량 및 발주 리스트",
                 bg=COLORS["card_bg"], fg=COLORS["text"],
                 font=(FONT_FAMILY, FONT_SIZES["heading"], "bold")).pack(side=tk.LEFT)

        safety_text = " (안전재고 반영)" if self.safety_stock_var.get() else ""
        tk.Label(header, text=safety_text,
                 bg=COLORS["card_bg"], fg=COLORS["info"],
                 font=(FONT_FAMILY, FONT_SIZES["small"])).pack(side=tk.LEFT, padx=5)

        # 필터 옵션
        filter_frame = tk.Frame(card, bg=COLORS["card_bg"])
        filter_frame.pack(fill=tk.X, padx=15, pady=(0, 5))

        self.filter_var = tk.StringVar(value="all")
        tk.Radiobutton(filter_frame, text="전체 부품", variable=self.filter_var,
                       value="all", bg=COLORS["card_bg"],
                       font=(FONT_FAMILY, FONT_SIZES["small"]),
                       command=lambda: self._apply_filter(result)).pack(side=tk.LEFT, padx=(0, 15))
        tk.Radiobutton(filter_frame, text="발주 필요 부품만", variable=self.filter_var,
                       value="order_only", bg=COLORS["card_bg"],
                       font=(FONT_FAMILY, FONT_SIZES["small"]),
                       command=lambda: self._apply_filter(result)).pack(side=tk.LEFT)

        # 트리뷰
        tree_frame = tk.Frame(card, bg=COLORS["card_bg"])
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 12))

        columns = ("no", "part_id", "part_name", "supplier", "unit",
                    "total_required", "current_stock", "safety_stock",
                    "shortage", "status")
        self.req_tree = ttk.Treeview(tree_frame, columns=columns,
                                      show="headings", height=15)

        col_config = [
            ("no", "No", 40, "center"),
            ("part_id", "품번", 100, "center"),
            ("part_name", "부품명", 180, "w"),
            ("supplier", "업체명", 100, "w"),
            ("unit", "단위", 50, "center"),
            ("total_required", "총소요량", 80, "center"),
            ("current_stock", "현재재고", 80, "center"),
            ("safety_stock", "안전재고", 80, "center"),
            ("shortage", "부족수량", 80, "center"),
            ("status", "발주필요", 80, "center"),
        ]

        for col_id, heading, width, anchor in col_config:
            self.req_tree.heading(col_id, text=heading)
            self.req_tree.column(col_id, width=width, anchor=anchor)

        scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.req_tree.yview)
        self.req_tree.configure(yscrollcommand=scrollbar_y.set)
        self.req_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)

        # 반응형: 부품명 컬럼 자동 조정
        def _on_req_resize(event):
            total = self.req_tree.winfo_width()
            fixed = 40 + 100 + 100 + 50 + 80 + 80 + 80 + 80 + 80 + 20  # 부품명 제외 고정 열
            remaining = max(120, total - fixed)
            self.req_tree.column("part_name", width=remaining)
        self.req_tree.bind("<Configure>", _on_req_resize)

        self.req_tree.tag_configure("order", background="#fef2f2", foreground=COLORS["danger"])
        self.req_tree.tag_configure("ok", foreground=COLORS["success"])

        self._apply_filter(result)

    def _apply_filter(self, result):
        """필터 적용"""
        self.req_tree.delete(*self.req_tree.get_children())

        filter_mode = self.filter_var.get()
        parts = result["parts_requirement"]

        for i, part in enumerate(parts, 1):
            if filter_mode == "order_only" and not part["order_needed"]:
                continue

            tag = "order" if part["order_needed"] else "ok"
            status = f"발주 {int(part['shortage'])}개" if part["order_needed"] else "충분"

            self.req_tree.insert("", tk.END, values=(
                i, part["part_id"], part["part_name"],
                part["supplier"], part["unit"],
                part["total_required"], part["current_stock"],
                part["safety_stock"], int(part["shortage"]),
                status
            ), tags=(tag,))

    # ═══════════════════════════════════════════
    # 엑셀 내보내기
    # ═══════════════════════════════════════════
    def _export_excel(self):
        """MRP 결과를 엑셀로 내보내기"""
        if not self.mrp_result:
            messagebox.showwarning("알림", "먼저 소요량 계산을 실행해 주세요.")
            return

        # 저장 경로 선택
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"MRP_발주리스트_{now}.xlsx"

        file_path = filedialog.asksaveasfilename(
            title="MRP 발주 리스트 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
            initialfile=default_name,
        )

        if not file_path:
            return

        try:
            self._generate_excel(file_path)
            messagebox.showinfo("저장 완료",
                                f"MRP 발주 리스트가 저장되었습니다.\n\n{file_path}")
            # 파일 열기
            os.startfile(file_path)
        except Exception as e:
            messagebox.showerror("저장 오류", f"엑셀 저장 중 오류가 발생했습니다:\n{e}")

    def _generate_excel(self, file_path):
        """엑셀 파일 생성"""
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

        wb = openpyxl.Workbook()

        # ── Sheet 1: 생산 계획 요약 ──
        ws1 = wb.active
        ws1.title = "생산계획"

        header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        header_font = Font(name="맑은 고딕", bold=True, color="FFFFFF", size=11)
        normal_font = Font(name="맑은 고딕", size=10)
        title_font = Font(name="맑은 고딕", bold=True, size=14)
        danger_font = Font(name="맑은 고딕", bold=True, size=10, color="DC2626")
        success_font = Font(name="맑은 고딕", size=10, color="16A34A")
        thin_border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin"))

        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        ws1.merge_cells("A1:H1")
        ws1["A1"] = f"자재소요계획 (MRP) - {now}"
        ws1["A1"].font = title_font
        ws1.row_dimensions[1].height = 30

        safety_text = "안전재고 반영: 예" if self.safety_stock_var.get() else "안전재고 반영: 아니오"
        ws1.merge_cells("A2:H2")
        ws1["A2"] = safety_text
        ws1["A2"].font = Font(name="맑은 고딕", size=9, color="666666")

        headers1 = ["No", "제품코드", "제품명", "현재재고", "목표수량",
                     "추가생산필요", "최대생산가능", "병목부품"]
        for col, h in enumerate(headers1, 1):
            cell = ws1.cell(row=4, column=col, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

        for i, item in enumerate(self.mrp_result["plan_summary"], 1):
            row = i + 4
            values = [i, item["product_id"], item["product_name"],
                      item["current_stock"], item["target_qty"],
                      item["need_to_produce"], item["max_producible"],
                      item["bottleneck"]]
            for col, val in enumerate(values, 1):
                cell = ws1.cell(row=row, column=col, value=val)
                cell.font = normal_font
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center")
                if col == 6 and item["need_to_produce"] > 0:
                    cell.font = danger_font

        # 열 너비
        widths1 = [6, 15, 25, 12, 12, 14, 14, 25]
        for i, w in enumerate(widths1, 1):
            ws1.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

        # ── Sheet 2: 전체 발주 리스트 ──
        ws2 = wb.create_sheet("발주리스트_전체")

        ws2.merge_cells("A1:J1")
        ws2["A1"] = f"부품 소요량 및 발주 리스트 - {now}"
        ws2["A1"].font = title_font
        ws2.row_dimensions[1].height = 30

        summary = (f"총 소요 부품: {len(self.mrp_result['parts_requirement'])}종  |  "
                   f"발주 필요: {self.mrp_result['total_order_items']}종  |  "
                   f"총 발주 수량: {self.mrp_result['total_order_qty']:,}개")
        ws2.merge_cells("A2:J2")
        ws2["A2"] = summary
        ws2["A2"].font = Font(name="맑은 고딕", size=9, color="666666")

        headers2 = ["No", "품번", "부품명", "업체명", "단위",
                     "총소요량", "현재재고", "안전재고", "부족수량", "발주필요"]
        for col, h in enumerate(headers2, 1):
            cell = ws2.cell(row=4, column=col, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

        order_fill = PatternFill(start_color="FEF2F2", end_color="FEF2F2", fill_type="solid")

        for i, part in enumerate(self.mrp_result["parts_requirement"], 1):
            row = i + 4
            status = f"발주 {int(part['shortage'])}개" if part["order_needed"] else "충분"
            values = [i, part["part_id"], part["part_name"], part["supplier"],
                      part["unit"], part["total_required"], part["current_stock"],
                      part["safety_stock"], int(part["shortage"]), status]
            for col, val in enumerate(values, 1):
                cell = ws2.cell(row=row, column=col, value=val)
                cell.font = normal_font
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center")
                if part["order_needed"]:
                    cell.fill = order_fill
                    if col in (9, 10):
                        cell.font = danger_font

        widths2 = [6, 15, 30, 15, 8, 12, 12, 12, 12, 14]
        for i, w in enumerate(widths2, 1):
            ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

        # ── Sheet 3: 발주 필요 부품만 (업체별 그룹핑) ──
        ws3 = wb.create_sheet("발주필요_업체별")

        ws3.merge_cells("A1:H1")
        ws3["A1"] = f"발주 필요 부품 (업체별) - {now}"
        ws3["A1"].font = title_font
        ws3.row_dimensions[1].height = 30

        # 업체별 그룹핑
        order_parts = [p for p in self.mrp_result["parts_requirement"] if p["order_needed"]]
        suppliers = {}
        for part in order_parts:
            supplier = part["supplier"] if part["supplier"] else "(업체 미지정)"
            if supplier not in suppliers:
                suppliers[supplier] = []
            suppliers[supplier].append(part)

        current_row = 3
        for supplier, parts in sorted(suppliers.items()):
            # 업체명 헤더
            ws3.merge_cells(f"A{current_row}:H{current_row}")
            cell = ws3.cell(row=current_row, column=1, value=f"▶ {supplier} ({len(parts)}종)")
            cell.font = Font(name="맑은 고딕", bold=True, size=11, color="1E293B")
            cell.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
            current_row += 1

            # 컬럼 헤더
            sub_headers = ["No", "품번", "부품명", "단위", "총소요량",
                           "현재재고", "부족수량", "비고"]
            for col, h in enumerate(sub_headers, 1):
                cell = ws3.cell(row=current_row, column=col, value=h)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center")
                cell.border = thin_border
            current_row += 1

            # 데이터
            for j, part in enumerate(parts, 1):
                values = [j, part["part_id"], part["part_name"], part["unit"],
                          part["total_required"], part["current_stock"],
                          int(part["shortage"]), ""]
                for col, val in enumerate(values, 1):
                    cell = ws3.cell(row=current_row, column=col, value=val)
                    cell.font = normal_font
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal="center")
                    if col == 7:
                        cell.font = danger_font
                current_row += 1

            current_row += 1  # 업체 간 빈 줄

        widths3 = [6, 15, 30, 8, 12, 12, 12, 20]
        for i, w in enumerate(widths3, 1):
            ws3.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

        wb.save(file_path)
