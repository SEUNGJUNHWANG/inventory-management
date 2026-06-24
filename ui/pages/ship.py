# -*- coding: utf-8 -*-
"""
재고관리 시스템 - 제품 출고(판매) 페이지
거래처 자동완성 + 실시간 재고/금액 카드
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
from core.constants import COLORS, FONT_FAMILY, FONT_SIZES
from core.auth import Session
from ui.widget_utils import flash_btn

# AutocompleteEntry 재사용
from ui.pages.transactions import AutocompleteEntry


def _fmt(n):
    """숫자를 천단위 콤마 문자열로 변환"""
    try:
        return f"{int(float(n)):,}"
    except Exception:
        return str(n)


class ShipPage:
    def __init__(self, app):
        self.app = app
        self._selected_product  = None
        self._selected_customer = None
        self._btn_ship          = None
        self._product_items     = []   # [(제품코드, 제품명), ...]
        self._customer_items    = []   # [(거래처코드, 거래처명), ...]

    # ──────────────────────────────────────────
    # 렌더링
    # ──────────────────────────────────────────
    def render(self):
        scroll_frame = self.app._create_scrollable_frame()

        # ── 헤더 ──
        header = tk.Frame(scroll_frame, bg=COLORS["bg"])
        header.pack(fill=tk.X, padx=5, pady=(0, 12))
        tk.Label(header, text="🚚 제품 출고(판매)", bg=COLORS["bg"],
                 fg=COLORS["text"],
                 font=(FONT_FAMILY, FONT_SIZES["title"], "bold")).pack(side=tk.LEFT)

        # ── 메인 카드 ──
        card = tk.Frame(scroll_frame, bg=COLORS["card_bg"],
                        highlightbackground=COLORS["border"], highlightthickness=1)
        card.pack(fill=tk.X, padx=5, pady=(0, 10))

        tk.Label(card,
                 text="제품코드와 거래처를 선택한 뒤 수량을 입력하고 [출고 처리] 버튼을 누르세요.",
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"],
                 font=(FONT_FAMILY, FONT_SIZES["small"])
                 ).pack(anchor="w", padx=20, pady=(14, 10))

        # ── 입력 폼 ──
        form = tk.Frame(card, bg=COLORS["card_bg"])
        form.pack(fill=tk.X, padx=20, pady=4)

        lbl_kw = dict(bg=COLORS["card_bg"],
                      font=(FONT_FAMILY, FONT_SIZES["small"], "bold"))
        entry_kw = dict(font=(FONT_FAMILY, FONT_SIZES["body"]))

        # 제품코드
        tk.Label(form, text="제품코드:", **lbl_kw).grid(
            row=0, column=0, sticky="e", padx=(0, 8), pady=7)
        self.ent_product = AutocompleteEntry(
            form, on_select=self._on_product_selected, width=26, **entry_kw)
        self.ent_product.grid(row=0, column=1, sticky="w", pady=7)
        self.ent_product.bind("<Return>", lambda e: self.ent_customer.focus_set())

        self._lbl_pname = tk.Label(form, text="", bg=COLORS["card_bg"],
                                   fg=COLORS["text_secondary"],
                                   font=(FONT_FAMILY, FONT_SIZES["small"]))
        self._lbl_pname.grid(row=0, column=2, sticky="w", padx=8)

        # 거래처
        tk.Label(form, text="거래처:", **lbl_kw).grid(
            row=1, column=0, sticky="e", padx=(0, 8), pady=7)
        self.ent_customer = AutocompleteEntry(
            form, on_select=self._on_customer_selected, width=26, **entry_kw)
        self.ent_customer.grid(row=1, column=1, sticky="w", pady=7)
        self.ent_customer.bind("<Return>", lambda e: self.ent_qty.focus_set())

        self._lbl_cname = tk.Label(form, text="", bg=COLORS["card_bg"],
                                   fg=COLORS["text_secondary"],
                                   font=(FONT_FAMILY, FONT_SIZES["small"]))
        self._lbl_cname.grid(row=1, column=2, sticky="w", padx=8)

        # 수량
        tk.Label(form, text="수량:", **lbl_kw).grid(
            row=2, column=0, sticky="e", padx=(0, 8), pady=7)
        qty_frame = tk.Frame(form, bg=COLORS["card_bg"])
        qty_frame.grid(row=2, column=1, sticky="w", pady=7)
        self.ent_qty = tk.Entry(qty_frame, width=12, **entry_kw)
        self.ent_qty.pack(side=tk.LEFT)
        self.ent_qty.bind("<KeyRelease>", self._on_qty_change)
        self.ent_qty.bind("<Return>", lambda e: self.ent_price.focus_set())

        self._lbl_stock_warn = tk.Label(qty_frame, text="", bg=COLORS["card_bg"],
                                        fg=COLORS["danger"],
                                        font=(FONT_FAMILY, FONT_SIZES["small"]))
        self._lbl_stock_warn.pack(side=tk.LEFT, padx=8)

        # 단가
        tk.Label(form, text="단가(원):", **lbl_kw).grid(
            row=3, column=0, sticky="e", padx=(0, 8), pady=7)
        price_frame = tk.Frame(form, bg=COLORS["card_bg"])
        price_frame.grid(row=3, column=1, sticky="w", pady=7)
        self.ent_price = tk.Entry(price_frame, width=16, **entry_kw)
        self.ent_price.pack(side=tk.LEFT)
        self.ent_price.bind("<KeyRelease>", self._on_price_change)
        self.ent_price.bind("<Return>", lambda e: self.ent_note.focus_set())
        tk.Label(price_frame, text="※ 판매가 자동입력, 수정 가능",
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"],
                 font=(FONT_FAMILY, FONT_SIZES["tiny"])).pack(side=tk.LEFT, padx=8)

        # 비고
        tk.Label(form, text="비고:", **lbl_kw).grid(
            row=4, column=0, sticky="e", padx=(0, 8), pady=7)
        self.ent_note = tk.Entry(form, width=44,
                                 font=(FONT_FAMILY, FONT_SIZES["small"]))
        self.ent_note.grid(row=4, column=1, columnspan=2, sticky="w", pady=7)
        self.ent_note.bind("<Return>", lambda e: flash_btn(self._btn_ship, self._do_ship) if self._btn_ship else self._do_ship())

        # ── 실시간 요약 카드 ──
        self._summary_frame = tk.Frame(card, bg="#f0fdf4",
                                       highlightbackground="#bbf7d0",
                                       highlightthickness=1)
        # 처음엔 숨김, 수량·단가 입력 시 표시
        self._lbl_summary = tk.Label(self._summary_frame, text="",
                                     bg="#f0fdf4", fg=COLORS["text"],
                                     font=(FONT_FAMILY, FONT_SIZES["body_large"], "bold"),
                                     padx=20, pady=12, justify="left")
        self._lbl_summary.pack(anchor="w")

        # ── 재고 현황 카드 ──
        self._stock_frame = tk.Frame(card, bg="#f8fafc",
                                     highlightbackground=COLORS["border"],
                                     highlightthickness=1)
        self._stock_inner = tk.Frame(self._stock_frame, bg="#f8fafc", padx=18, pady=12)
        self._stock_inner.pack(fill=tk.X)

        top = tk.Frame(self._stock_inner, bg="#f8fafc")
        top.pack(fill=tk.X, pady=(0, 8))
        self._lbl_stock_name  = tk.Label(top, text="", bg="#f8fafc", fg=COLORS["text"],
                                         font=(FONT_FAMILY, FONT_SIZES["heading"], "bold"), anchor="w")
        self._lbl_stock_name.pack(side=tk.LEFT)
        self._lbl_stock_badge = tk.Label(top, text="", bg=COLORS["primary"], fg="white",
                                         font=(FONT_FAMILY, FONT_SIZES["tiny"], "bold"),
                                         padx=10, pady=3)
        self._lbl_stock_badge.pack(side=tk.LEFT, padx=10)

        stats = tk.Frame(self._stock_inner, bg="#f8fafc")
        stats.pack(fill=tk.X)
        self._stat_stock = self._make_stat(stats, "현재 재고", COLORS["primary"])
        self._stat_stock.pack(side=tk.LEFT, padx=(0, 40))
        self._stat_price = self._make_stat(stats, "판매가(원)", COLORS["text_secondary"])
        self._stat_price.pack(side=tk.LEFT, padx=(0, 40))
        self._stat_after = self._make_stat(stats, "출고 후 잔여", COLORS["warning"])
        self._stat_after.pack(side=tk.LEFT)

        # ── 출고 버튼 ──
        self._btn_frame = tk.Frame(card, bg=COLORS["card_bg"])
        self._btn_frame.pack(fill=tk.X, padx=20, pady=(12, 16))
        btn_frame = self._btn_frame

        if Session.has_write("ship"):
            self._btn_ship = tk.Button(
                btn_frame, text="  출고 처리  ",
                font=(FONT_FAMILY, FONT_SIZES["body_large"], "bold"),
                bg="#0ea5e9", fg="white",
                activebackground="#0284c7", activeforeground="white",
                relief="flat", cursor="hand2", padx=24, pady=10,
                command=self._do_ship,
            )
            self._btn_ship.pack(side=tk.LEFT)

        self._lbl_result = tk.Label(
            card, text="", bg=COLORS["card_bg"],
            font=(FONT_FAMILY, FONT_SIZES["body_large"]),
            wraplength=700, justify="left",
        )
        self._lbl_result.pack(anchor="w", padx=20, pady=(0, 16))

        # ── 최근 판매 이력 (간략) ──
        hist_label = tk.Frame(scroll_frame, bg=COLORS["bg"])
        hist_label.pack(fill=tk.X, padx=5, pady=(8, 4))
        tk.Label(hist_label, text="최근 판매 이력", bg=COLORS["bg"],
                 fg=COLORS["text"],
                 font=(FONT_FAMILY, FONT_SIZES["heading"], "bold")).pack(side=tk.LEFT)

        hist_card = tk.Frame(scroll_frame, bg=COLORS["card_bg"],
                             highlightbackground=COLORS["border"], highlightthickness=1)
        hist_card.pack(fill=tk.BOTH, expand=True, padx=5)

        hist_cols = ("일시", "제품명", "거래처명", "수량", "단가", "금액", "잔여재고")
        self._hist_tree = ttk.Treeview(hist_card, columns=hist_cols,
                                       show="headings", height=10)
        hist_widths = {"일시": 150, "제품명": 200, "거래처명": 160,
                       "수량": 70, "단가": 100, "금액": 110, "잔여재고": 80}
        for col in hist_cols:
            self._hist_tree.heading(col, text=col)
            self._hist_tree.column(col, width=hist_widths.get(col, 100), anchor="center")

        hist_vsb = ttk.Scrollbar(hist_card, orient="vertical",
                                  command=self._hist_tree.yview)
        self._hist_tree.configure(yscrollcommand=hist_vsb.set)
        self._hist_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        hist_vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # 캐시 로드
        self._load_caches()

    # ──────────────────────────────────────────
    # stat 위젯 생성 헬퍼
    # ──────────────────────────────────────────
    def _make_stat(self, parent, label, color):
        f = tk.Frame(parent, bg="#f8fafc")
        tk.Label(f, text=label, bg="#f8fafc", fg=COLORS["text_secondary"],
                 font=(FONT_FAMILY, FONT_SIZES["tiny"])).pack(anchor="w")
        val = tk.Label(f, text="—", bg="#f8fafc", fg=color,
                       font=(FONT_FAMILY, FONT_SIZES["stat"], "bold"))
        val.pack(anchor="w")
        f._val = val
        return f

    # ──────────────────────────────────────────
    # 캐시 로드 (자동완성 목록)
    # ──────────────────────────────────────────
    def _load_caches(self):
        def load():
            try:
                products  = self.app.db.get_all_products()
                customers = self.app.db.get_all_customers()
                p_items = [(p.get("제품코드",""), p.get("제품명","")) for p in products]
                c_items = [(c.get("거래처코드",""), c.get("거래처명","")) for c in customers]
                sales   = self.app.db.get_all_sales()
                self.app.root.after(0, lambda: self._apply_caches(p_items, c_items, sales))
            except Exception as e:
                msg = str(e)
                self.app.root.after(0, lambda: messagebox.showerror("오류", msg))

        threading.Thread(target=load, daemon=True).start()

    def _apply_caches(self, p_items, c_items, sales):
        self._product_items  = p_items
        self._customer_items = c_items
        self.ent_product.set_items(p_items)
        self.ent_customer.set_items(c_items)
        self._render_history(sales)

    # ──────────────────────────────────────────
    # 최근 판매 이력 렌더링
    # ──────────────────────────────────────────
    def _render_history(self, sales):
        self._hist_tree.delete(*self._hist_tree.get_children())
        recent = list(reversed(sales[-50:]))   # 최근 50건, 최신순
        for r in recent:
            self._hist_tree.insert("", "end", values=(
                r.get("일시", ""),
                r.get("제품명", ""),
                r.get("거래처명", ""),
                _fmt(r.get("수량", 0)),
                _fmt(r.get("단가", 0)),
                _fmt(r.get("금액", 0)),
                _fmt(r.get("잔여재고", 0)),
            ))

    # ──────────────────────────────────────────
    # 제품 선택 콜백
    # ──────────────────────────────────────────
    def _on_product_selected(self, code, name):
        if not code:
            self._selected_product = None
            self._lbl_pname.configure(text="")
            self._stock_frame.pack_forget()
            return

        self._lbl_pname.configure(text="")
        def lookup():
            prod = self.app.db.get_product_by_id(code)
            if prod:
                self.app.root.after(0, lambda p=prod: self._show_product(p))
            else:
                self._selected_product = None
                self.app.root.after(0, lambda: self._stock_frame.pack_forget())
        threading.Thread(target=lookup, daemon=True).start()

    def _show_product(self, prod):
        self._selected_product = prod
        name    = prod.get("제품명", "")
        current = int(prod.get("현재재고", 0))
        sp      = prod.get("판매가", 0)

        self._lbl_pname.configure(text=name)
        self._lbl_stock_name.configure(text=name)
        badge_text = "정상 재고" if current > 0 else "재고 없음"
        badge_bg   = COLORS["success"] if current > 0 else COLORS["danger"]
        self._lbl_stock_badge.configure(text=badge_text, bg=badge_bg)
        self._stat_stock._val.configure(text=f"{current:,}개")
        self._stat_price._val.configure(
            text=f"{int(sp):,}" if sp else "—")
        self._stat_after._val.configure(text="—")

        # 단가 자동 채우기
        self.ent_price.delete(0, tk.END)
        if sp:
            self.ent_price.insert(0, str(int(float(sp))))

        self._stock_frame.pack(fill=tk.X, padx=20, pady=(0, 12), before=self._btn_frame)
        self._update_summary()

    # ──────────────────────────────────────────
    # 거래처 선택 콜백
    # ──────────────────────────────────────────
    def _on_customer_selected(self, code, name):
        if not code:
            self._selected_customer = None
            self._lbl_cname.configure(text="")
            return
        self._selected_customer = {"거래처코드": code, "거래처명": name}
        self._lbl_cname.configure(text=name)

    # ──────────────────────────────────────────
    # 수량/단가 변경 시 실시간 요약 갱신
    # ──────────────────────────────────────────
    def _on_qty_change(self, event=None):
        self._update_summary()

    def _on_price_change(self, event=None):
        self._update_summary()

    def _update_summary(self):
        if not self._selected_product:
            self._summary_frame.pack_forget()
            return

        current = int(self._selected_product.get("현재재고", 0))
        qty_str   = self.ent_qty.get().strip()
        price_str = self.ent_price.get().strip().replace(",", "")

        try:
            qty = int(qty_str) if qty_str else 0
        except ValueError:
            qty = 0
        try:
            price = float(price_str) if price_str else 0
        except ValueError:
            price = 0

        # 재고 초과 경고
        if qty > 0 and qty > current:
            self._lbl_stock_warn.configure(
                text=f"재고 초과! (가능: {current:,}개)")
        else:
            self._lbl_stock_warn.configure(text="")

        # 출고 후 잔여
        after = current - qty if qty > 0 else None
        if after is not None and after >= 0:
            self._stat_after._val.configure(text=f"{after:,}개",
                fg=COLORS["success"] if after > 0 else COLORS["danger"])
        else:
            self._stat_after._val.configure(text="—", fg=COLORS["warning"])

        # 요약 카드
        if qty > 0 and price > 0 and qty <= current:
            amount = qty * price
            summary = (
                f"수량 {qty:,}개   x   단가 {int(price):,}원"
                f"   =   합계 {int(amount):,}원\n"
                f"출고 후 잔여재고: {after:,}개"
            )
            self._lbl_summary.configure(text=summary)
            self._summary_frame.pack(fill=tk.X, padx=20, pady=(0, 8),
                                     before=self._btn_frame)
        else:
            self._summary_frame.pack_forget()

    # ──────────────────────────────────────────
    # 출고 처리
    # ──────────────────────────────────────────
    def _do_ship(self):
        product_id = self.ent_product.get().strip()
        qty_str    = self.ent_qty.get().strip()
        price_str  = self.ent_price.get().strip().replace(",", "")
        note       = self.ent_note.get().strip()

        if not product_id:
            messagebox.showwarning("입력 오류", "제품코드를 선택해 주세요.")
            return
        if not self._selected_customer:
            messagebox.showwarning("입력 오류", "거래처를 선택해 주세요.")
            return
        if not qty_str:
            messagebox.showwarning("입력 오류", "수량을 입력해 주세요.")
            return

        try:
            qty = int(qty_str)
            if qty <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("입력 오류", "수량은 1 이상의 정수를 입력해 주세요.")
            return

        try:
            price = float(price_str) if price_str else 0
            if price < 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("입력 오류", "단가를 올바르게 입력해 주세요.")
            return

        # 재고 초과 확인
        if self._selected_product:
            current = int(self._selected_product.get("현재재고", 0))
            if qty > current:
                messagebox.showwarning("재고 부족",
                    f"현재재고({current:,}개)보다 출고수량({qty:,}개)이 많습니다.")
                return

        c = self._selected_customer
        cust_code = c.get("거래처코드", "")
        cust_name = c.get("거래처명", "")

        # 최종 확인
        if not messagebox.askyesno("출고 확인",
            f"[{product_id}] {qty:,}개를\n"
            f"'{cust_name}'에 단가 {int(price):,}원으로 출고하시겠습니까?"):
            return

        if self._btn_ship:
            self._btn_ship.configure(text="처리 중...", state="disabled")
        self._lbl_result.configure(text="")

        def process():
            success, msg = self.app.db.ship_product(
                product_id, qty, cust_code, cust_name, price, note)
            self.app.root.after(0, lambda: self._show_result(success, msg))

        threading.Thread(target=process, daemon=True).start()

    def _show_result(self, success, msg):
        color = COLORS["success"] if success else COLORS["danger"]
        self._lbl_result.configure(text=msg, fg=color)
        if self._btn_ship:
            self._btn_ship.configure(text="  출고 처리  ", state="normal")

        if success:
            # 폼 초기화
            self.ent_product.delete(0, tk.END)
            self.ent_customer.delete(0, tk.END)
            self.ent_qty.delete(0, tk.END)
            self.ent_price.delete(0, tk.END)
            self.ent_note.delete(0, tk.END)
            self._selected_product  = None
            self._selected_customer = None
            self._lbl_pname.configure(text="")
            self._lbl_cname.configure(text="")
            self._lbl_stock_warn.configure(text="")
            self._stock_frame.pack_forget()
            self._summary_frame.pack_forget()
            # 이력 새로고침
            self._load_caches()
