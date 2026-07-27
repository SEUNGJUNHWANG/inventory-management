"""
재고관리 시스템 - 제품 관리 페이지
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
from core.constants import COLORS, FONT_FAMILY, FONT_SIZES, PRODUCTS_COLUMNS
from ui.widget_utils import flash_btn
from core.auth import Session


def _bind_tree_scroll(tree):
    """Treeview에 마우스 휠 스크롤 바인딩 (hover 기반)"""
    def _on_mw(e):
        try:
            tree.yview_scroll(int(-1 * (e.delta / 120)), "units")
        except Exception:
            pass
    tree.bind("<Enter>", lambda e: tree.bind_all("<MouseWheel>", _on_mw))
    tree.bind("<Leave>", lambda e: tree.unbind_all("<MouseWheel>"))


def _fmt_price(v):
    """숫자를 천단위 콤마 포함 문자열로 변환. 0이면 '-' 표시."""
    try:
        f = float(v)
        if f == 0:
            return "-"
        return f"{f:,.0f}"
    except Exception:
        return str(v) if v else "-"


def _fmt_margin(cost, selling_price):
    """마진율 계산 및 포맷: (판매가 - 원가) / 판매가 × 100"""
    try:
        c = float(cost)
        s = float(selling_price)
        if s <= 0:
            return "-"
        margin = (s - c) / s * 100
        return f"{margin:.1f}%"
    except Exception:
        return "-"


class ProductsPage:
    def __init__(self, app):
        self.app = app
        self.products_tree = None
        self.products_menu = None
        self._action_overlay = None
        self._hovered_item = None

    def render(self):
        scroll_frame = self.app._create_scrollable_frame()

        # ── 헤더 ──
        header = tk.Frame(scroll_frame, bg=COLORS["bg"])
        header.pack(fill=tk.X, padx=5, pady=(0, 10))
        tk.Label(header, text="📦 제품 관리", bg=COLORS["bg"],
                 fg=COLORS["text"], font=(FONT_FAMILY, FONT_SIZES["title"], "bold")).pack(side=tk.LEFT)
        if Session.has_write("products"):
            tk.Button(header, text="+ 제품 추가", font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                      bg=COLORS["primary"], fg="white", padx=15, pady=5,
                      cursor="hand2", command=self._add_product_dialog).pack(side=tk.RIGHT)

        # ── 테이블 카드 ──
        card = tk.Frame(scroll_frame, bg=COLORS["card_bg"],
                        highlightbackground=COLORS["border"], highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True, padx=5)

        self.products_tree = ttk.Treeview(card, columns=PRODUCTS_COLUMNS, show="headings", height=20)
        for col in PRODUCTS_COLUMNS:
            self.products_tree.heading(col, text=col)
            self.products_tree.column(col, width=120, anchor="center")

        # 컬럼별 너비 조정
        self.products_tree.column("제품코드", width=110)
        self.products_tree.column("제품명",   width=220)
        self.products_tree.column("규격",     width=140)
        self.products_tree.column("현재재고", width=80)
        self.products_tree.column("원가",     width=95)
        self.products_tree.column("판매가",   width=95)
        self.products_tree.column("마진율",   width=75)
        # 비고는 반응형으로 자동 조정

        prod_scroll = ttk.Scrollbar(card, orient="vertical", command=self.products_tree.yview)
        self.products_tree.configure(yscrollcommand=prod_scroll.set)
        self.products_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        prod_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 반응형: 비고 컬럼 너비 자동 조정
        def _on_tree_resize(event):
            total = self.products_tree.winfo_width()
            fixed = 110 + 220 + 140 + 80 + 95 + 95 + 75 + 20  # 스크롤바 포함
            remaining = max(80, total - fixed)
            self.products_tree.column("비고", width=remaining)
        self.products_tree.bind("<Configure>", _on_tree_resize)
        _bind_tree_scroll(self.products_tree)

        # 더블클릭 → 수정
        self.products_tree.bind("<Double-1>", lambda e: self._edit_product_dialog())

        # 우클릭 메뉴
        self.products_menu = tk.Menu(self.app.root, tearoff=0)
        if Session.has_write("products"):
            self.products_menu.add_command(label="✏️ 수정", command=self._edit_product_dialog)
            self.products_menu.add_separator()
            self.products_menu.add_command(label="🗑️ 삭제", command=self._delete_product)
        self.products_tree.bind("<Button-3>", self._right_click)

        # ── 인라인 액션 오버레이 ──
        self._action_overlay = tk.Frame(
            card, bg="#f1f5f9", relief="solid", bd=1, cursor="arrow"
        )
        if Session.has_write("products"):
            btn_edit = tk.Label(
                self._action_overlay, text="✏️", bg="#f1f5f9",
                font=(FONT_FAMILY, 12), cursor="hand2", padx=6,
            )
            btn_edit.pack(side=tk.LEFT)
            btn_del = tk.Label(
                self._action_overlay, text="🗑️", bg="#f1f5f9",
                font=(FONT_FAMILY, 12), cursor="hand2", padx=6,
            )
            btn_del.pack(side=tk.LEFT)

            btn_edit.bind("<ButtonRelease-1>", lambda e: self._edit_product_dialog())
            btn_del.bind("<ButtonRelease-1>",  lambda e: self._delete_product())
            self._action_overlay.bind("<Leave>", lambda e: self._action_overlay.place_forget())
            btn_edit.bind("<Enter>", lambda e: btn_edit.configure(bg="#dbeafe"))
            btn_edit.bind("<Leave>", lambda e: btn_edit.configure(bg="#f1f5f9"))
            btn_del.bind("<Enter>",  lambda e: btn_del.configure(bg="#fee2e2"))
            btn_del.bind("<Leave>",  lambda e: btn_del.configure(bg="#f1f5f9"))

        self.products_tree.bind("<Motion>", self._on_tree_hover)
        self.products_tree.bind("<Leave>",  lambda e: self._action_overlay.place_forget())

        self._load_data()

    def _load_data(self):
        def load():
            try:
                products = self.app.db.get_all_products()
                costs    = self.app.db.get_all_product_costs()
                self.app.root.after(0, lambda: render(products, costs))
            except Exception as e:
                err_msg = str(e)
                self.app.root.after(0, lambda: messagebox.showerror("오류", err_msg))

        def render(products, costs):
            self.products_tree.delete(*self.products_tree.get_children())
            for p in products:
                pid   = p.get("제품코드", "")
                cost  = costs.get(str(pid), 0.0)
                sp    = p.get("판매가", 0)
                self.products_tree.insert("", "end", values=(
                    pid,
                    p.get("제품명", ""),
                    p.get("규격", ""),
                    p.get("현재재고", 0),
                    _fmt_price(cost),
                    _fmt_price(sp),
                    _fmt_margin(cost, sp),
                    p.get("비고", ""),
                ))

        threading.Thread(target=load, daemon=True).start()

    def _on_tree_hover(self, event):
        if not self._action_overlay:
            return
        item = self.products_tree.identify_row(event.y)
        if not item:
            self._action_overlay.place_forget()
            return
        bbox = self.products_tree.bbox(item)
        if not bbox:
            self._action_overlay.place_forget()
            return
        _, row_y, _, row_h = bbox
        tree_x = self.products_tree.winfo_x()
        tree_y = self.products_tree.winfo_y()
        tree_w = self.products_tree.winfo_width()
        overlay_w = 76
        scrollbar_w = 18
        x = tree_x + tree_w - overlay_w - scrollbar_w
        y = tree_y + row_y
        self._action_overlay.place(x=x, y=y, width=overlay_w, height=max(row_h, 24))
        self._action_overlay.lift()
        self._hovered_item = item
        self.products_tree.selection_set(item)

    def _right_click(self, event):
        item = self.products_tree.identify_row(event.y)
        if item:
            self.products_tree.selection_set(item)
            self.products_menu.post(event.x_root, event.y_root)

    def _add_product_dialog(self):
        dialog = tk.Toplevel(self.app.root)
        dialog.title("제품 추가")
        self.app.center_dialog(dialog, 420, 330)
        dialog.resizable(False, False)
        dialog.transient(self.app.root)
        dialog.grab_set()

        fields = {}
        labels = [
            ("제품코드",  "",  False),
            ("제품명",    "",  False),
            ("규격",      "",  False),
            ("현재재고",  "0", False),
            ("판매가(원)", "0", False),
            ("비고",      "",  False),
        ]

        for i, (label, default, readonly) in enumerate(labels):
            tk.Label(dialog, text=label + ":", font=(FONT_FAMILY, FONT_SIZES["small"])).grid(
                row=i, column=0, padx=10, pady=5, sticky="e")
            entry = tk.Entry(dialog, font=(FONT_FAMILY, FONT_SIZES["small"]), width=30)
            entry.grid(row=i, column=1, padx=10, pady=5)
            entry.insert(0, default)
            if readonly:
                entry.configure(state="readonly")
            fields[label] = entry

        def save():
            try:
                sp_raw = fields["판매가(원)"].get().replace(",", "").strip()
                self.app.db.add_product(
                    fields["제품코드"].get().strip(),
                    fields["제품명"].get().strip(),
                    fields["규격"].get().strip(),
                    int(fields["현재재고"].get()),
                    float(sp_raw) if sp_raw else 0,
                    fields["비고"].get().strip(),
                )
                messagebox.showinfo("성공", "제품이 추가되었습니다.")
                dialog.destroy()
                self._load_data()
            except Exception as e:
                messagebox.showerror("오류", str(e))

        save_btn = tk.Button(dialog, text="저장",
                             font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                             bg=COLORS["primary"], fg="white", padx=20, pady=5,
                             command=save)
        save_btn.grid(row=len(labels), column=0, columnspan=2, pady=15)
        dialog.bind("<Return>", lambda e: flash_btn(save_btn, save))

    def _edit_product_dialog(self):
        selected = self.products_tree.selection()
        if not selected:
            return
        values = self.products_tree.item(selected[0])["values"]
        # values = (제품코드, 제품명, 규격, 현재재고, 원가(표시용), 판매가(표시용), 비고)

        dialog = tk.Toplevel(self.app.root)
        dialog.title("제품 수정")
        self.app.center_dialog(dialog, 420, 410)
        dialog.resizable(False, False)
        dialog.transient(self.app.root)
        dialog.grab_set()

        fields = {}

        # 원가는 읽기전용 안내 레이블로만 표시
        # values = (제품코드, 제품명, 규격, 현재재고, 원가, 판매가, 마진율, 비고)
        label_defs = [
            ("제품코드",   str(values[0]), True),
            ("제품명",     str(values[1]), False),
            ("규격",       str(values[2]), False),
            ("현재재고",   str(values[3]), False),
            ("판매가(원)", str(values[5]).replace(",", "").replace("-", "0"), False),
            ("비고",       str(values[7]), False),
            ("변경사유",   "", False),
        ]

        for i, (label, default, readonly) in enumerate(label_defs):
            tk.Label(dialog, text=label + ":", font=(FONT_FAMILY, FONT_SIZES["small"])).grid(
                row=i, column=0, padx=10, pady=5, sticky="e")
            entry = tk.Entry(dialog, font=(FONT_FAMILY, FONT_SIZES["small"]), width=30)
            entry.grid(row=i, column=1, padx=10, pady=5)
            entry.insert(0, default)
            if readonly:
                entry.configure(state="readonly")
            fields[label] = entry

        # 변경사유 안내 (판매가 변경 시에만 이력에 기록됨)
        tk.Label(dialog, text="※ 판매가 변경 시 변경사유가 이력에 기록됩니다.",
                 bg=dialog.cget("bg"), fg=COLORS["text_secondary"],
                 font=(FONT_FAMILY, FONT_SIZES["tiny"])).grid(
            row=len(label_defs) - 1, column=1, sticky="w", padx=(0, 10))

        # 원가 · 마진율 안내 (BOM 기반 자동계산, 읽기전용)
        cost_text = f"원가: {values[4]}원  │  마진율: {values[6]}  (BOM 기준 자동계산)"
        tk.Label(dialog, text=cost_text, bg=dialog.cget("bg"),
                 fg=COLORS["text_secondary"],
                 font=(FONT_FAMILY, FONT_SIZES["tiny"])).grid(
            row=len(label_defs), column=0, columnspan=2, pady=(0, 4))

        def save():
            try:
                sp_raw = fields["판매가(원)"].get().replace(",", "").strip()
                self.app.db.update_product(
                    fields["제품코드"].get().strip(),
                    fields["제품명"].get().strip(),
                    fields["규격"].get().strip(),
                    int(fields["현재재고"].get()),
                    float(sp_raw) if sp_raw else 0,
                    fields["비고"].get().strip(),
                    change_reason=fields["변경사유"].get().strip(),
                )
                messagebox.showinfo("성공", "제품 정보가 수정되었습니다.")
                dialog.destroy()
                self._load_data()
            except Exception as e:
                messagebox.showerror("오류", str(e))

        save_btn = tk.Button(dialog, text="저장",
                             font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                             bg=COLORS["primary"], fg="white", padx=20, pady=5,
                             command=save)
        save_btn.grid(row=len(label_defs) + 1, column=0, columnspan=2, pady=10)
        dialog.bind("<Return>", lambda e: flash_btn(save_btn, save))

    def _delete_product(self):
        selected = self.products_tree.selection()
        if not selected:
            return
        values = self.products_tree.item(selected[0])["values"]
        if messagebox.askyesno("삭제 확인", f"'{values[1]}({values[0]})' 제품을 삭제하시겠습니까?"):
            try:
                self.app.db.delete_product(str(values[0]))
                messagebox.showinfo("성공", "제품이 삭제되었습니다.")
                self._load_data()
            except Exception as e:
                messagebox.showerror("오류", str(e))
