# -*- coding: utf-8 -*-
"""
재고관리 시스템 - 입고 / 출고 / 생산 페이지
자동완성 + 실시간 재고 카드 개선 버전
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
from core.constants import COLORS, FONT_FAMILY, FONT_SIZES
from core.auth import Session
from ui.widget_utils import flash_btn


# ═════════════════════════════════════════
# 자동완성 Entry 위젯
# ═════════════════════════════════════════
class AutocompleteEntry(tk.Entry):
    """타이핑 시 드롭다운 자동완성 Entry 위젯"""

    def __init__(self, master, on_select=None, **kwargs):
        super().__init__(master, **kwargs)
        self._items = []
        self._on_select = on_select
        self._dropdown = None
        self._lb = None
        self._matches = []
        self._global_click_id = None   # 루트 창 글로벌 클릭 바인딩 ID

        self.bind("<KeyRelease>", self._on_key_release)
        self.bind("<Button-1>",   self._on_click_show)   # 클릭 시 목록 표시
        self.bind("<FocusOut>",   self._schedule_hide)
        self.bind("<Down>",       self._focus_list)
        self.bind("<Up>",         self._focus_list)
        self.bind("<Escape>",     lambda e: self._hide())

    def set_items(self, items):
        self._items = items

    # ── 클릭 시 목록 표시 (최대 10개) ──
    def _on_click_show(self, event=None):
        if self._dropdown:
            return  # 이미 열려 있으면 무시
        if not self._items:
            return
        val = self.get().strip().lower()
        if val:
            matches = [i for i in self._items
                       if val in i[0].lower() or val in i[1].lower()][:10]
        else:
            matches = self._items[:10]
        if matches:
            self._show(matches)

    # ── 키 입력 처리 ──
    def _on_key_release(self, event):
        if event.keysym in ("Return", "Tab", "Escape", "Up", "Down"):
            return

        val = self.get().strip().lower()
        if not val:
            # 내용이 지워지면 전체 목록 최대 10개 표시
            if self._items:
                self._show(self._items[:10])
            else:
                self._hide()
            if self._on_select:
                self._on_select("", None)
            return

        matches = [
            item for item in self._items
            if val in item[0].lower() or val in item[1].lower()
        ][:10]

        if matches:
            self._show(matches)
        else:
            self._hide()

        exact = next((i for i in self._items if i[0].lower() == val), None)
        if exact and self._on_select:
            self._on_select(exact[0], exact[1])

    # ── 드롭다운 표시 ──
    def _show(self, matches):
        self._hide()
        self._matches = matches

        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height() + 1
        w = max(340, self.winfo_width())

        # 위젯을 먼저 숨겨서 만든 뒤 실제 높이를 측정하여 딱 맞게 조정
        self._dropdown = tk.Toplevel(self.winfo_toplevel())
        self._dropdown.wm_overrideredirect(True)
        self._dropdown.withdraw()          # 크기 측정 전 숨김

        outer = tk.Frame(
            self._dropdown,
            bg=COLORS["card_bg"],
            highlightbackground=COLORS["primary"],
            highlightthickness=1,
        )
        outer.pack(fill=tk.BOTH, expand=True)

        self._lb = tk.Listbox(
            outer,
            font=(FONT_FAMILY, FONT_SIZES["small"]),
            bg=COLORS["card_bg"],
            fg=COLORS["text"],
            selectbackground=COLORS["primary"],
            selectforeground="white",
            activestyle="none",
            borderwidth=0,
            highlightthickness=0,
            cursor="hand2",
        )
        self._lb.pack(fill=tk.BOTH, expand=True, padx=3, pady=3)

        for pid, name in matches:
            self._lb.insert(tk.END, f"  {pid}   {name}")

        # 실제 렌더링 크기 측정 후 창 크기 확정 (최대 280px)
        self._dropdown.update_idletasks()
        row_h  = self._lb.winfo_reqheight()   # 아이템 수에 맞는 실제 높이
        h = min(row_h + 8, 280)               # +8: outer frame 테두리 여백
        self._dropdown.wm_geometry(f"{w}x{h}+{x}+{y}")
        self._dropdown.deiconify()             # 크기 확정 후 표시
        self._dropdown.lift()

        # ButtonPress 로 FocusOut보다 먼저 처리 → 드롭다운이 닫히기 전에 선택
        self._lb.bind("<ButtonPress-1>", self._on_lb_click)
        self._lb.bind("<Return>", self._on_lb_select)
        self._lb.bind("<FocusOut>", self._schedule_hide)
        self._lb.bind("<Escape>", lambda e: self._hide())
        self._lb.bind("<Up>", self._on_lb_navigate)
        self._lb.bind("<Down>", self._on_lb_navigate)
        self._lb.bind("<Motion>", self._on_lb_motion)

        # ── 글로벌 클릭 감지: 드롭다운·Entry 바깥 클릭 시 닫기 ──
        # (Label·Frame 등 포커스 없는 위젯 클릭 시 FocusOut이 발생하지 않는 버그 보완)
        try:
            root = self.winfo_toplevel()
            self._global_click_id = root.bind(
                "<Button-1>", self._on_global_click, add="+"
            )
        except Exception:
            pass

    # ── 드롭다운 숨기기 ──
    def _hide(self):
        # 글로벌 클릭 바인딩 해제
        try:
            if self._global_click_id:
                self.winfo_toplevel().unbind("<Button-1>", self._global_click_id)
                self._global_click_id = None
        except Exception:
            pass
        if self._dropdown:
            try:
                self._dropdown.destroy()
            except Exception:
                pass
            self._dropdown = None
            self._lb = None

    def _on_global_click(self, event):
        """루트 창의 아무 곳이나 클릭 시 드롭다운/Entry 바깥이면 닫기"""
        if self._dropdown is None:
            return
        try:
            cx, cy = event.x_root, event.y_root
            # 드롭다운 영역
            dx, dy = self._dropdown.winfo_rootx(), self._dropdown.winfo_rooty()
            dw, dh = self._dropdown.winfo_width(), self._dropdown.winfo_height()
            # Entry 영역
            ex, ey = self.winfo_rootx(), self.winfo_rooty()
            ew, eh = self.winfo_width(), self.winfo_height()

            in_dropdown = dx <= cx <= dx + dw and dy <= cy <= dy + dh
            in_entry    = ex <= cx <= ex + ew and ey <= cy <= ey + eh

            if not in_dropdown and not in_entry:
                self._hide()
        except Exception:
            self._hide()

    def _schedule_hide(self, event=None):
        """FocusOut 시: 포커스가 Entry·드롭다운 내부에 있으면 닫지 않음"""
        def check_and_hide():
            try:
                focused = self.focus_get()
                if focused is None:
                    self._hide()
                    return
                if focused == self:
                    return
                if self._dropdown and str(focused).startswith(str(self._dropdown)):
                    return
            except Exception:
                pass
            self._hide()
        self.after(150, check_and_hide)

    def _focus_list(self, event=None):
        """방향키로 드롭다운 항목 탐색 시작"""
        if self._lb and self._lb.size() > 0:
            self._lb.focus_set()
            cur = self._lb.curselection()
            if not cur:
                self._lb.selection_set(0)
                self._lb.activate(0)
        return "break"

    def _on_lb_motion(self, event):
        """마우스 오버 시 해당 항목 하이라이트"""
        idx = self._lb.nearest(event.y)
        if idx >= 0:
            self._lb.selection_clear(0, tk.END)
            self._lb.selection_set(idx)
            self._lb.activate(idx)

    def _on_lb_navigate(self, event):
        """리스트박스 위아래 이동, 맨 위에서 Up 누르면 Entry로 복귀"""
        size = self._lb.size()
        cur = self._lb.curselection()
        idx = cur[0] if cur else -1

        if event.keysym == "Down":
            new_idx = min(idx + 1, size - 1)
        else:
            new_idx = idx - 1
            if new_idx < 0:
                self.focus_set()
                return "break"

        self._lb.selection_clear(0, tk.END)
        self._lb.selection_set(new_idx)
        self._lb.activate(new_idx)
        self._lb.see(new_idx)
        return "break"

    def _on_lb_click(self, event=None):
        """마우스 클릭으로 선택 (ButtonPress = FocusOut보다 먼저 실행됨)"""
        idx = self._lb.nearest(event.y)
        if idx < 0 or idx >= len(self._matches):
            return
        pid, name = self._matches[idx]
        self.delete(0, tk.END)
        self.insert(0, pid)
        self._hide()
        self.focus_set()
        if self._on_select:
            self._on_select(pid, name)

    def _on_lb_select(self, event=None):
        """Enter 키로 현재 선택 항목 확정"""
        if not self._lb:
            return
        sel = self._lb.curselection()
        if not sel:
            return
        pid, name = self._matches[sel[0]]
        self.delete(0, tk.END)
        self.insert(0, pid)
        self._hide()
        self.focus_set()
        if self._on_select:
            self._on_select(pid, name)


# ═════════════════════════════════════════
# 실시간 재고 정보 카드 위젯
# ═════════════════════════════════════════
class StockInfoCard(tk.Frame):
    """선택된 부품/제품의 재고 현황 카드"""

    def __init__(self, master, **kwargs):
        super().__init__(master, bg=COLORS["card_bg"], **kwargs)
        self._build()

    def _build(self):
        inner = tk.Frame(
            self,
            bg="#f8fafc",
            highlightbackground=COLORS["border"],
            highlightthickness=1,
            padx=18,
            pady=14,
        )
        inner.pack(fill=tk.X)

        top = tk.Frame(inner, bg="#f8fafc")
        top.pack(fill=tk.X, pady=(0, 10))

        self._name_lbl = tk.Label(
            top, text="", bg="#f8fafc", fg=COLORS["text"],
            font=(FONT_FAMILY, FONT_SIZES["heading"], "bold"), anchor="w",
        )
        self._name_lbl.pack(side=tk.LEFT)

        self._badge = tk.Label(
            top, text="", bg=COLORS["success"], fg="white",
            font=(FONT_FAMILY, FONT_SIZES["tiny"], "bold"),
            padx=10, pady=3,
        )
        self._badge.pack(side=tk.LEFT, padx=12)

        stats = tk.Frame(inner, bg="#f8fafc")
        stats.pack(fill=tk.X)

        self._stat_current = self._make_stat(stats, "현재 재고", COLORS["primary"])
        self._stat_current.pack(side=tk.LEFT, padx=(0, 40))

        self._stat_safety = self._make_stat(stats, "안전 재고", COLORS["text_secondary"])
        self._stat_safety.pack(side=tk.LEFT, padx=(0, 40))

        self._stat_extra = self._make_stat(stats, "단  가", COLORS["text_secondary"])
        self._stat_extra.pack(side=tk.LEFT, padx=(0, 40))

        self._stat_supplier = self._make_stat(stats, "업체명", COLORS["text_secondary"])
        self._stat_supplier.pack(side=tk.LEFT)

    def _make_stat(self, parent, label, color):
        f = tk.Frame(parent, bg="#f8fafc")
        tk.Label(f, text=label, bg="#f8fafc",
                 fg=COLORS["text_secondary"],
                 font=(FONT_FAMILY, FONT_SIZES["tiny"])).pack(anchor="w")
        val = tk.Label(f, text="—", bg="#f8fafc", fg=color,
                       font=(FONT_FAMILY, FONT_SIZES["stat"], "bold"))
        val.pack(anchor="w")
        f._val = val
        return f

    def update_part(self, part: dict):
        current  = int(part.get("현재재고", 0))
        safety   = int(part.get("안전재고", 0))
        price    = part.get("단가", 0)
        name     = part.get("부품명", "")
        supplier = part.get("업체명", "") or "—"

        if safety > 0 and current <= safety:
            status, badge_bg = "⚠  재고 부족", COLORS["danger"]
            stock_color = COLORS["danger"]
        elif safety > 0 and current <= safety * 1.2:
            status, badge_bg = "△  주의", COLORS["warning"]
            stock_color = COLORS["warning"]
        else:
            status, badge_bg = "●  정상", COLORS["success"]
            stock_color = COLORS["success"]

        self._name_lbl.configure(text=name)
        self._badge.configure(text=status, bg=badge_bg)
        self._stat_current._val.configure(text=f"{current:,}개", fg=stock_color)
        self._stat_safety._val.configure(text=f"{safety:,}개", fg=COLORS["text_secondary"])
        self._stat_extra._val.configure(
            text=f"{int(price):,}원" if price else "—",
            fg=COLORS["text_secondary"],
        )
        self._stat_supplier._val.configure(text=supplier, fg=COLORS["text_secondary"])
        self.pack(fill=tk.X, padx=20, pady=(0, 12))

    def update_product(self, product: dict, bom_count: int = 0, cost: float = 0):
        current = int(product.get("현재재고", 0))
        name    = product.get("제품명", "")

        self._name_lbl.configure(text=name)
        self._badge.configure(text="●  제품", bg=COLORS["info"])
        self._stat_current._val.configure(text=f"{current:,}개", fg=COLORS["primary"])
        self._stat_safety._val.configure(
            text=f"BOM  {bom_count}종", fg=COLORS["text_secondary"])
        # 원가: BOM 기반으로 계산된 제품 원가 표시
        cost_text = f"{int(cost):,}원" if cost else "—"
        self._stat_extra._val.configure(text=cost_text, fg=COLORS["text_secondary"])
        self._stat_supplier._val.configure(text="—", fg=COLORS["text_secondary"])
        self.pack(fill=tk.X, padx=20, pady=(0, 12))

    def clear(self):
        self.pack_forget()


# ═════════════════════════════════════════
# 부품 입고 페이지
# ═════════════════════════════════════════
class ReceivePage:
    def __init__(self, app):
        self.app = app
        self._selected_part = None

    def render(self):
        card = self.app._create_card("📥 부품 입고")

        tk.Label(
            card,
            text="품번을 입력하면 자동완성 목록이 나타납니다. 선택 후 수량을 입력하고 Enter 또는 버튼을 누르세요.",
            bg=COLORS["card_bg"], fg=COLORS["text_secondary"],
            font=(FONT_FAMILY, FONT_SIZES["small"]),
        ).pack(anchor="w", padx=20, pady=(0, 12))

        form = tk.Frame(card, bg=COLORS["card_bg"])
        form.pack(fill=tk.X, padx=20, pady=5)

        tk.Label(form, text="품번:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).grid(
            row=0, column=0, sticky="e", padx=(0, 8), pady=7)
        self.part_id = AutocompleteEntry(
            form, on_select=self._on_part_selected,
            font=(FONT_FAMILY, FONT_SIZES["body"]), width=26,
        )
        self.part_id.grid(row=0, column=1, pady=7, sticky="w")
        self.part_id.focus_set()
        self.part_id.bind("<Return>", lambda e: self.qty.focus_set())

        tk.Label(form, text="수량:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).grid(
            row=1, column=0, sticky="e", padx=(0, 8), pady=7)
        self.qty = tk.Entry(form, font=(FONT_FAMILY, FONT_SIZES["body"]), width=16)
        self.qty.grid(row=1, column=1, pady=7, sticky="w")
        self._btn_receive = None

        tk.Label(form, text="비고:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).grid(
            row=2, column=0, sticky="e", padx=(0, 8), pady=7)
        self.note = tk.Entry(form, font=(FONT_FAMILY, FONT_SIZES["small"]), width=42)
        self.note.grid(row=2, column=1, columnspan=2, pady=7, sticky="w")

        self._card_container = tk.Frame(card, bg=COLORS["card_bg"])
        self._card_container.pack(fill=tk.X)
        self.stock_card = StockInfoCard(self._card_container)

        btn_frame = tk.Frame(card, bg=COLORS["card_bg"])
        btn_frame.pack(fill=tk.X, padx=20, pady=(8, 15))
        if Session.has_write("receive"):
            self._btn_receive = tk.Button(
                btn_frame, text="  입고 처리  ",
                font=(FONT_FAMILY, FONT_SIZES["body_large"], "bold"),
                bg=COLORS["success"], fg="white", padx=20, pady=8,
                cursor="hand2", relief="flat", command=self._do_receive,
            )
            self._btn_receive.pack(side=tk.LEFT)
        self.qty.bind("<Return>", lambda e: flash_btn(self._btn_receive, self._do_receive) if self._btn_receive else self._do_receive())

        self.result_label = tk.Label(
            card, text="", bg=COLORS["card_bg"],
            font=(FONT_FAMILY, FONT_SIZES["body_large"]),
            wraplength=600, justify="left",
        )
        self.result_label.pack(anchor="w", padx=20, pady=(0, 15))

        self._load_parts_cache()

    def _load_parts_cache(self):
        def load():
            try:
                parts = self.app.db.get_all_parts()
                items = [(p.get("품번", ""), p.get("부품명", "")) for p in parts]
                self.app.root.after(0, lambda: self.part_id.set_items(items))
            except Exception:
                pass
        threading.Thread(target=load, daemon=True).start()

    def _on_part_selected(self, part_id, part_name):
        if not part_id:
            self.stock_card.clear()
            self._selected_part = None
            return

        def lookup():
            part = self.app.db.get_part_by_id(part_id)
            if part:
                self._selected_part = part
                self.app.root.after(0, lambda: self.stock_card.update_part(part))
            else:
                self._selected_part = None
                self.app.root.after(0, self.stock_card.clear)

        threading.Thread(target=lookup, daemon=True).start()

    def _do_receive(self):
        part_id = self.part_id.get().strip()
        qty_str = self.qty.get().strip()
        note    = self.note.get().strip()

        if not part_id or not qty_str:
            messagebox.showwarning("입력 오류", "품번과 수량을 입력해 주세요.")
            return
        try:
            qty = int(qty_str)
            if qty <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("입력 오류", "수량은 양의 정수를 입력해 주세요.")
            return

        def process():
            success, msg = self.app.db.receive_part(part_id, qty, note)
            self.app.root.after(0, lambda: self._show_result(success, msg))

        threading.Thread(target=process, daemon=True).start()

    def _show_result(self, success, msg):
        color = COLORS["success"] if success else COLORS["danger"]
        self.result_label.configure(text=msg, fg=color)
        if success:
            self.part_id.delete(0, tk.END)
            self.qty.delete(0, tk.END)
            self.note.delete(0, tk.END)
            self.stock_card.clear()
            self._selected_part = None
            self.part_id.focus_set()
            self._load_parts_cache()


# ═════════════════════════════════════════
# 부품 출고 페이지
# ═════════════════════════════════════════
class IssuePage:
    def __init__(self, app):
        self.app = app
        self._selected_part = None

    def render(self):
        card = self.app._create_card("📤 부품 출고")

        tk.Label(
            card,
            text="품번을 입력하면 자동완성 목록이 나타납니다. 선택 후 수량을 입력하고 Enter 또는 버튼을 누르세요.",
            bg=COLORS["card_bg"], fg=COLORS["text_secondary"],
            font=(FONT_FAMILY, FONT_SIZES["small"]),
        ).pack(anchor="w", padx=20, pady=(0, 12))

        form = tk.Frame(card, bg=COLORS["card_bg"])
        form.pack(fill=tk.X, padx=20, pady=5)

        tk.Label(form, text="품번:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).grid(
            row=0, column=0, sticky="e", padx=(0, 8), pady=7)
        self.part_id = AutocompleteEntry(
            form, on_select=self._on_part_selected,
            font=(FONT_FAMILY, FONT_SIZES["body"]), width=26,
        )
        self.part_id.grid(row=0, column=1, pady=7, sticky="w")
        self.part_id.focus_set()
        self.part_id.bind("<Return>", lambda e: self.qty.focus_set())

        tk.Label(form, text="수량:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).grid(
            row=1, column=0, sticky="e", padx=(0, 8), pady=7)
        self.qty = tk.Entry(form, font=(FONT_FAMILY, FONT_SIZES["body"]), width=16)
        self.qty.grid(row=1, column=1, pady=7, sticky="w")
        self._btn_issue = None

        tk.Label(form, text="비고:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).grid(
            row=2, column=0, sticky="e", padx=(0, 8), pady=7)
        self.note = tk.Entry(form, font=(FONT_FAMILY, FONT_SIZES["small"]), width=42)
        self.note.grid(row=2, column=1, columnspan=2, pady=7, sticky="w")

        self._card_container = tk.Frame(card, bg=COLORS["card_bg"])
        self._card_container.pack(fill=tk.X)
        self.stock_card = StockInfoCard(self._card_container)

        btn_frame = tk.Frame(card, bg=COLORS["card_bg"])
        btn_frame.pack(fill=tk.X, padx=20, pady=(8, 15))
        if Session.has_write("issue"):
            self._btn_issue = tk.Button(
                btn_frame, text="  출고 처리  ",
                font=(FONT_FAMILY, FONT_SIZES["body_large"], "bold"),
                bg=COLORS["warning"], fg="white", padx=20, pady=8,
                cursor="hand2", relief="flat", command=self._do_issue,
            )
            self._btn_issue.pack(side=tk.LEFT)
        self.qty.bind("<Return>", lambda e: flash_btn(self._btn_issue, self._do_issue) if self._btn_issue else self._do_issue())

        self.result_label = tk.Label(
            card, text="", bg=COLORS["card_bg"],
            font=(FONT_FAMILY, FONT_SIZES["body_large"]),
            wraplength=600, justify="left",
        )
        self.result_label.pack(anchor="w", padx=20, pady=(0, 15))

        self._load_parts_cache()

    def _load_parts_cache(self):
        def load():
            try:
                parts = self.app.db.get_all_parts()
                items = [(p.get("품번", ""), p.get("부품명", "")) for p in parts]
                self.app.root.after(0, lambda: self.part_id.set_items(items))
            except Exception:
                pass
        threading.Thread(target=load, daemon=True).start()

    def _on_part_selected(self, part_id, part_name):
        if not part_id:
            self.stock_card.clear()
            self._selected_part = None
            return

        def lookup():
            part = self.app.db.get_part_by_id(part_id)
            if part:
                self._selected_part = part
                self.app.root.after(0, lambda: self.stock_card.update_part(part))
            else:
                self._selected_part = None
                self.app.root.after(0, self.stock_card.clear)

        threading.Thread(target=lookup, daemon=True).start()

    def _do_issue(self):
        part_id = self.part_id.get().strip()
        qty_str = self.qty.get().strip()
        note    = self.note.get().strip()

        if not part_id or not qty_str:
            messagebox.showwarning("입력 오류", "품번과 수량을 입력해 주세요.")
            return
        try:
            qty = int(qty_str)
            if qty <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("입력 오류", "수량은 양의 정수를 입력해 주세요.")
            return

        def process():
            success, msg = self.app.db.issue_part(part_id, qty, note)
            self.app.root.after(0, lambda: self._show_result(success, msg))

        threading.Thread(target=process, daemon=True).start()

    def _show_result(self, success, msg):
        color = COLORS["success"] if success else COLORS["danger"]
        self.result_label.configure(text=msg, fg=color)
        if success:
            self.part_id.delete(0, tk.END)
            self.qty.delete(0, tk.END)
            self.note.delete(0, tk.END)
            self.stock_card.clear()
            self._selected_part = None
            self.part_id.focus_set()
            self._load_parts_cache()


# ═════════════════════════════════════════
# 제품 생산 페이지
# ═════════════════════════════════════════
class ProducePage:
    def __init__(self, app):
        self.app = app
        self._selected_product = None

    def render(self):
        card = self.app._create_card("🏭 제품 생산 (BOM 기반 자동 출고)")

        tk.Label(
            card,
            text="제품코드를 입력하면 자동완성 목록이 나타납니다. 선택 후 생산수량을 입력하세요.",
            bg=COLORS["card_bg"], fg=COLORS["text_secondary"],
            font=(FONT_FAMILY, FONT_SIZES["small"]),
        ).pack(anchor="w", padx=20, pady=(0, 12))

        form = tk.Frame(card, bg=COLORS["card_bg"])
        form.pack(fill=tk.X, padx=20, pady=5)

        tk.Label(form, text="제품코드:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).grid(
            row=0, column=0, sticky="e", padx=(0, 8), pady=7)
        self.product_id = AutocompleteEntry(
            form, on_select=self._on_product_selected,
            font=(FONT_FAMILY, FONT_SIZES["body"]), width=26,
        )
        self.product_id.grid(row=0, column=1, pady=7, sticky="w")
        self.product_id.focus_set()
        self.product_id.bind("<Return>", lambda e: self.qty.focus_set())

        tk.Label(form, text="생산수량:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).grid(
            row=1, column=0, sticky="e", padx=(0, 8), pady=7)
        self.qty = tk.Entry(form, font=(FONT_FAMILY, FONT_SIZES["body"]), width=16)
        self.qty.grid(row=1, column=1, pady=7, sticky="w")
        self._btn_produce = None

        tk.Label(form, text="비고:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).grid(
            row=2, column=0, sticky="e", padx=(0, 8), pady=7)
        self.note = tk.Entry(form, font=(FONT_FAMILY, FONT_SIZES["small"]), width=42)
        self.note.grid(row=2, column=1, columnspan=2, pady=7, sticky="w")

        self._card_container = tk.Frame(card, bg=COLORS["card_bg"])
        self._card_container.pack(fill=tk.X)
        self.stock_card = StockInfoCard(self._card_container)

        btn_frame = tk.Frame(card, bg=COLORS["card_bg"])
        btn_frame.pack(fill=tk.X, padx=20, pady=(8, 15))
        if Session.has_write("produce"):
            self._btn_produce = tk.Button(
                btn_frame, text="  🏭  생산 처리  ",
                font=(FONT_FAMILY, FONT_SIZES["body_large"], "bold"),
                bg=COLORS["danger"], fg="white", padx=20, pady=8,
                cursor="hand2", relief="flat", command=self._do_produce,
            )
            self._btn_produce.pack(side=tk.LEFT)
        self.qty.bind("<Return>", lambda e: flash_btn(self._btn_produce, self._do_produce) if self._btn_produce else self._do_produce())

        self.result_text = tk.Text(
            card, font=(FONT_FAMILY, FONT_SIZES["small"]),
            height=10, width=70, state="disabled", wrap="word",
        )
        self.result_text.pack(padx=20, pady=(0, 15))

        self._load_products_cache()

    def _load_products_cache(self):
        def load():
            try:
                products = self.app.db.get_all_products()
                items = [(p.get("제품코드", ""), p.get("제품명", "")) for p in products]
                self.app.root.after(0, lambda: self.product_id.set_items(items))
            except Exception:
                pass
        threading.Thread(target=load, daemon=True).start()

    def _on_product_selected(self, product_id, product_name):
        if not product_id:
            self.stock_card.clear()
            self._selected_product = None
            return

        def lookup():
            product = self.app.db.get_product_by_id(product_id)
            if product:
                self._selected_product = product
                try:
                    bom = self.app.db.get_bom_for_product(product_id)
                    bom_count = len(bom)
                except Exception:
                    bom_count = 0
                # 원가 조회 (BOM 단가 기반 계산)
                try:
                    cost, _ = self.app.db.get_product_cost(product_id)
                except Exception:
                    cost = 0
                self.app.root.after(
                    0, lambda p=product, c=bom_count, v=cost:
                        self.stock_card.update_product(p, c, v))
            else:
                self._selected_product = None
                self.app.root.after(0, self.stock_card.clear)

        threading.Thread(target=lookup, daemon=True).start()

    def _do_produce(self):
        product_id = self.product_id.get().strip()
        qty_str    = self.qty.get().strip()
        note       = self.note.get().strip()

        if not product_id or not qty_str:
            messagebox.showwarning("입력 오류", "제품코드와 생산수량을 입력해 주세요.")
            return
        try:
            qty = int(qty_str)
            if qty <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("입력 오류", "생산수량은 양의 정수를 입력해 주세요.")
            return

        def process():
            success, msg, warnings = self.app.db.produce_product(product_id, qty, note, force=False)
            if success is None:
                # 재고 부족 → 확인 다이얼로그 (메인 스레드에서 실행)
                self.app.root.after(
                    0, lambda: self._confirm_negative_stock(product_id, qty, note, msg, warnings)
                )
            else:
                self.app.root.after(0, lambda: self._show_result(success, msg, warnings))

        threading.Thread(target=process, daemon=True).start()

    def _confirm_negative_stock(self, product_id, qty, note, msg, shortage_list):
        """재고 부족 시 마이너스 처리 여부를 사용자에게 확인"""
        detail = "\n".join(f"  • {s}" for s in shortage_list)
        answer = messagebox.askyesno(
            "재고 부족 — 마이너스 처리 확인",
            f"⚠️ 일부 부품의 재고가 부족합니다:\n\n{detail}\n\n"
            "그래도 생산을 진행하시겠습니까?\n"
            "(부족한 부품은 마이너스 재고로 처리됩니다)",
            icon="warning",
        )
        if answer:
            def process_force():
                success, msg2, warnings2 = self.app.db.produce_product(
                    product_id, qty, note, force=True
                )
                self.app.root.after(0, lambda: self._show_result(success, msg2, warnings2))
            threading.Thread(target=process_force, daemon=True).start()

    def _show_result(self, success, msg, warnings=None):
        color = COLORS["success"] if success else COLORS["danger"]
        result = msg
        if warnings:
            result += "\n\n⚠️ 재고 경고:\n" + "\n".join(warnings)

        self.result_text.configure(state="normal")
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert("1.0", result)
        self.result_text.configure(state="disabled", fg=color)

        if success:
            self.product_id.delete(0, tk.END)
            self.qty.delete(0, tk.END)
            self.note.delete(0, tk.END)
            self.stock_card.clear()
            self._selected_product = None
            self.product_id.focus_set()
            self._load_products_cache()
