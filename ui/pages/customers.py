# -*- coding: utf-8 -*-
"""
재고관리 시스템 - 거래처 관리 페이지
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
from core.constants import COLORS, FONT_FAMILY, FONT_SIZES, CUSTOMERS_COLUMNS
from core.auth import Session
from ui.widget_utils import flash_btn


def _bind_tree_scroll(tree):
    def _on_mw(e):
        try:
            tree.yview_scroll(int(-1 * (e.delta / 120)), "units")
        except Exception:
            pass
    tree.bind("<Enter>", lambda e: tree.bind_all("<MouseWheel>", _on_mw))
    tree.bind("<Leave>", lambda e: tree.unbind_all("<MouseWheel>"))


# 테이블에 표시할 컬럼 (거래처코드, 거래처명, 담당자, 연락처, 이메일, 주소, 비고)
_TREE_COLS = CUSTOMERS_COLUMNS
_COL_WIDTHS = {
    "거래처코드": 110,
    "거래처명":   200,
    "담당자":     100,
    "연락처":     120,
    "이메일":     180,
    "주소":       220,
    "비고":       150,
}


class CustomersPage:
    def __init__(self, app):
        self.app = app
        self.tree = None
        self._ctx_menu = None
        self._action_overlay = None
        self._hovered_item = None

    # ──────────────────────────────────────────
    # 렌더링
    # ──────────────────────────────────────────
    def render(self):
        scroll_frame = self.app._create_scrollable_frame()

        # ── 헤더 ──
        header = tk.Frame(scroll_frame, bg=COLORS["bg"])
        header.pack(fill=tk.X, padx=5, pady=(0, 10))
        tk.Label(header, text="🏢 거래처 관리", bg=COLORS["bg"],
                 fg=COLORS["text"],
                 font=(FONT_FAMILY, FONT_SIZES["title"], "bold")).pack(side=tk.LEFT)
        if Session.has_write("customers"):
            tk.Button(header, text="+ 거래처 추가",
                      font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                      bg=COLORS["primary"], fg="white", padx=15, pady=5,
                      cursor="hand2", command=self._add_dialog).pack(side=tk.RIGHT)

        # ── 안내 ──
        tk.Label(scroll_frame,
                 text="※ 거래처 정보를 등록하면 제품 출고(판매) 시 자동완성으로 선택할 수 있습니다.",
                 bg=COLORS["bg"], fg=COLORS["text_secondary"],
                 font=(FONT_FAMILY, FONT_SIZES["small"])).pack(anchor="w", padx=5, pady=(0, 6))

        # ── 테이블 카드 ──
        card = tk.Frame(scroll_frame, bg=COLORS["card_bg"],
                        highlightbackground=COLORS["border"], highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True, padx=5)

        self.tree = ttk.Treeview(card, columns=_TREE_COLS, show="headings", height=18)
        for col in _TREE_COLS:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=_COL_WIDTHS.get(col, 120), anchor="center")

        vsb = ttk.Scrollbar(card, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        _bind_tree_scroll(self.tree)
        self.tree.bind("<Double-1>", lambda e: self._edit_dialog())

        # 우클릭 메뉴
        self._ctx_menu = tk.Menu(self.app.root, tearoff=0)
        if Session.has_write("customers"):
            self._ctx_menu.add_command(label="✏️ 수정", command=self._edit_dialog)
            self._ctx_menu.add_separator()
            self._ctx_menu.add_command(label="🗑️ 삭제", command=self._delete)
        self.tree.bind("<Button-3>", self._right_click)

        # ── 인라인 액션 오버레이 ──
        self._action_overlay = tk.Frame(card, bg="#f1f5f9", relief="solid", bd=1)
        if Session.has_write("customers"):
            btn_edit = tk.Label(self._action_overlay, text="✏️", bg="#f1f5f9",
                                font=(FONT_FAMILY, 12), cursor="hand2", padx=6)
            btn_edit.pack(side=tk.LEFT)
            btn_del = tk.Label(self._action_overlay, text="🗑️", bg="#f1f5f9",
                               font=(FONT_FAMILY, 12), cursor="hand2", padx=6)
            btn_del.pack(side=tk.LEFT)

            btn_edit.bind("<ButtonRelease-1>", lambda e: self._edit_dialog())
            btn_del.bind("<ButtonRelease-1>",  lambda e: self._delete())
            self._action_overlay.bind("<Leave>", lambda e: self._action_overlay.place_forget())
            btn_edit.bind("<Enter>", lambda e: btn_edit.configure(bg="#dbeafe"))
            btn_edit.bind("<Leave>", lambda e: btn_edit.configure(bg="#f1f5f9"))
            btn_del.bind("<Enter>",  lambda e: btn_del.configure(bg="#fee2e2"))
            btn_del.bind("<Leave>",  lambda e: btn_del.configure(bg="#f1f5f9"))

        self.tree.bind("<Motion>", self._on_hover)
        self.tree.bind("<Leave>",  lambda e: self._action_overlay.place_forget())

        self._load_data()

    # ──────────────────────────────────────────
    # 데이터 로드
    # ──────────────────────────────────────────
    def _load_data(self):
        def load():
            try:
                rows = self.app.db.get_all_customers()
                self.app.root.after(0, lambda: _render(rows))
            except Exception as e:
                msg = str(e)
                self.app.root.after(0, lambda: messagebox.showerror("오류", msg))

        def _render(rows):
            self.tree.delete(*self.tree.get_children())
            for r in rows:
                self.tree.insert("", "end", values=(
                    r.get("거래처코드", ""),
                    r.get("거래처명", ""),
                    r.get("담당자", ""),
                    r.get("연락처", ""),
                    r.get("이메일", ""),
                    r.get("주소", ""),
                    r.get("비고", ""),
                ))

        threading.Thread(target=load, daemon=True).start()

    # ──────────────────────────────────────────
    # 호버 오버레이
    # ──────────────────────────────────────────
    def _on_hover(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            self._action_overlay.place_forget()
            return
        bbox = self.tree.bbox(item)
        if not bbox:
            self._action_overlay.place_forget()
            return
        _, row_y, _, row_h = bbox
        tree_x = self.tree.winfo_x()
        tree_y = self.tree.winfo_y()
        tree_w = self.tree.winfo_width()
        overlay_w = 76
        x = tree_x + tree_w - overlay_w - 18
        y = tree_y + row_y
        self._action_overlay.place(x=x, y=y, width=overlay_w, height=max(row_h, 24))
        self._action_overlay.lift()
        self._hovered_item = item
        self.tree.selection_set(item)

    def _right_click(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self._ctx_menu.post(event.x_root, event.y_root)

    # ──────────────────────────────────────────
    # 폼 헬퍼
    # ──────────────────────────────────────────
    def _build_form(self, parent, defaults=None):
        """거래처 입력 폼 구성. returns: fields dict"""
        if defaults is None:
            defaults = {}
        fields_def = [
            ("거래처코드", "코드를 입력하세요 (예: C001)"),
            ("거래처명",   "거래처 이름"),
            ("담당자",     "담당자 이름"),
            ("연락처",     "전화번호"),
            ("이메일",     "이메일 주소"),
            ("주소",       "주소"),
            ("비고",       ""),
        ]
        fields = {}
        for i, (label, placeholder) in enumerate(fields_def):
            tk.Label(parent, text=label + ":",
                     font=(FONT_FAMILY, FONT_SIZES["small"])).grid(
                row=i, column=0, sticky="e", padx=(12, 8), pady=5)
            ent = tk.Entry(parent, font=(FONT_FAMILY, FONT_SIZES["small"]), width=32)
            ent.grid(row=i, column=1, sticky="w", pady=5, padx=(0, 12))
            val = defaults.get(label, "")
            if val:
                ent.insert(0, str(val))
            elif placeholder:
                # placeholder 색상 힌트
                ent.insert(0, placeholder)
                ent.configure(fg="#94a3b8")
                def _on_focus_in(e, w=ent, ph=placeholder):
                    if w.get() == ph:
                        w.delete(0, tk.END)
                        w.configure(fg=COLORS["text"])
                def _on_focus_out(e, w=ent, ph=placeholder):
                    if not w.get():
                        w.insert(0, ph)
                        w.configure(fg="#94a3b8")
                ent.bind("<FocusIn>",  _on_focus_in)
                ent.bind("<FocusOut>", _on_focus_out)
            fields[label] = ent
        return fields

    def _get_field(self, fields, key, placeholder=""):
        val = fields[key].get().strip()
        if val == placeholder:
            return ""
        return val

    # ──────────────────────────────────────────
    # 거래처 추가 다이얼로그
    # ──────────────────────────────────────────
    def _add_dialog(self):
        dialog = tk.Toplevel(self.app.root)
        dialog.title("거래처 추가")
        self.app.center_dialog(dialog, 460, 380)
        dialog.resizable(False, False)
        dialog.transient(self.app.root)
        dialog.grab_set()

        form = tk.Frame(dialog)
        form.pack(fill=tk.X, padx=4, pady=16)
        placeholders = {
            "거래처코드": "코드를 입력하세요 (예: C001)",
            "거래처명":   "거래처 이름",
            "담당자":     "담당자 이름",
            "연락처":     "전화번호",
            "이메일":     "이메일 주소",
            "주소":       "주소",
        }
        fields = self._build_form(form, defaults={})

        def save():
            code = self._get_field(fields, "거래처코드", placeholders.get("거래처코드",""))
            name = self._get_field(fields, "거래처명",   placeholders.get("거래처명",""))
            if not code or not name:
                messagebox.showwarning("입력 오류", "거래처코드와 거래처명은 필수입니다.", parent=dialog)
                return
            try:
                ok = self.app.db.add_customer(
                    code,
                    name,
                    self._get_field(fields, "담당자",  placeholders.get("담당자","")),
                    self._get_field(fields, "연락처",  placeholders.get("연락처","")),
                    self._get_field(fields, "이메일",  placeholders.get("이메일","")),
                    self._get_field(fields, "주소",    placeholders.get("주소","")),
                    fields["비고"].get().strip(),
                )
                if ok:
                    messagebox.showinfo("성공", f"'{name}' 거래처가 추가되었습니다.", parent=dialog)
                    dialog.destroy()
                    self._load_data()
                else:
                    messagebox.showerror("오류", f"거래처코드 '{code}'가 이미 존재합니다.", parent=dialog)
            except Exception as e:
                messagebox.showerror("오류", str(e), parent=dialog)

        save_btn = tk.Button(dialog, text="저장",
                             font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                             bg=COLORS["primary"], fg="white", padx=24, pady=6,
                             cursor="hand2", command=save)
        save_btn.pack(pady=12)
        dialog.bind("<Return>", lambda e: flash_btn(save_btn, save))

    # ──────────────────────────────────────────
    # 거래처 수정 다이얼로그
    # ──────────────────────────────────────────
    def _edit_dialog(self):
        selected = self.tree.selection()
        if not selected:
            return
        vals = self.tree.item(selected[0])["values"]
        # vals = (거래처코드, 거래처명, 담당자, 연락처, 이메일, 주소, 비고)
        keys = ["거래처코드", "거래처명", "담당자", "연락처", "이메일", "주소", "비고"]
        defaults = {k: str(v) for k, v in zip(keys, vals)}

        dialog = tk.Toplevel(self.app.root)
        dialog.title(f"거래처 수정 - {vals[0]}")
        self.app.center_dialog(dialog, 460, 380)
        dialog.resizable(False, False)
        dialog.transient(self.app.root)
        dialog.grab_set()

        form = tk.Frame(dialog)
        form.pack(fill=tk.X, padx=4, pady=16)
        fields = self._build_form(form, defaults=defaults)
        # 거래처코드 읽기전용
        fields["거래처코드"].configure(state="readonly")

        def save():
            code = str(vals[0])
            name = fields["거래처명"].get().strip()
            if not name:
                messagebox.showwarning("입력 오류", "거래처명은 필수입니다.", parent=dialog)
                return
            try:
                self.app.db.update_customer(
                    code, name,
                    fields["담당자"].get().strip(),
                    fields["연락처"].get().strip(),
                    fields["이메일"].get().strip(),
                    fields["주소"].get().strip(),
                    fields["비고"].get().strip(),
                )
                messagebox.showinfo("성공", "거래처 정보가 수정되었습니다.", parent=dialog)
                dialog.destroy()
                self._load_data()
            except Exception as e:
                messagebox.showerror("오류", str(e), parent=dialog)

        save_btn = tk.Button(dialog, text="저장",
                             font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                             bg=COLORS["primary"], fg="white", padx=24, pady=6,
                             cursor="hand2", command=save)
        save_btn.pack(pady=12)
        dialog.bind("<Return>", lambda e: flash_btn(save_btn, save))

    # ──────────────────────────────────────────
    # 거래처 삭제
    # ──────────────────────────────────────────
    def _delete(self):
        selected = self.tree.selection()
        if not selected:
            return
        vals = self.tree.item(selected[0])["values"]
        code = str(vals[0])
        name = str(vals[1])
        if not messagebox.askyesno("삭제 확인",
                                   f"'{name}({code})' 거래처를 삭제하시겠습니까?"):
            return
        try:
            ok = self.app.db.delete_customer(code)
            if ok:
                messagebox.showinfo("성공", "거래처가 삭제되었습니다.")
                self._load_data()
            else:
                messagebox.showwarning("오류", "해당 거래처를 찾을 수 없습니다.")
        except Exception as e:
            messagebox.showerror("오류", str(e))
