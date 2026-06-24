# -*- coding: utf-8 -*-
"""
재고관리 시스템 - 사용자 관리 페이지
조회 권한 / 쓰기 권한 분리 버전
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
from core.constants import (
    COLORS, FONT_FAMILY, FONT_SIZES,
    ALL_MENU_IDS, MENU_LABELS,
    WRITE_MENU_IDS, WRITE_MENU_LABELS,
)
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


_USER_COLUMNS = ("아이디", "이름", "조회 메뉴", "쓰기 권한", "활성화")


class UsersPage:
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
        if not Session.can_manage_users():
            scroll_frame = self.app._create_scrollable_frame()
            tk.Label(scroll_frame, text="접근 권한이 없습니다.",
                     bg=COLORS["bg"], fg=COLORS["danger"],
                     font=(FONT_FAMILY, FONT_SIZES["title"], "bold")).pack(pady=80)
            return

        scroll_frame = self.app._create_scrollable_frame()

        # ── 헤더 ──
        header = tk.Frame(scroll_frame, bg=COLORS["bg"])
        header.pack(fill=tk.X, padx=5, pady=(0, 10))
        tk.Label(header, text="👤 사용자 관리", bg=COLORS["bg"],
                 fg=COLORS["text"],
                 font=(FONT_FAMILY, FONT_SIZES["title"], "bold")).pack(side=tk.LEFT)
        tk.Button(header, text="+ 사용자 추가",
                  font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                  bg=COLORS["primary"], fg="white", padx=15, pady=5,
                  cursor="hand2", command=self._add_dialog).pack(side=tk.RIGHT)

        # ── 안내 ──
        info = tk.Frame(scroll_frame, bg="#eff6ff",
                        highlightbackground="#bfdbfe", highlightthickness=1)
        info.pack(fill=tk.X, padx=5, pady=(0, 8))
        tk.Label(info,
                 text="  조회 권한: 해당 메뉴 화면 접근 가능   |   "
                      "쓰기 권한: 데이터 추가·수정·삭제·처리 버튼 사용 가능\n"
                      "  관리자(사용자 관리 권한 보유자)는 모든 쓰기 권한을 자동으로 가집니다.",
                 bg="#eff6ff", fg="#1e40af",
                 font=(FONT_FAMILY, FONT_SIZES["small"]),
                 justify="left").pack(anchor="w", padx=8, pady=6)

        # ── 테이블 카드 ──
        card = tk.Frame(scroll_frame, bg=COLORS["card_bg"],
                        highlightbackground=COLORS["border"], highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True, padx=5)

        self.tree = ttk.Treeview(card, columns=_USER_COLUMNS,
                                  show="headings", height=15)
        col_w = {"아이디": 130, "이름": 170, "조회 메뉴": 120,
                 "쓰기 권한": 120, "활성화": 90}
        for col in _USER_COLUMNS:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=col_w.get(col, 120), anchor="center")

        vsb = ttk.Scrollbar(card, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        _bind_tree_scroll(self.tree)
        self.tree.bind("<Double-1>", lambda e: self._edit_dialog())

        # 우클릭 메뉴
        self._ctx_menu = tk.Menu(self.app.root, tearoff=0)
        self._ctx_menu.add_command(label="✏️ 수정", command=self._edit_dialog)
        self._ctx_menu.add_command(label="🔑 비밀번호 초기화", command=self._reset_pw_dialog)
        self._ctx_menu.add_separator()
        self._ctx_menu.add_command(label="🗑️ 삭제", command=self._delete_user)
        self.tree.bind("<Button-3>", self._right_click)

        # ── 인라인 액션 오버레이 ──
        self._action_overlay = tk.Frame(card, bg="#f1f5f9", relief="solid", bd=1)
        for text, cmd, hover in [
            ("✏️", self._edit_dialog,    "#dbeafe"),
            ("🔑", self._reset_pw_dialog, "#fef9c3"),
            ("🗑️", self._delete_user,    "#fee2e2"),
        ]:
            lbl = tk.Label(self._action_overlay, text=text, bg="#f1f5f9",
                           font=(FONT_FAMILY, 12), cursor="hand2", padx=5)
            lbl.pack(side=tk.LEFT)
            lbl.bind("<ButtonRelease-1>", lambda e, c=cmd: c())
            lbl.bind("<Enter>", lambda e, w=lbl, h=hover: w.configure(bg=h))
            lbl.bind("<Leave>", lambda e, w=lbl: w.configure(bg="#f1f5f9"))

        self._action_overlay.bind("<Leave>", lambda e: self._action_overlay.place_forget())
        self.tree.bind("<Motion>", self._on_hover)
        self.tree.bind("<Leave>",  lambda e: self._action_overlay.place_forget())

        self._load_data()

    # ──────────────────────────────────────────
    # 데이터 로드
    # ──────────────────────────────────────────
    def _load_data(self):
        def load():
            try:
                users = self.app.db.get_all_users()
                self.app.root.after(0, lambda: _render(users))
            except Exception as e:
                msg = str(e)
                self.app.root.after(0, lambda: messagebox.showerror("오류", msg))

        def _render(users):
            self.tree.delete(*self.tree.get_children())
            for u in users:
                all_perms = u.get("메뉴권한", [])
                view_cnt  = len([m for m in all_perms if not m.endswith("_w")])
                write_cnt = len([m for m in all_perms if m.endswith("_w")])
                is_admin  = "users" in all_perms
                active    = "활성" if u.get("활성화", True) else "비활성"

                write_text = "전체(관리자)" if is_admin else f"{write_cnt}개"
                self.tree.insert("", "end", values=(
                    u.get("아이디", ""),
                    u.get("이름", ""),
                    f"{view_cnt} / {len(ALL_MENU_IDS)}",
                    write_text,
                    active,
                ))

        threading.Thread(target=load, daemon=True).start()

    # ──────────────────────────────────────────
    # 호버 오버레이
    # ──────────────────────────────────────────
    def _on_hover(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            self._action_overlay.place_forget(); return
        bbox = self.tree.bbox(item)
        if not bbox:
            self._action_overlay.place_forget(); return
        _, row_y, _, row_h = bbox
        ow = 96
        x  = self.tree.winfo_x() + self.tree.winfo_width() - ow - 18
        y  = self.tree.winfo_y() + row_y
        self._action_overlay.place(x=x, y=y, width=ow, height=max(row_h, 24))
        self._action_overlay.lift()
        self._hovered_item = item
        self.tree.selection_set(item)

    def _right_click(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self._ctx_menu.post(event.x_root, event.y_root)

    # ──────────────────────────────────────────
    # 권한 체크박스 빌더 (조회 / 쓰기 구분)
    # ──────────────────────────────────────────
    def _build_perm_section(self, parent, current_perms):
        """
        조회 권한 / 쓰기 권한 체크박스 섹션 생성.
        Returns: (view_vars dict, write_vars dict)
        """
        # ── 조회 권한 ──
        view_frame = tk.LabelFrame(
            parent, text="  조회 권한 (메뉴 화면 접근)  ",
            bg=COLORS["bg"], fg=COLORS["primary"],
            font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
            padx=10, pady=6,
        )
        view_frame.pack(fill=tk.X, padx=16, pady=(4, 6))

        view_vars = {}
        for i, mid in enumerate(ALL_MENU_IDS):
            var = tk.BooleanVar(value=(mid in current_perms))
            view_vars[mid] = var
            col, row = i % 2, i // 2
            tk.Checkbutton(view_frame,
                           text=MENU_LABELS.get(mid, mid),
                           variable=var, bg=COLORS["bg"],
                           font=(FONT_FAMILY, FONT_SIZES["small"])).grid(
                row=row, column=col, sticky="w", padx=6, pady=2)

        # 전체 선택/해제
        v_btn = tk.Frame(parent, bg=COLORS["bg"])
        v_btn.pack(anchor="w", padx=24, pady=(0, 4))
        tk.Button(v_btn, text="조회 전체 선택",
                  font=(FONT_FAMILY, FONT_SIZES["tiny"]),
                  command=lambda: [v.set(True) for v in view_vars.values()],
                  padx=8).pack(side=tk.LEFT, padx=4)
        tk.Button(v_btn, text="조회 전체 해제",
                  font=(FONT_FAMILY, FONT_SIZES["tiny"]),
                  command=lambda: [v.set(False) for v in view_vars.values()],
                  padx=8).pack(side=tk.LEFT)

        # ── 쓰기 권한 ──
        write_frame = tk.LabelFrame(
            parent, text="  쓰기 권한 (추가·수정·삭제·처리 버튼)  ",
            bg="#fefce8", fg="#92400e",
            font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
            padx=10, pady=6,
        )
        write_frame.pack(fill=tk.X, padx=16, pady=(2, 6))

        write_vars = {}
        for i, wid in enumerate(WRITE_MENU_IDS):
            var = tk.BooleanVar(value=(wid in current_perms))
            write_vars[wid] = var
            col, row = i % 2, i // 2
            tk.Checkbutton(write_frame,
                           text=WRITE_MENU_LABELS.get(wid, wid),
                           variable=var, bg="#fefce8",
                           font=(FONT_FAMILY, FONT_SIZES["small"])).grid(
                row=row, column=col, sticky="w", padx=6, pady=2)

        w_btn = tk.Frame(parent, bg=COLORS["bg"])
        w_btn.pack(anchor="w", padx=24, pady=(0, 6))
        tk.Button(w_btn, text="쓰기 전체 선택",
                  font=(FONT_FAMILY, FONT_SIZES["tiny"]),
                  command=lambda: [v.set(True) for v in write_vars.values()],
                  padx=8).pack(side=tk.LEFT, padx=4)
        tk.Button(w_btn, text="쓰기 전체 해제",
                  font=(FONT_FAMILY, FONT_SIZES["tiny"]),
                  command=lambda: [v.set(False) for v in write_vars.values()],
                  padx=8).pack(side=tk.LEFT)

        return view_vars, write_vars

    def _collect_perms(self, view_vars, write_vars):
        """체크박스 상태에서 권한 리스트 수집"""
        perms = [mid for mid in ALL_MENU_IDS if view_vars[mid].get()]
        perms += [wid for wid in WRITE_MENU_IDS if write_vars[wid].get()]
        return perms

    # ──────────────────────────────────────────
    # 사용자 추가 다이얼로그
    # ──────────────────────────────────────────
    def _add_dialog(self):
        dialog = tk.Toplevel(self.app.root)
        dialog.title("사용자 추가")
        self.app.center_dialog(dialog, 520, 680)
        dialog.resizable(False, True)
        dialog.transient(self.app.root)
        dialog.grab_set()

        # 스크롤 가능한 내부
        canvas = tk.Canvas(dialog, highlightthickness=0)
        vsb = ttk.Scrollbar(dialog, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        inner = tk.Frame(canvas, bg=COLORS["bg"])
        cw = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(cw, width=e.width))
        canvas.bind("<Enter>", lambda e: canvas.bind_all(
            "<MouseWheel>", lambda ev: canvas.yview_scroll(int(-1*(ev.delta/120)), "units")))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # 기본 정보
        info_f = tk.Frame(inner, bg=COLORS["bg"])
        info_f.pack(fill=tk.X, padx=16, pady=12)
        lkw = dict(bg=COLORS["bg"], font=(FONT_FAMILY, FONT_SIZES["small"]))
        ekw = dict(font=(FONT_FAMILY, FONT_SIZES["small"]), width=28)

        for row, (label, key) in enumerate([
            ("아이디", "id"), ("이름", "name"), ("비밀번호", "pw")
        ]):
            tk.Label(info_f, text=label+":", **lkw).grid(
                row=row, column=0, sticky="e", padx=(0,8), pady=5)
        ent_id = tk.Entry(info_f, **ekw); ent_id.grid(row=0, column=1, pady=5, sticky="w")
        ent_name = tk.Entry(info_f, **ekw); ent_name.grid(row=1, column=1, pady=5, sticky="w")
        ent_pw = tk.Entry(info_f, show="●", **ekw); ent_pw.grid(row=2, column=1, pady=5, sticky="w")
        ent_pw.insert(0, "admin1234")

        tk.Label(info_f, text="활성화:", **lkw).grid(row=3, column=0, sticky="e", padx=(0,8), pady=5)
        var_active = tk.BooleanVar(value=True)
        tk.Checkbutton(info_f, variable=var_active, bg=COLORS["bg"]).grid(
            row=3, column=1, sticky="w", pady=5)

        # 권한 섹션
        tk.Label(inner, text="권한 설정",
                 bg=COLORS["bg"], fg=COLORS["text"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).pack(
            anchor="w", padx=16, pady=(4, 2))

        view_vars, write_vars = self._build_perm_section(inner, [])

        def save():
            uid  = ent_id.get().strip()
            name = ent_name.get().strip()
            pw   = ent_pw.get()
            if not uid or not name or not pw:
                messagebox.showwarning("입력 오류", "아이디·이름·비밀번호는 필수입니다.", parent=dialog)
                return
            perms = self._collect_perms(view_vars, write_vars)
            try:
                ok = self.app.db.add_user(uid, pw, name, perms, var_active.get())
                if ok:
                    messagebox.showinfo("성공", f"'{name}' 사용자가 추가되었습니다.", parent=dialog)
                    dialog.destroy()
                    self._load_data()
                else:
                    messagebox.showerror("오류", f"아이디 '{uid}'가 이미 존재합니다.", parent=dialog)
            except Exception as e:
                messagebox.showerror("오류", str(e), parent=dialog)

        save_btn = tk.Button(inner, text="저장",
                             font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                             bg=COLORS["primary"], fg="white",
                             padx=24, pady=6, cursor="hand2", command=save)
        save_btn.pack(pady=14)
        dialog.bind("<Return>", lambda e: flash_btn(save_btn, save))

    # ──────────────────────────────────────────
    # 사용자 수정 다이얼로그
    # ──────────────────────────────────────────
    def _edit_dialog(self):
        selected = self.tree.selection()
        if not selected:
            return
        uid = str(self.tree.item(selected[0])["values"][0])

        try:
            all_users = self.app.db.get_all_users()
            udata = next((u for u in all_users if str(u.get("아이디")) == uid), None)
        except Exception as e:
            messagebox.showerror("오류", str(e)); return
        if not udata:
            messagebox.showerror("오류", "사용자를 찾을 수 없습니다."); return

        dialog = tk.Toplevel(self.app.root)
        dialog.title(f"사용자 수정 - {uid}")
        self.app.center_dialog(dialog, 520, 660)
        dialog.resizable(False, True)
        dialog.transient(self.app.root)
        dialog.grab_set()

        canvas = tk.Canvas(dialog, highlightthickness=0)
        vsb = ttk.Scrollbar(dialog, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        inner = tk.Frame(canvas, bg=COLORS["bg"])
        cw = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(cw, width=e.width))
        canvas.bind("<Enter>", lambda e: canvas.bind_all(
            "<MouseWheel>", lambda ev: canvas.yview_scroll(int(-1*(ev.delta/120)), "units")))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        info_f = tk.Frame(inner, bg=COLORS["bg"])
        info_f.pack(fill=tk.X, padx=16, pady=12)
        lkw = dict(bg=COLORS["bg"], font=(FONT_FAMILY, FONT_SIZES["small"]))
        ekw = dict(font=(FONT_FAMILY, FONT_SIZES["small"]), width=28)

        tk.Label(info_f, text="아이디:", **lkw).grid(row=0, column=0, sticky="e", padx=(0,8), pady=5)
        ent_id = tk.Entry(info_f, **ekw); ent_id.grid(row=0, column=1, pady=5, sticky="w")
        ent_id.insert(0, uid); ent_id.configure(state="readonly")

        tk.Label(info_f, text="이름:", **lkw).grid(row=1, column=0, sticky="e", padx=(0,8), pady=5)
        ent_name = tk.Entry(info_f, **ekw); ent_name.grid(row=1, column=1, pady=5, sticky="w")
        ent_name.insert(0, udata.get("이름", ""))

        tk.Label(info_f, text="활성화:", **lkw).grid(row=2, column=0, sticky="e", padx=(0,8), pady=5)
        var_active = tk.BooleanVar(value=udata.get("활성화", True))
        tk.Checkbutton(info_f, variable=var_active, bg=COLORS["bg"]).grid(
            row=2, column=1, sticky="w", pady=5)

        tk.Label(inner, text="권한 설정",
                 bg=COLORS["bg"], fg=COLORS["text"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).pack(
            anchor="w", padx=16, pady=(4, 2))

        current_perms = udata.get("메뉴권한", [])
        view_vars, write_vars = self._build_perm_section(inner, current_perms)

        def save():
            name = ent_name.get().strip()
            if not name:
                messagebox.showwarning("입력 오류", "이름을 입력하세요.", parent=dialog); return
            perms = self._collect_perms(view_vars, write_vars)
            try:
                self.app.db.update_user(uid, name, perms, var_active.get())
                messagebox.showinfo("성공", "사용자 정보가 수정되었습니다.", parent=dialog)
                dialog.destroy()
                self._load_data()
            except Exception as e:
                messagebox.showerror("오류", str(e), parent=dialog)

        save_btn2 = tk.Button(inner, text="저장",
                              font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                              bg=COLORS["primary"], fg="white",
                              padx=24, pady=6, cursor="hand2", command=save)
        save_btn2.pack(pady=14)
        dialog.bind("<Return>", lambda e: flash_btn(save_btn2, save))

    # ──────────────────────────────────────────
    # 비밀번호 초기화
    # ──────────────────────────────────────────
    def _reset_pw_dialog(self):
        selected = self.tree.selection()
        if not selected:
            return
        uid = str(self.tree.item(selected[0])["values"][0])

        dialog = tk.Toplevel(self.app.root)
        dialog.title(f"비밀번호 초기화 - {uid}")
        self.app.center_dialog(dialog, 380, 190)
        dialog.resizable(False, False)
        dialog.transient(self.app.root)
        dialog.grab_set()

        tk.Label(dialog, text=f"'{uid}' 계정의 새 비밀번호:",
                 font=(FONT_FAMILY, FONT_SIZES["small"])).pack(padx=24, pady=(20,6), anchor="w")
        ent_pw = tk.Entry(dialog, font=(FONT_FAMILY, FONT_SIZES["small"]),
                          width=30, show="●", relief="solid", bd=1)
        ent_pw.pack(padx=24, ipady=5)
        ent_pw.insert(0, "admin1234")

        def reset():
            pw = ent_pw.get()
            if not pw:
                messagebox.showwarning("입력 오류", "새 비밀번호를 입력하세요.", parent=dialog); return
            try:
                self.app.db.reset_password(uid, pw)
                messagebox.showinfo("성공", f"'{uid}' 비밀번호가 초기화되었습니다.", parent=dialog)
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("오류", str(e), parent=dialog)

        tk.Button(dialog, text="초기화",
                  font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                  bg=COLORS["warning"], fg="white",
                  padx=20, pady=5, cursor="hand2", command=reset).pack(pady=14)

    # ──────────────────────────────────────────
    # 사용자 삭제
    # ──────────────────────────────────────────
    def _delete_user(self):
        selected = self.tree.selection()
        if not selected:
            return
        vals = self.tree.item(selected[0])["values"]
        uid, name = str(vals[0]), str(vals[1])
        if not messagebox.askyesno("삭제 확인",
                                   f"'{name}({uid})' 계정을 삭제하시겠습니까?\n되돌릴 수 없습니다."):
            return
        try:
            ok = self.app.db.delete_user(uid)
            if ok:
                messagebox.showinfo("성공", "계정이 삭제되었습니다.")
                self._load_data()
            else:
                messagebox.showwarning("오류", "해당 계정을 찾을 수 없습니다.")
        except Exception as e:
            messagebox.showerror("오류", str(e))
