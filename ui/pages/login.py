# -*- coding: utf-8 -*-
"""
재고관리 시스템 - 로그인 페이지
포털 스타일 UI: 통일된 입력 필드, 고정 카드 너비
"""

import tkinter as tk
from tkinter import messagebox
import threading
from core.constants import COLORS, FONT_FAMILY, FONT_SIZES, APP_NAME, APP_VERSION
from core.auth import Session

# ── 로그인 카드 고정 너비 ──
_CARD_INNER_W = 320   # 입력 필드 영역 픽셀 너비
_CARD_PADX    = 40    # 카드 좌우 여백
_INPUT_BG     = "#f8fafc"
_INPUT_BORDER = "#cbd5e1"
_INPUT_FOCUS  = "#3b82f6"


class LoginPage:
    def __init__(self, app):
        self.app = app
        self._entry_id   = None
        self._entry_pw   = None
        self._btn_login  = None
        self._lbl_error  = None
        self._show_pw    = False
        self._id_wrap    = None
        self._pw_wrap    = None

    # ──────────────────────────────────────────
    # 렌더링
    # ──────────────────────────────────────────
    def render(self):
        for w in self.app.content_frame.winfo_children():
            w.destroy()

        # 배경 (content_frame 전체 채움)
        bg = tk.Frame(self.app.content_frame, bg=COLORS["bg"])
        bg.place(relx=0, rely=0, relwidth=1, relheight=1)

        # 카드를 중앙에 place
        card_w = _CARD_INNER_W + _CARD_PADX * 2
        card = tk.Frame(bg, bg="white",
                        highlightbackground="#e2e8f0",
                        highlightthickness=1)
        card.place(relx=0.5, rely=0.5, anchor="center", width=card_w)

        # ── 내부 패딩 래퍼 ──
        inner = tk.Frame(card, bg="white")
        inner.pack(fill=tk.X, padx=_CARD_PADX, pady=36)

        # ── 로고 영역 ──
        tk.Label(inner, text="📦", bg="white",
                 font=(FONT_FAMILY, 40)).pack(pady=(0, 6))
        tk.Label(inner, text=APP_NAME, bg="white",
                 fg=COLORS["text"],
                 font=(FONT_FAMILY, FONT_SIZES["subtitle"], "bold")).pack()
        tk.Label(inner, text="업무 포털에 로그인하세요", bg="white",
                 fg=COLORS["text_secondary"],
                 font=(FONT_FAMILY, FONT_SIZES["small"])).pack(pady=(2, 0))

        # 구분선
        sep = tk.Frame(inner, bg=COLORS["border"], height=1)
        sep.pack(fill=tk.X, pady=24)

        # ── 아이디 ──
        tk.Label(inner, text="아이디", bg="white",
                 fg=COLORS["text"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).pack(anchor="w")

        self._id_wrap = tk.Frame(inner, bg=_INPUT_BG,
                                  highlightbackground=_INPUT_BORDER,
                                  highlightthickness=1)
        self._id_wrap.pack(fill=tk.X, pady=(5, 14))

        self._entry_id = tk.Entry(
            self._id_wrap,
            font=(FONT_FAMILY, FONT_SIZES["body"]),
            relief="flat", bd=0, bg=_INPUT_BG,
            fg=COLORS["text"],
            insertbackground=COLORS["primary"],
        )
        self._entry_id.pack(fill=tk.X, ipady=9, padx=12)

        # ── 비밀번호 ──
        tk.Label(inner, text="비밀번호", bg="white",
                 fg=COLORS["text"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).pack(anchor="w")

        self._pw_wrap = tk.Frame(inner, bg=_INPUT_BG,
                                  highlightbackground=_INPUT_BORDER,
                                  highlightthickness=1)
        self._pw_wrap.pack(fill=tk.X, pady=(5, 6))

        self._entry_pw = tk.Entry(
            self._pw_wrap,
            font=(FONT_FAMILY, FONT_SIZES["body"]),
            show="●", relief="flat", bd=0,
            bg=_INPUT_BG, fg=COLORS["text"],
            insertbackground=COLORS["primary"],
        )
        self._entry_pw.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=9, padx=(12, 0))

        # 눈 아이콘 (비밀번호 표시/숨김)
        self._btn_eye = tk.Label(
            self._pw_wrap, text="👁", bg=_INPUT_BG,
            fg="#94a3b8", font=(FONT_FAMILY, 13),
            cursor="hand2", padx=10,
        )
        self._btn_eye.pack(side=tk.RIGHT, pady=2)
        self._btn_eye.bind("<ButtonRelease-1>", lambda e: self._toggle_pw())

        # ── 포커스 테두리 색상 변경 ──
        for entry, wrap in [(self._entry_id, self._id_wrap),
                             (self._entry_pw, self._pw_wrap)]:
            entry.bind("<FocusIn>",  lambda e, w=wrap: w.configure(
                highlightbackground=_INPUT_FOCUS))
            entry.bind("<FocusOut>", lambda e, w=wrap: w.configure(
                highlightbackground=_INPUT_BORDER))

        # ── 오류 메시지 ──
        self._lbl_error = tk.Label(inner, text="", bg="white",
                                   fg=COLORS["danger"],
                                   font=(FONT_FAMILY, FONT_SIZES["small"]),
                                   wraplength=_CARD_INNER_W)
        self._lbl_error.pack(pady=(6, 0))

        # ── 로그인 버튼 ──
        self._btn_login = tk.Button(
            inner, text="로그인",
            font=(FONT_FAMILY, FONT_SIZES["body_large"], "bold"),
            bg=COLORS["primary"], fg="white",
            activebackground="#2563eb", activeforeground="white",
            relief="flat", cursor="hand2",
            pady=10, bd=0,
            command=self._do_login,
        )
        self._btn_login.pack(fill=tk.X, pady=(14, 0))

        # 버튼 호버 효과
        self._btn_login.bind("<Enter>", lambda e: self._btn_login.configure(bg="#2563eb"))
        self._btn_login.bind("<Leave>", lambda e: self._btn_login.configure(bg=COLORS["primary"]))

        # ── 푸터 ──
        tk.Label(inner, text=f"v{APP_VERSION}  ·  Tenova", bg="white",
                 fg="#94a3b8",
                 font=(FONT_FAMILY, FONT_SIZES["tiny"])).pack(pady=(20, 0))

        # ── 키 바인딩 ──
        self._entry_id.bind("<Return>", lambda e: self._entry_pw.focus())
        self._entry_pw.bind("<Return>", lambda e: self._do_login())
        self._entry_id.focus()

    # ──────────────────────────────────────────
    # 눈 아이콘 토글
    # ──────────────────────────────────────────
    def _toggle_pw(self):
        self._show_pw = not self._show_pw
        self._entry_pw.configure(show="" if self._show_pw else "●")
        self._btn_eye.configure(text="🙈" if self._show_pw else "👁")

    # ──────────────────────────────────────────
    # 로그인 처리
    # ──────────────────────────────────────────
    def _do_login(self):
        uid = self._entry_id.get().strip()
        pw  = self._entry_pw.get()

        if not uid or not pw:
            self._lbl_error.configure(text="아이디와 비밀번호를 입력해주세요.")
            return

        self._btn_login.configure(text="확인 중...", state="disabled",
                                   bg="#93c5fd")
        self._lbl_error.configure(text="")

        def auth_worker():
            try:
                user = self.app.db.authenticate_user(uid, pw)
            except Exception as e:
                err = str(e)
                self.app.root.after(0, lambda: self._on_fail(f"연결 오류: {err}"))
                return

            if user:
                self.app.root.after(0, lambda: self._on_success(user))
            else:
                self.app.root.after(0, lambda: self._on_fail(
                    "아이디 또는 비밀번호가 올바르지 않습니다."))

        threading.Thread(target=auth_worker, daemon=True).start()

    def _on_success(self, user: dict):
        Session.login(user)
        self.app.on_login_success()

    def _on_fail(self, msg: str):
        self._lbl_error.configure(text=msg)
        self._btn_login.configure(text="로그인", state="normal",
                                   bg=COLORS["primary"])
        self._entry_pw.delete(0, tk.END)
        self._entry_pw.focus()
