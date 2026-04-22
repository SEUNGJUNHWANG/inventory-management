"""
재고관리 시스템 - 공통 UI 위젯
카드, 스크롤 프레임 등 재사용 가능한 UI 컴포넌트
"""

import tkinter as tk
from tkinter import ttk
from core.constants import COLORS, FONT_FAMILY, FONT_SIZES


def create_card(parent, title):
    """카드 형태의 프레임 생성"""
    card = tk.Frame(parent, bg=COLORS["card_bg"], bd=0, highlightthickness=1,
                    highlightbackground="#e2e8f0")
    card.pack(fill=tk.X, pady=(0, 10))

    # 카드 헤더
    header = tk.Frame(card, bg=COLORS["card_bg"])
    header.pack(fill=tk.X, padx=20, pady=(15, 10))
    tk.Label(header, text=title, bg=COLORS["card_bg"], fg=COLORS["text"],
             font=(FONT_FAMILY, FONT_SIZES["heading"], "bold")).pack(anchor="w")

    # 카드 바디
    body = tk.Frame(card, bg=COLORS["card_bg"])
    body.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 15))

    return card, body


def create_scrollable_frame(parent):
    """스크롤 가능한 프레임 생성"""
    canvas = tk.Canvas(parent, bg=COLORS["bg"], highlightthickness=0)
    scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
    scrollable = tk.Frame(canvas, bg=COLORS["bg"])

    scrollable.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scrollable, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # 마우스 휠 바인딩
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    return scrollable, canvas


def create_form_field(parent, label_text, row, default_value="", readonly=False, width=30):
    """폼 필드 (라벨 + 입력) 생성"""
    tk.Label(parent, text=label_text, bg=COLORS["card_bg"], fg=COLORS["text"],
             font=(FONT_FAMILY, FONT_SIZES["small"])).grid(
        row=row, column=0, sticky="e", padx=(0, 10), pady=5)

    entry = tk.Entry(parent, font=(FONT_FAMILY, FONT_SIZES["small"]), width=width)
    entry.grid(row=row, column=1, sticky="w", pady=5)
    if default_value:
        entry.insert(0, str(default_value))
    if readonly:
        entry.configure(state="readonly")

    return entry


def create_stat_card(parent, title, value, color, row=0, col=0):
    """통계 카드 위젯 생성"""
    frame = tk.Frame(parent, bg=COLORS["card_bg"], bd=0, highlightthickness=1,
                     highlightbackground="#e2e8f0")
    frame.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")

    tk.Label(frame, text=title, bg=COLORS["card_bg"], fg=COLORS["text_secondary"],
             font=(FONT_FAMILY, FONT_SIZES["small"])).pack(padx=15, pady=(10, 0))
    tk.Label(frame, text=str(value), bg=COLORS["card_bg"], fg=color,
             font=(FONT_FAMILY, 24, "bold")).pack(padx=15, pady=(0, 10))

    return frame
