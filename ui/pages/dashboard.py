"""
재고관리 시스템 - 대시보드 페이지
"""

import tkinter as tk
from tkinter import ttk
import threading
from core.constants import COLORS, FONT_FAMILY, FONT_SIZES


class DashboardPage:
    def __init__(self, app):
        self.app = app

    def render(self):
        """대시보드 페이지 렌더링"""
        canvas = tk.Canvas(self.app.content_frame, bg=COLORS["bg"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.app.content_frame, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=COLORS["bg"])

        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

        # 타이틀
        tk.Label(scroll_frame, text="📊 대시보드", bg=COLORS["bg"], fg=COLORS["text"],
                 font=(FONT_FAMILY, FONT_SIZES["title"], "bold")).pack(anchor="w", padx=5, pady=(0, 15))

        # 로딩 표시
        loading = tk.Label(scroll_frame, text="데이터 로딩 중...", bg=COLORS["bg"],
                           fg=COLORS["text_secondary"], font=(FONT_FAMILY, 12))
        loading.pack(pady=20)

        def load_data():
            try:
                parts = self.app.db.get_all_parts()
                products = self.app.db.get_all_products()
                alerts = self.app.db.get_safety_stock_alerts()
                history = self.app.db.get_all_history()
                recent = history[-10:] if len(history) > 10 else history
                recent.reverse()
                self.app.root.after(0, lambda: self._render_dashboard(
                    scroll_frame, loading, parts, products, alerts, recent))
            except Exception as e:
                err_msg = str(e)
                self.app.root.after(0, lambda: loading.configure(text=f"데이터 로드 실패: {err_msg}"))

        threading.Thread(target=load_data, daemon=True).start()

    def _render_dashboard(self, parent, loading_label, parts, products, alerts, recent_history):
        loading_label.destroy()

        # 요약 카드
        summary_frame = tk.Frame(parent, bg=COLORS["bg"])
        summary_frame.pack(fill=tk.X, padx=5, pady=(0, 15))

        cards_data = [
            ("부품 종류", f"{len(parts)}건", COLORS["primary"], "🔩"),
            ("제품 종류", f"{len(products)}건", COLORS["success"], "📦"),
            ("안전재고 경고", f"{len(alerts)}건",
             COLORS["danger"] if alerts else COLORS["success"], "⚠️"),
            ("총 부품 재고", f"{sum(int(p.get('현재재고', 0)) for p in parts):,}개",
             COLORS["warning"], "📊"),
        ]

        for i, (title, value, color, icon) in enumerate(cards_data):
            card = tk.Frame(summary_frame, bg=COLORS["card_bg"], padx=20, pady=15,
                            highlightbackground=COLORS["border"], highlightthickness=1)
            card.pack(side=tk.LEFT, fill=tk.X, expand=True,
                      padx=(0 if i == 0 else 5, 5 if i < 3 else 0))

            tk.Label(card, text=f"{icon} {title}", bg=COLORS["card_bg"],
                     fg=COLORS["text_secondary"], font=(FONT_FAMILY, FONT_SIZES["small"])).pack(anchor="w")
            tk.Label(card, text=value, bg=COLORS["card_bg"],
                     fg=color, font=(FONT_FAMILY, FONT_SIZES["stat"], "bold")).pack(anchor="w", pady=(5, 0))

        # 안전재고 경고 섹션
        if alerts:
            alert_card = tk.Frame(parent, bg=COLORS["alert_bg"], padx=15, pady=10,
                                  highlightbackground=COLORS["danger"], highlightthickness=1)
            alert_card.pack(fill=tk.X, padx=5, pady=(0, 15))

            tk.Label(alert_card, text="⚠️ 안전재고 경고", bg=COLORS["alert_bg"],
                     fg=COLORS["danger"], font=(FONT_FAMILY, 12, "bold")).pack(anchor="w")

            for a in alerts:
                tk.Label(alert_card,
                         text=f"  • {a['부품명']}({a['품번']}): 현재 {a['현재재고']}개 / 안전재고 {a['안전재고']}개",
                         bg=COLORS["alert_bg"], fg=COLORS["danger"],
                         font=(FONT_FAMILY, FONT_SIZES["small"])).pack(anchor="w", pady=1)

        # 최근 이력
        hist_card = tk.Frame(parent, bg=COLORS["card_bg"], padx=15, pady=10,
                             highlightbackground=COLORS["border"], highlightthickness=1)
        hist_card.pack(fill=tk.X, padx=5, pady=(0, 15))

        tk.Label(hist_card, text="📜 최근 입출고 이력", bg=COLORS["card_bg"],
                 fg=COLORS["text"], font=(FONT_FAMILY, 12, "bold")).pack(anchor="w", pady=(0, 10))

        if recent_history:
            cols = ("일시", "구분", "유형", "품번", "품명", "수량", "잔여재고")
            tree = ttk.Treeview(hist_card, columns=cols, show="headings",
                                height=min(10, len(recent_history)))
            for col in cols:
                tree.heading(col, text=col)
                tree.column(col, width=100, anchor="center")
            tree.column("일시", width=150)
            tree.column("품명", width=200)

            # 반응형: 스크롤바 추가 및 품명 컬럼 자동 조정
            tree_scroll = ttk.Scrollbar(hist_card, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=tree_scroll.set)

            def _on_dash_tree_resize(event, t=tree):
                total = t.winfo_width()
                fixed = 150 + 100 + 100 + 100 + 100 + 100 + 20  # 일시+구분+유형+품번+수량+잔여+스크롤바
                remaining = max(120, total - fixed)
                t.column("품명", width=remaining)
            tree.bind("<Configure>", _on_dash_tree_resize)

            for h in recent_history:
                tree.insert("", "end", values=(
                    h.get("일시", ""), h.get("구분", ""), h.get("유형", ""),
                    h.get("품번/제품코드", ""), h.get("품명", ""),
                    h.get("수량", ""), h.get("잔여재고", ""),
                ))
            tree_frame_d = tk.Frame(hist_card, bg=COLORS["card_bg"])
            tree_frame_d.pack(fill=tk.BOTH, expand=True)
            tree.master = tree_frame_d  # 부모 변경
            tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, in_=tree_frame_d)
            tree_scroll.pack(side=tk.RIGHT, fill=tk.Y, in_=tree_frame_d)
        else:
            tk.Label(hist_card, text="아직 입출고 이력이 없습니다.",
                     bg=COLORS["card_bg"], fg=COLORS["text_secondary"],
                     font=(FONT_FAMILY, FONT_SIZES["small"])).pack(pady=10)
