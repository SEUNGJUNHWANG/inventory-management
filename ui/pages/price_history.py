"""
재고관리 시스템 - 단가변경이력 페이지
- 월별 대시보드 (변경 건수 카드 4개 + 최근 6개월 바 차트)
- 상세 이력 테이블 (연·월·업체·인상/인하 필터 + 엑셀 내보내기)
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
from datetime import datetime, date
from collections import defaultdict
from core.constants import COLORS, FONT_FAMILY, FONT_SIZES


class PriceHistoryPage:
    """단가변경이력 페이지"""

    def __init__(self, app):
        self.app = app
        self._all_records = []      # 전체 이력 캐시
        self._suppliers   = []      # 업체 목록 캐시

    # ═══════════════════════════════════════════
    # 렌더링
    # ═══════════════════════════════════════════
    def render(self):
        self.scroll_frame = self.app._create_scrollable_frame()

        # 타이틀
        title_f = tk.Frame(self.scroll_frame, bg=COLORS["bg"])
        title_f.pack(fill=tk.X, padx=5, pady=(5, 8))
        tk.Label(title_f, text="💰 단가변경이력",
                 bg=COLORS["bg"], fg=COLORS["text"],
                 font=(FONT_FAMILY, FONT_SIZES["title"], "bold")).pack(side=tk.LEFT)

        # 새로고침 버튼
        tk.Button(title_f, text="🔄 새로고침",
                  font=(FONT_FAMILY, FONT_SIZES["small"]),
                  bg="#e2e8f0", fg=COLORS["text"], padx=10, pady=3,
                  cursor="hand2",
                  command=self._load_data).pack(side=tk.RIGHT)

        # ── 대시보드 카드 영역 (로딩 전 플레이스홀더) ──
        self.dashboard_frame = tk.Frame(self.scroll_frame, bg=COLORS["bg"])
        self.dashboard_frame.pack(fill=tk.X, padx=5, pady=(0, 6))

        # ── 바 차트 영역 ──
        self.chart_card = tk.Frame(self.scroll_frame, bg=COLORS["card_bg"],
                                    highlightbackground=COLORS["border"],
                                    highlightthickness=1)
        self.chart_card.pack(fill=tk.X, padx=5, pady=(0, 6))

        # ── 필터 + 테이블 영역 ──
        self._build_filter_bar()
        self._build_table()

        # 데이터 로드
        self._load_data()

    # ═══════════════════════════════════════════
    # 필터 바
    # ═══════════════════════════════════════════
    def _build_filter_bar(self):
        card = tk.Frame(self.scroll_frame, bg=COLORS["card_bg"],
                        highlightbackground=COLORS["border"], highlightthickness=1)
        card.pack(fill=tk.X, padx=5, pady=(0, 4))

        row = tk.Frame(card, bg=COLORS["card_bg"])
        row.pack(fill=tk.X, padx=15, pady=10)

        now = datetime.now()

        # 연도
        tk.Label(row, text="연도:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"])).pack(side=tk.LEFT)
        self.year_var = tk.StringVar(value=str(now.year))
        years = [str(y) for y in range(now.year, now.year - 5, -1)]
        year_cb = ttk.Combobox(row, textvariable=self.year_var,
                                values=years, width=7, state="readonly")
        year_cb.pack(side=tk.LEFT, padx=(4, 12))

        # 월
        tk.Label(row, text="월:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"])).pack(side=tk.LEFT)
        self.month_var = tk.StringVar(value="전체")
        months = ["전체"] + [str(m) for m in range(1, 13)]
        month_cb = ttk.Combobox(row, textvariable=self.month_var,
                                 values=months, width=6, state="readonly")
        month_cb.pack(side=tk.LEFT, padx=(4, 12))

        # 업체
        tk.Label(row, text="업체:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"])).pack(side=tk.LEFT)
        self.supplier_var = tk.StringVar(value="전체")
        self.supplier_cb = ttk.Combobox(row, textvariable=self.supplier_var,
                                         values=["전체"], width=14, state="readonly")
        self.supplier_cb.pack(side=tk.LEFT, padx=(4, 12))

        # 방향
        tk.Label(row, text="구분:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"])).pack(side=tk.LEFT)
        self.dir_var = tk.StringVar(value="전체")
        dir_cb = ttk.Combobox(row, textvariable=self.dir_var,
                               values=["전체", "인상(+)", "인하(-)", "신규"],
                               width=9, state="readonly")
        dir_cb.pack(side=tk.LEFT, padx=(4, 16))

        # 조회 버튼
        tk.Button(row, text="🔍 조회",
                  font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                  bg=COLORS["primary"], fg="white", padx=14, pady=3,
                  cursor="hand2", command=self._apply_filter).pack(side=tk.LEFT, padx=(0, 6))

        # 엑셀 내보내기 버튼
        tk.Button(row, text="📥 엑셀 내보내기",
                  font=(FONT_FAMILY, FONT_SIZES["small"]),
                  bg=COLORS["success"], fg="white", padx=12, pady=3,
                  cursor="hand2", command=self._export_excel).pack(side=tk.LEFT)

    # ═══════════════════════════════════════════
    # 테이블
    # ═══════════════════════════════════════════
    def _build_table(self):
        card = tk.Frame(self.scroll_frame, bg=COLORS["card_bg"],
                        highlightbackground=COLORS["border"], highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True, padx=5, pady=(0, 10))

        hdr = tk.Frame(card, bg=COLORS["card_bg"])
        hdr.pack(fill=tk.X, padx=15, pady=(10, 4))
        tk.Label(hdr, text="상세 변경 이력",
                 bg=COLORS["card_bg"], fg=COLORS["text"],
                 font=(FONT_FAMILY, FONT_SIZES["heading"], "bold")).pack(side=tk.LEFT)
        self.count_label = tk.Label(hdr, text="(0건)",
                                     bg=COLORS["card_bg"], fg=COLORS["text_secondary"],
                                     font=(FONT_FAMILY, FONT_SIZES["small"]))
        self.count_label.pack(side=tk.LEFT, padx=8)

        tf = tk.Frame(card, bg=COLORS["card_bg"])
        tf.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 12))

        cols = ("no", "dt", "part_id", "part_name", "supplier",
                "old_price", "new_price", "rate", "changed_by", "reason")
        self.tree = ttk.Treeview(tf, columns=cols, show="headings", height=16)

        cfg = [
            ("no",          "No",     40,  "center"),
            ("dt",          "변경일시", 140, "center"),
            ("part_id",     "품번",    100, "center"),
            ("part_name",   "부품명",  160, "w"),
            ("supplier",    "업체명",   90, "w"),
            ("old_price",   "이전단가",  90, "center"),
            ("new_price",   "변경단가",  90, "center"),
            ("rate",        "변경률",   70, "center"),
            ("changed_by",  "변경자",   70, "center"),
            ("reason",      "변경사유", 130, "w"),
        ]
        for cid, heading, width, anchor in cfg:
            self.tree.heading(cid, text=heading,
                               command=lambda c=cid: self._sort_by(c))
            self.tree.column(cid, width=width, anchor=anchor,
                              minwidth=40, stretch=(cid not in ("no",)))

        vsb = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.tag_configure("up",   background="#fef2f2", foreground="#dc2626")
        self.tree.tag_configure("down", background="#f0fdf4", foreground="#16a34a")
        self.tree.tag_configure("new",  background="#eff6ff", foreground="#2563eb")

        self._sort_col = "dt"
        self._sort_rev = True

    # ═══════════════════════════════════════════
    # 데이터 로드
    # ═══════════════════════════════════════════
    def _load_data(self):
        def load():
            try:
                records = self.app.db.get_price_history()
                monthly = self.app.db.get_price_history_monthly_summary()
                self.app.root.after(0, lambda: self._on_data_loaded(records, monthly))
            except Exception as e:
                err = str(e)
                self.app.root.after(0, lambda: messagebox.showerror("오류", err))
        threading.Thread(target=load, daemon=True).start()

    def _on_data_loaded(self, records, monthly):
        self._all_records = records

        # 업체 목록 갱신
        suppliers = sorted({r["업체명"] for r in records if r["업체명"]})
        self._suppliers = suppliers
        self.supplier_cb.configure(values=["전체"] + suppliers)

        # 대시보드 갱신
        self._render_dashboard(records, monthly)

        # 바 차트 갱신
        self._render_chart(monthly)

        # 테이블 (현재 필터 적용)
        self._apply_filter()

    # ═══════════════════════════════════════════
    # 대시보드 카드
    # ═══════════════════════════════════════════
    def _render_dashboard(self, records, monthly):
        for w in self.dashboard_frame.winfo_children():
            w.destroy()

        now_ym = datetime.now().strftime("%Y-%m")
        this_month = [r for r in records if r["변경일시"][:7] == now_ym]
        total_m    = len(this_month)
        up_m       = sum(1 for r in this_month if r["변경률"].startswith("+"))
        down_m     = sum(1 for r in this_month if r["변경률"].startswith("-"))

        # 평균 변동률 (이번 달)
        rates = []
        for r in this_month:
            rate_str = r["변경률"].replace("%", "").replace("+", "")
            try:
                rates.append(float(rate_str))
            except ValueError:
                pass
        avg_rate = sum(rates) / len(rates) if rates else 0
        avg_str  = f"{avg_rate:+.1f}%" if rates else "-"

        stats = [
            ("이번달 변경",  f"{total_m}건",  COLORS["primary"]),
            ("평균 변동률",  avg_str,          COLORS["warning"] if avg_rate >= 0 else COLORS["success"]),
            ("인상 건수",    f"{up_m}건 🔴",   COLORS["danger"]),
            ("인하 건수",    f"{down_m}건 🟢", COLORS["success"]),
        ]

        for label, value, color in stats:
            c = tk.Frame(self.dashboard_frame, bg="white",
                         highlightbackground=color, highlightthickness=2)
            c.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=3)
            tk.Label(c, text=label, bg="white", fg=COLORS["text_secondary"],
                     font=(FONT_FAMILY, FONT_SIZES["small"])).pack(pady=(10, 2))
            tk.Label(c, text=value, bg="white", fg=color,
                     font=(FONT_FAMILY, FONT_SIZES["stat"], "bold")).pack(pady=(0, 10))

    # ═══════════════════════════════════════════
    # 바 차트 (최근 6개월)
    # ═══════════════════════════════════════════
    def _render_chart(self, monthly):
        for w in self.chart_card.winfo_children():
            w.destroy()

        hdr = tk.Frame(self.chart_card, bg=COLORS["card_bg"])
        hdr.pack(fill=tk.X, padx=15, pady=(10, 4))
        tk.Label(hdr, text="월별 단가 변경 현황 (최근 6개월)",
                 bg=COLORS["card_bg"], fg=COLORS["text"],
                 font=(FONT_FAMILY, FONT_SIZES["heading"], "bold")).pack(side=tk.LEFT)

        # 최근 6개월 목록 생성
        now = datetime.now()
        months_6 = []
        for i in range(5, -1, -1):
            m = now.month - i
            y = now.year
            while m <= 0:
                m += 12
                y -= 1
            months_6.append(f"{y}-{str(m).zfill(2)}")

        chart_frame = tk.Frame(self.chart_card, bg=COLORS["card_bg"])
        chart_frame.pack(fill=tk.X, padx=15, pady=(0, 14))

        max_val = max((monthly.get(ym, {}).get("total", 0) for ym in months_6), default=1)
        if max_val == 0:
            max_val = 1

        BAR_W      = 60
        BAR_MAX_H  = 100
        CANVAS_H   = BAR_MAX_H + 50

        for ym in months_6:
            data    = monthly.get(ym, {"total": 0, "up": 0, "down": 0})
            total   = data["total"]
            up_cnt  = data["up"]
            dn_cnt  = data["down"]

            col_f = tk.Frame(chart_frame, bg=COLORS["card_bg"])
            col_f.pack(side=tk.LEFT, expand=True)

            canvas = tk.Canvas(col_f, width=BAR_W, height=CANVAS_H,
                                bg=COLORS["card_bg"], highlightthickness=0)
            canvas.pack()

            bar_h = int(BAR_MAX_H * total / max_val) if total > 0 else 2
            top_y = CANVAS_H - 30 - bar_h
            bot_y = CANVAS_H - 30

            # 인상(빨강) / 인하(초록) / 신규(파랑) 분할
            if total > 0:
                up_h  = int(bar_h * up_cnt  / total)
                dn_h  = int(bar_h * dn_cnt  / total)
                etc_h = bar_h - up_h - dn_h
                x0, x1 = BAR_W // 2 - 14, BAR_W // 2 + 14
                y_cur = bot_y
                for h, color in [(dn_h, "#22c55e"), (etc_h, "#6366f1"), (up_h, "#ef4444")]:
                    if h > 0:
                        canvas.create_rectangle(x0, y_cur - h, x1, y_cur,
                                                fill=color, outline="")
                        y_cur -= h
            else:
                canvas.create_rectangle(
                    BAR_W // 2 - 14, bot_y - 2, BAR_W // 2 + 14, bot_y,
                    fill=COLORS["border"], outline="")

            # 건수 라벨
            if total > 0:
                canvas.create_text(BAR_W // 2, top_y - 6,
                                   text=str(total), font=(FONT_FAMILY, 9, "bold"),
                                   fill=COLORS["text"])

            # 월 라벨
            label_m = ym[5:] + "월"
            canvas.create_text(BAR_W // 2, CANVAS_H - 14,
                                text=label_m, font=(FONT_FAMILY, 9),
                                fill=COLORS["text_secondary"])

        # 범례
        legend = tk.Frame(self.chart_card, bg=COLORS["card_bg"])
        legend.pack(pady=(0, 10))
        for color, text in [("#ef4444", "인상"), ("#22c55e", "인하"), ("#6366f1", "신규")]:
            tk.Frame(legend, bg=color, width=12, height=12).pack(side=tk.LEFT, padx=(8, 2))
            tk.Label(legend, text=text, bg=COLORS["card_bg"],
                     fg=COLORS["text_secondary"],
                     font=(FONT_FAMILY, FONT_SIZES["small"])).pack(side=tk.LEFT, padx=(0, 6))

    # ═══════════════════════════════════════════
    # 필터 적용 & 테이블 렌더링
    # ═══════════════════════════════════════════
    def _apply_filter(self):
        year     = self.year_var.get()
        month    = self.month_var.get()
        supplier = self.supplier_var.get()
        dir_sel  = self.dir_var.get()

        filtered = []
        for r in self._all_records:
            dt = r["변경일시"]
            if not dt.startswith(year):
                continue
            if month != "전체":
                ym = f"{year}-{month.zfill(2)}"
                if not dt.startswith(ym):
                    continue
            if supplier != "전체" and r["업체명"] != supplier:
                continue
            rate = r["변경률"]
            if dir_sel == "인상(+)" and not rate.startswith("+"):
                continue
            if dir_sel == "인하(-)" and not rate.startswith("-"):
                continue
            if dir_sel == "신규" and rate != "신규":
                continue
            filtered.append(r)

        # 정렬
        filtered.sort(key=lambda x: x.get("변경일시", ""), reverse=True)

        self._render_table(filtered)

    def _render_table(self, rows):
        self.tree.delete(*self.tree.get_children())
        self.count_label.configure(text=f"({len(rows)}건)")

        for i, r in enumerate(rows, 1):
            rate = r["변경률"]
            if rate.startswith("+"):
                tag = "up"
            elif rate.startswith("-"):
                tag = "down"
            else:
                tag = "new"

            old_p = f"{r['이전단가']:,.0f}" if r['이전단가'] else "-"
            new_p = f"{r['변경단가']:,.0f}" if r['변경단가'] else "-"

            self.tree.insert("", tk.END, values=(
                i,
                r["변경일시"],
                r["품번"],
                r["부품명"],
                r["업체명"],
                old_p,
                new_p,
                rate,
                r["변경자"],
                r["변경사유"],
            ), tags=(tag,))

    # ═══════════════════════════════════════════
    # 정렬
    # ═══════════════════════════════════════════
    def _sort_by(self, col):
        if self._sort_col == col:
            self._sort_rev = not self._sort_rev
        else:
            self._sort_col = col
            self._sort_rev = True
        self._apply_filter()

    # ═══════════════════════════════════════════
    # 엑셀 내보내기
    # ═══════════════════════════════════════════
    def _export_excel(self):
        rows = []
        for iid in self.tree.get_children():
            rows.append(self.tree.item(iid)["values"])

        if not rows:
            messagebox.showinfo("알림", "내보낼 데이터가 없습니다.")
            return

        now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = filedialog.asksaveasfilename(
            title="단가변경이력 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
            initialfile=f"단가변경이력_{now_str}.xlsx",
        )
        if not path:
            return

        try:
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "단가변경이력"

            hdr_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
            hdr_font = Font(name="맑은 고딕", bold=True, color="FFFFFF", size=10)
            nrm_font = Font(name="맑은 고딕", size=10)
            up_fill  = PatternFill(start_color="FEF2F2", end_color="FEF2F2", fill_type="solid")
            dn_fill  = PatternFill(start_color="F0FDF4", end_color="F0FDF4", fill_type="solid")
            nw_fill  = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
            thin     = Border(left=Side(style="thin"), right=Side(style="thin"),
                              top=Side(style="thin"),  bottom=Side(style="thin"))

            headers = ["No", "변경일시", "품번", "부품명", "업체명",
                       "이전단가", "변경단가", "변경률", "변경자", "변경사유"]
            for c, h in enumerate(headers, 1):
                cell = ws.cell(row=1, column=c, value=h)
                cell.font   = hdr_font
                cell.fill   = hdr_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin

            for r_idx, row in enumerate(rows, 2):
                rate_val = str(row[7])
                fill = up_fill if rate_val.startswith("+") else (
                       dn_fill if rate_val.startswith("-") else nw_fill)
                for c_idx, val in enumerate(row, 1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=val)
                    cell.font   = nrm_font
                    cell.fill   = fill
                    cell.border = thin
                    cell.alignment = Alignment(horizontal="center")

            widths = [6, 18, 14, 22, 14, 12, 12, 10, 10, 20]
            for i, w in enumerate(widths, 1):
                ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

            wb.save(path)
            messagebox.showinfo("저장 완료", f"저장되었습니다.\n{path}")
            import os
            os.startfile(path)
        except Exception as e:
            messagebox.showerror("오류", str(e))
