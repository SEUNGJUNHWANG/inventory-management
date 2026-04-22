"""
재고관리 시스템 - 입출고 이력 페이지
- 개별 이력 취소 (기존)
- 생산 일괄 취소: '생산입고' 이력 우클릭 시 제품 재고 + 소요 부품 재고를 한 번에 원복
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
from datetime import datetime, timedelta
from core.constants import COLORS, FONT_FAMILY, FONT_SIZES, HISTORY_COLUMNS


class HistoryPage:
    def __init__(self, app):
        self.app = app
        self.history_tree = None
        self.history_menu = None
        self.hist_start = None
        self.hist_end = None

    def render(self):
        scroll_frame = self.app._create_scrollable_frame()

        header = tk.Frame(scroll_frame, bg=COLORS["bg"])
        header.pack(fill=tk.X, padx=5, pady=(0, 10))
        tk.Label(header, text="📜 입출고 이력", bg=COLORS["bg"],
                 fg=COLORS["text"], font=(FONT_FAMILY, FONT_SIZES["title"], "bold")).pack(side=tk.LEFT)

        # 필터
        filter_frame = tk.Frame(scroll_frame, bg=COLORS["bg"])
        filter_frame.pack(fill=tk.X, padx=5, pady=(0, 10))

        tk.Label(filter_frame, text="기간:", bg=COLORS["bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"])).pack(side=tk.LEFT)
        self.hist_start = tk.Entry(filter_frame, font=(FONT_FAMILY, FONT_SIZES["small"]), width=12)
        self.hist_start.pack(side=tk.LEFT, padx=3)
        self.hist_start.insert(0, (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d"))
        tk.Label(filter_frame, text="~", bg=COLORS["bg"]).pack(side=tk.LEFT)
        self.hist_end = tk.Entry(filter_frame, font=(FONT_FAMILY, FONT_SIZES["small"]), width=12)
        self.hist_end.pack(side=tk.LEFT, padx=3)
        self.hist_end.insert(0, datetime.now().strftime("%Y-%m-%d"))

        tk.Button(filter_frame, text="조회", font=(FONT_FAMILY, 9),
                  command=self._load_data).pack(side=tk.LEFT, padx=5)
        tk.Button(filter_frame, text="전체 조회", font=(FONT_FAMILY, 9),
                  command=lambda: self._load_data(all_data=True)).pack(side=tk.LEFT)

        card = tk.Frame(scroll_frame, bg=COLORS["card_bg"],
                        highlightbackground=COLORS["border"], highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True, padx=5)

        self.history_tree = ttk.Treeview(card, columns=HISTORY_COLUMNS, show="headings", height=20)
        for col in HISTORY_COLUMNS:
            self.history_tree.heading(col, text=col)
            self.history_tree.column(col, width=100, anchor="center")
        self.history_tree.column("No", width=50)
        self.history_tree.column("일시", width=155)
        self.history_tree.column("품명", width=200)
        self.history_tree.column("비고", width=150)

        # 반응형: 품명/비고 컬럼 너비 자동 조정
        def _on_hist_resize(event):
            total = self.history_tree.winfo_width()
            # No(50) + 일시(155) + 유형(100) + 수량(100) + 단위(80) + 스크롤바(20)
            fixed = 50 + 155 + 100 + 100 + 80 + 20
            remaining = max(200, total - fixed)
            part_w = int(remaining * 0.55)
            note_w = remaining - part_w
            self.history_tree.column("품명", width=part_w)
            self.history_tree.column("비고", width=note_w)
        self.history_tree.bind("<Configure>", _on_hist_resize)

        # 행 색상 태그 설정
        self.history_tree.tag_configure("생산입고", background="#dbeafe", foreground="#1e40af")
        self.history_tree.tag_configure("생산출고", background="#fef3c7", foreground="#92400e")
        self.history_tree.tag_configure("취소", background="#fecaca", foreground="#991b1b")

        hist_scroll = ttk.Scrollbar(card, orient="vertical", command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=hist_scroll.set)
        self.history_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        hist_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 우클릭 메뉴 (동적으로 구성)
        self.history_menu = tk.Menu(self.app.root, tearoff=0)
        self.history_tree.bind("<Button-3>", self._right_click)

        self._load_data()

    def _load_data(self, all_data=False):
        def load():
            try:
                if all_data:
                    history = self.app.db.get_all_history()
                else:
                    start = self.hist_start.get().strip()
                    end = self.hist_end.get().strip()
                    history = self.app.db.get_history_by_date_range(start, end)
                history.reverse()
                self.app.root.after(0, lambda: render(history, all_data))
            except Exception as e:
                err_msg = str(e)
                self.app.root.after(0, lambda: messagebox.showerror("오류", err_msg))

        def render(history, is_all):
            self.history_tree.delete(*self.history_tree.get_children())
            all_hist = self.app.db.get_all_history() if not is_all else history[::-1]
            for idx, h in enumerate(history):
                row_no = len(all_hist) - idx + 1 if not is_all else len(history) - idx + 1
                h_type = h.get("유형", "")
                h_direction = h.get("구분", "")

                # 행 색상 태그 결정
                tag = "normal"
                if h_type == "생산입고":
                    tag = "생산입고"
                elif h_type == "생산출고":
                    tag = "생산출고"
                elif h_direction == "취소":
                    tag = "취소"

                self.history_tree.insert("", "end", values=(
                    row_no,
                    h.get("일시", ""), h.get("구분", ""), h_type,
                    h.get("품번/제품코드", ""), h.get("품명", ""),
                    h.get("수량", ""), h.get("잔여재고", ""),
                    h.get("관련제품", ""), h.get("비고", ""),
                ), tags=(tag,))

        threading.Thread(target=load, daemon=True).start()

    def _right_click(self, event):
        item = self.history_tree.identify_row(event.y)
        if not item:
            return
        self.history_tree.selection_set(item)

        values = self.history_tree.item(item)["values"]
        h_type = str(values[3])  # 유형 컬럼

        # 메뉴 초기화 후 동적 구성
        self.history_menu.delete(0, tk.END)

        if h_type == "생산입고":
            # 생산입고 이력: 일괄 취소 메뉴 표시
            self.history_menu.add_command(
                label="🏭 생산 전체 취소 (제품 + 부품 재고 일괄 원복)",
                command=self._cancel_production,
                font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
            )
            self.history_menu.add_separator()
            self.history_menu.add_command(
                label="이 이력만 취소 (제품 재고만 원복)",
                command=self._cancel_history,
            )
        else:
            # 기타 이력: 기존 개별 취소 메뉴
            self.history_menu.add_command(
                label="이 이력 취소 (원복)",
                command=self._cancel_history,
            )

        self.history_menu.post(event.x_root, event.y_root)

    def _cancel_production(self):
        """생산 일괄 취소 - 제품 재고 + 소요 부품 재고 한 번에 원복"""
        selected = self.history_tree.selection()
        if not selected:
            return
        values = self.history_tree.item(selected[0])["values"]
        row_no = int(values[0])
        h_type = str(values[3])

        if h_type != "생산입고":
            messagebox.showwarning("알림", "생산입고 이력만 일괄 취소할 수 있습니다.")
            return

        # 상세 확인 대화상자
        info = (
            f"일시: {values[1]}\n"
            f"제품: {values[5]} ({values[4]})\n"
            f"생산수량: {values[6]}개\n\n"
            f"이 생산 건을 전체 취소하시겠습니까?\n\n"
            f"다음 작업이 일괄 수행됩니다:\n"
            f"  1. 제품 재고 -{values[6]}개 원복\n"
            f"  2. 소요된 모든 부품 재고 원복\n"
            f"  3. 취소 이력 자동 기록"
        )

        if not messagebox.askyesno("생산 전체 취소", info, icon="warning"):
            return

        # 처리 중 표시
        self.history_tree.configure(cursor="watch")

        def process():
            try:
                success, msg, details = self.app.db.cancel_production(row_no)
                self.app.root.after(0, lambda: self._show_production_cancel_result(success, msg, details))
            except Exception as e:
                err_msg = str(e)
                self.app.root.after(0, lambda: self._show_production_cancel_result(False, f"오류 발생: {err_msg}", []))

        threading.Thread(target=process, daemon=True).start()

    def _show_production_cancel_result(self, success, msg, details):
        """생산 일괄 취소 결과 표시"""
        self.history_tree.configure(cursor="")

        if success:
            # 결과를 보여주는 상세 대화상자
            result_dialog = tk.Toplevel(self.app.root)
            result_dialog.title("생산 일괄 취소 완료")
            result_dialog.transient(self.app.root)
            result_dialog.grab_set()

            # 크기 및 위치 설정
            dw, dh = 600, 450
            sw = result_dialog.winfo_screenwidth()
            sh = result_dialog.winfo_screenheight()
            x = (sw - dw) // 2
            y = (sh - dh) // 2
            result_dialog.geometry(f"{dw}x{dh}+{x}+{y}")
            result_dialog.resizable(True, True)

            # 성공 아이콘 및 제목
            header_frame = tk.Frame(result_dialog, bg="#f0fdf4", padx=15, pady=12)
            header_frame.pack(fill=tk.X)
            tk.Label(header_frame, text="✅ 생산 일괄 취소가 완료되었습니다.",
                     bg="#f0fdf4", fg="#166534",
                     font=(FONT_FAMILY, FONT_SIZES["heading"], "bold")).pack(anchor="w")

            # 상세 결과 텍스트
            text_frame = tk.Frame(result_dialog)
            text_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

            result_text = tk.Text(text_frame, font=(FONT_FAMILY, FONT_SIZES["small"]),
                                  wrap="word", state="normal")
            result_scroll = ttk.Scrollbar(text_frame, orient="vertical", command=result_text.yview)
            result_text.configure(yscrollcommand=result_scroll.set)
            result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            result_scroll.pack(side=tk.RIGHT, fill=tk.Y)

            result_text.insert("1.0", msg)
            result_text.configure(state="disabled")

            # 닫기 버튼
            btn_frame = tk.Frame(result_dialog, padx=15, pady=10)
            btn_frame.pack(fill=tk.X)
            tk.Button(btn_frame, text="확인", font=(FONT_FAMILY, FONT_SIZES["body"], "bold"),
                      bg=COLORS["primary"], fg="white", padx=25, pady=6,
                      cursor="hand2",
                      command=result_dialog.destroy).pack(side=tk.RIGHT)

            # 이력 목록 새로고침
            self._load_data()
        else:
            messagebox.showerror("생산 취소 실패", msg)

    def _cancel_history(self):
        """기존 개별 이력 취소"""
        selected = self.history_tree.selection()
        if not selected:
            return
        values = self.history_tree.item(selected[0])["values"]
        row_no = int(values[0])

        info = f"일시: {values[1]}\n유형: {values[3]}\n품명: {values[5]}\n수량: {values[6]}"
        if not messagebox.askyesno("이력 취소", f"다음 이력을 취소하시겠습니까?\n\n{info}\n\n재고가 원복됩니다."):
            return

        def process():
            success, msg = self.app.db.cancel_history(row_no)
            self.app.root.after(0, lambda: self._show_result(success, msg))

        threading.Thread(target=process, daemon=True).start()

    def _show_result(self, success, msg):
        if success:
            messagebox.showinfo("취소 완료", msg)
            self._load_data()
        else:
            messagebox.showerror("취소 실패", msg)
