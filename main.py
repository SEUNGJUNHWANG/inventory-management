# -*- coding: utf-8 -*-
"""
재고관리 시스템 - 메인 애플리케이션
반응형 UI: 화면 해상도에 맞게 창 크기 및 레이아웃 자동 조정
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox

# ─────────────────────────────────────────
# 경로 설정 (EXE 빌드 시에도 정상 동작)
# ─────────────────────────────────────────
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 모듈 임포트를 위해 BASE_DIR을 sys.path에 추가
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from core.constants import (
    APP_TITLE, APP_VERSION, APP_NAME,
    COLORS, FONT_FAMILY, FONT_SIZES,
    WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT,
    SIDEBAR_WIDTH, MENU_ITEMS,
)
from core.config import load_config, save_config, get_log_path
from core.database import GoogleSheetsDB
from core.updater import check_for_updates

# 페이지 모듈 임포트
from ui.pages.dashboard import DashboardPage
from ui.pages.parts import PartsPage
from ui.pages.products import ProductsPage
from ui.pages.bom import BomPage
from ui.pages.transactions import ReceivePage, IssuePage, ProducePage
from ui.pages.mrp import MrpPage
from ui.pages.history import HistoryPage
from ui.pages.report import ReportPage
from ui.pages.settings import SettingsPage


class InventoryApp:
    """재고관리 시스템 메인 애플리케이션 (반응형 UI)"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title(APP_TITLE)
        self.root.configure(bg=COLORS["bg"])

        # ── 화면 해상도 감지 및 창 크기 최적화 ──
        self._setup_window_size()

        # 상태 변수
        self.db = None
        self.connected = False
        self.current_page = None
        self.sidebar_buttons = {}

        # 스타일 설정
        self._setup_styles()

        # 레이아웃 구성
        self._build_layout()

        # 페이지 인스턴스 생성
        self.pages = {
            "dashboard": DashboardPage(self),
            "parts": PartsPage(self),
            "products": ProductsPage(self),
            "bom": BomPage(self),
            "receive": ReceivePage(self),
            "issue": IssuePage(self),
            "produce": ProducePage(self),
            "mrp": MrpPage(self),
            "history": HistoryPage(self),
            "report": ReportPage(self),
            "settings": SettingsPage(self),
        }

        # 자동 연결 시도
        self._auto_connect()

        # 초기 페이지 표시
        if self.connected:
            self._show_page("dashboard")
        else:
            self._show_page("settings")

        # 업데이트 확인 (백그라운드, 3초 딜레이 후 실행)
        self.root.after(3000, lambda: check_for_updates(self.root, silent=True))

    # ─────────────────────────────────────────
    # 창 크기 최적화
    # ─────────────────────────────────────────
    def _setup_window_size(self):
        """화면 해상도를 감지하여 최적 창 크기 설정"""
        # 화면 크기 감지
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()

        # 작업 표시줄 등을 고려한 유효 영역 (화면의 95%)
        usable_w = int(screen_w * 0.95)
        usable_h = int(screen_h * 0.95)

        # 최소 크기 보장
        win_w = max(WINDOW_MIN_WIDTH, usable_w)
        win_h = max(WINDOW_MIN_HEIGHT, usable_h)

        # 화면 중앙에 배치
        x = (screen_w - win_w) // 2
        y = (screen_h - win_h) // 2

        self.root.geometry(f"{win_w}x{win_h}+{x}+{y}")
        self.root.minsize(WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT)

        # 화면 크기 저장 (다른 메서드에서 참조)
        self.screen_w = screen_w
        self.screen_h = screen_h
        self.win_w = win_w
        self.win_h = win_h

    # ─────────────────────────────────────────
    # 스타일 설정
    # ─────────────────────────────────────────
    def _setup_styles(self):
        """ttk 스타일 설정 (화면 크기에 따라 행 높이 조정)"""
        style = ttk.Style()
        style.theme_use("clam")

        # 화면 높이에 따라 트리뷰 행 높이 조정
        row_height = 26 if self.screen_h <= 768 else 28

        style.configure("Treeview",
                        font=(FONT_FAMILY, FONT_SIZES["small"]),
                        rowheight=row_height,
                        background="white",
                        fieldbackground="white")
        style.configure("Treeview.Heading",
                        font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                        background="#e2e8f0",
                        foreground="#1e293b")
        style.map("Treeview",
                  background=[("selected", "#dbeafe")],
                  foreground=[("selected", "#1e293b")])

        # 스크롤바 스타일
        style.configure("Vertical.TScrollbar", width=12)
        style.configure("Horizontal.TScrollbar", width=12)

    # ─────────────────────────────────────────
    # 레이아웃 구성
    # ─────────────────────────────────────────
    def _build_layout(self):
        """메인 레이아웃 구성 (사이드바 + 콘텐츠 영역) - 반응형"""
        # ── 사이드바 ──
        # 화면 크기에 따라 사이드바 너비 조정
        sidebar_w = SIDEBAR_WIDTH if self.screen_w >= 1440 else max(180, SIDEBAR_WIDTH - 20)

        self.sidebar = tk.Frame(self.root, bg=COLORS["sidebar"], width=sidebar_w)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y)
        self.sidebar.pack_propagate(False)

        # ── 로고/타이틀 영역 ──
        # 화면 높이에 따라 로고 패딩 조정
        logo_pady_top = 15 if self.screen_h <= 900 else 20
        logo_pady_bottom = 15 if self.screen_h <= 900 else 25
        logo_icon_size = 22 if self.screen_h <= 900 else 26

        logo_frame = tk.Frame(self.sidebar, bg=COLORS["sidebar"])
        logo_frame.pack(fill=tk.X, pady=(logo_pady_top, logo_pady_bottom))

        tk.Label(logo_frame, text="📦", bg=COLORS["sidebar"],
                 font=(FONT_FAMILY, logo_icon_size)).pack()
        tk.Label(logo_frame, text=APP_NAME, bg=COLORS["sidebar"],
                 fg=COLORS["sidebar_text"],
                 font=(FONT_FAMILY, FONT_SIZES["body_large"], "bold")).pack()
        tk.Label(logo_frame, text=f"v{APP_VERSION}", bg=COLORS["sidebar"],
                 fg=COLORS["text_secondary"],
                 font=(FONT_FAMILY, FONT_SIZES["tiny"])).pack()

        # ── 메뉴 버튼 ──
        # 화면 높이에 따라 메뉴 버튼 패딩 자동 계산
        # 전체 메뉴 수: 11개, 로고영역 약 100px, 연결상태 약 40px
        # 남은 높이를 메뉴 버튼에 균등 배분
        num_menus = len(MENU_ITEMS)
        logo_area = logo_pady_top + logo_pady_bottom + logo_icon_size * 2 + 50
        status_area = 50
        available_h = self.win_h - logo_area - status_area
        menu_btn_h = max(6, min(12, available_h // num_menus - 4))

        # 메뉴 폰트 크기도 화면 크기에 따라 조정
        menu_font_size = FONT_SIZES["small"] if self.screen_h <= 900 else FONT_SIZES["body"]

        for page_id, label in MENU_ITEMS:
            btn = tk.Label(
                self.sidebar, text=f"  {label}",
                bg=COLORS["sidebar"], fg=COLORS["sidebar_text"],
                font=(FONT_FAMILY, menu_font_size),
                anchor="w", padx=12, pady=menu_btn_h, cursor="hand2",
            )
            btn.pack(fill=tk.X)
            btn.bind("<Enter>", lambda e, b=btn, pid=page_id:
                     b.configure(bg=COLORS["sidebar_hover"]) if pid != self.current_page else None)
            btn.bind("<Leave>", lambda e, b=btn, pid=page_id:
                     b.configure(bg=COLORS["sidebar"]) if pid != self.current_page else None)
            btn.bind("<Button-1>", lambda e, pid=page_id: self._show_page(pid))
            self.sidebar_buttons[page_id] = btn

        # ── 연결 상태 표시 ──
        self.conn_label = tk.Label(
            self.sidebar, text="● 미연결",
            bg=COLORS["sidebar"], fg=COLORS["danger"],
            font=(FONT_FAMILY, FONT_SIZES["tiny"]),
        )
        self.conn_label.pack(side=tk.BOTTOM, fill=tk.X, padx=15, pady=10)

        # ── 콘텐츠 영역 ──
        self.content_frame = tk.Frame(self.root, bg=COLORS["bg"])
        self.content_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # ─────────────────────────────────────────
    # 자동 연결
    # ─────────────────────────────────────────
    def _auto_connect(self):
        """저장된 설정으로 자동 연결 시도"""
        config = load_config()
        json_path = config.get("json_key_path", "")
        sheet_url = config.get("spreadsheet_url", "")

        if json_path and os.path.exists(json_path) and sheet_url:
            try:
                self.db = GoogleSheetsDB(json_path, sheet_url)
                self.connected = True
                self.conn_label.configure(text="● 연결됨", fg=COLORS["success"])
            except Exception as e:
                self.connected = False
                try:
                    with open(get_log_path(), "a", encoding="utf-8") as f:
                        f.write(f"Auto-connect failed: {e}\n")
                except:
                    pass

    # ─────────────────────────────────────────
    # 페이지 전환
    # ─────────────────────────────────────────
    def _show_page(self, page_id):
        """페이지 전환"""
        # 사이드바 버튼 상태 업데이트
        for pid, btn in self.sidebar_buttons.items():
            if pid == page_id:
                btn.configure(bg=COLORS["sidebar_active"], fg="white")
            else:
                btn.configure(bg=COLORS["sidebar"], fg=COLORS["sidebar_text"])

        self.current_page = page_id

        # 콘텐츠 영역 초기화
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        # 연결 확인 (설정 페이지는 예외)
        if page_id != "settings" and not self.connected:
            self._show_not_connected()
            return

        # 페이지 렌더링
        try:
            page = self.pages.get(page_id)
            if page:
                page.render()
            else:
                tk.Label(self.content_frame, text=f"페이지 '{page_id}'를 찾을 수 없습니다.",
                         bg=COLORS["bg"], fg=COLORS["danger"],
                         font=(FONT_FAMILY, FONT_SIZES["heading"])).pack(pady=50)
        except Exception as e:
            error_msg = str(e)
            tk.Label(self.content_frame, text=f"페이지 로드 오류: {error_msg}",
                     bg=COLORS["bg"], fg=COLORS["danger"],
                     font=(FONT_FAMILY, FONT_SIZES["body"])).pack(pady=50)
            try:
                with open(get_log_path(), "a", encoding="utf-8") as f:
                    import traceback
                    f.write(f"Page load error ({page_id}): {traceback.format_exc()}\n")
            except:
                pass

    def _show_not_connected(self):
        """미연결 상태 안내"""
        frame = tk.Frame(self.content_frame, bg=COLORS["bg"])
        frame.pack(expand=True)
        tk.Label(frame, text="🔗", bg=COLORS["bg"],
                 font=(FONT_FAMILY, 48)).pack(pady=(0, 10))
        tk.Label(frame, text="구글 시트에 연결되지 않았습니다.",
                 bg=COLORS["bg"], fg=COLORS["text"],
                 font=(FONT_FAMILY, FONT_SIZES["heading"], "bold")).pack(pady=5)
        tk.Label(frame, text="설정 메뉴에서 구글 시트 연결을 먼저 진행해 주세요.",
                 bg=COLORS["bg"], fg=COLORS["text_secondary"],
                 font=(FONT_FAMILY, FONT_SIZES["body"])).pack(pady=5)
        tk.Button(frame, text="⚙️ 설정으로 이동",
                  font=(FONT_FAMILY, FONT_SIZES["body_large"], "bold"),
                  bg=COLORS["primary"], fg="white", padx=20, pady=10,
                  cursor="hand2",
                  command=lambda: self._show_page("settings")).pack(pady=20)

    # ─────────────────────────────────────────
    # 공통 UI 헬퍼 메서드 (페이지에서 self.app.xxx로 호출)
    # ─────────────────────────────────────────
    def _create_card(self, title=""):
        """카드 위젯 생성"""
        card = tk.Frame(self.content_frame, bg=COLORS["card_bg"],
                        highlightbackground=COLORS["border"], highlightthickness=1)
        card.pack(fill=tk.X, padx=15, pady=(15, 5))

        if title:
            tk.Label(card, text=title, bg=COLORS["card_bg"],
                     fg=COLORS["text"],
                     font=(FONT_FAMILY, FONT_SIZES["heading"], "bold")).pack(
                anchor="w", padx=20, pady=(15, 10))
        return card

    def _create_scrollable_frame(self):
        """스크롤 가능한 프레임 생성 (반응형: 캔버스 너비 자동 조정)"""
        container = tk.Frame(self.content_frame, bg=COLORS["bg"])
        container.pack(fill=tk.BOTH, expand=True)

        # 세로 스크롤바
        v_scrollbar = ttk.Scrollbar(container, orient="vertical")
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 캔버스
        canvas = tk.Canvas(container, bg=COLORS["bg"], highlightthickness=0,
                            yscrollcommand=v_scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scrollbar.config(command=canvas.yview)

        # 스크롤 가능한 내부 프레임
        scrollable = tk.Frame(canvas, bg=COLORS["bg"])
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor="nw")

        # 내부 프레임 크기 변경 시 스크롤 영역 업데이트
        def _on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        scrollable.bind("<Configure>", _on_frame_configure)

        # 캔버스 너비 변경 시 내부 프레임 너비 자동 조정 (반응형 핵심)
        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)

        canvas.bind("<Configure>", _on_canvas_configure)

        # 마우스 휠 스크롤
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        return scrollable

    def get_content_width(self):
        """현재 콘텐츠 영역 너비 반환 (트리뷰 컬럼 자동 조정에 사용)"""
        self.root.update_idletasks()
        w = self.content_frame.winfo_width()
        return w if w > 100 else self.win_w - self.sidebar.winfo_width()

    def run(self):
        """앱 실행"""
        self.root.mainloop()


# ─────────────────────────────────────────
# 엔트리 포인트
# ─────────────────────────────────────────
if __name__ == "__main__":
    try:
        app = InventoryApp()
        app.run()
    except Exception as e:
        import traceback
        try:
            log_path = os.path.join(BASE_DIR, "startup_error.log")
            with open(log_path, "w", encoding="utf-8") as f:
                f.write(traceback.format_exc())
        except:
            pass
        messagebox.showerror("시작 오류", f"앱 시작 중 오류가 발생했습니다:\n{e}")
