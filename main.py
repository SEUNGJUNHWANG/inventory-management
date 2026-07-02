# -*- coding: utf-8 -*-
"""
재고관리 시스템 - 메인 애플리케이션
반응형 UI: 창 크기 변경 시 실시간으로 레이아웃 자동 조정
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from core.constants import (
    APP_TITLE, APP_VERSION, APP_NAME,
    COLORS, FONT_FAMILY, FONT_SIZES,
    WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT,
    SIDEBAR_WIDTH, MENU_ITEMS, MENU_LABELS,
)
from core.config import load_config, save_config, get_log_path
from core.database import GoogleSheetsDB
from core.updater import check_for_updates
from core.auth import Session

from ui.pages.dashboard import DashboardPage
from ui.pages.parts import PartsPage
from ui.pages.products import ProductsPage
from ui.pages.bom import BomPage
from ui.pages.transactions import ReceivePage, IssuePage, ProducePage
from ui.pages.mrp import MrpPage
from ui.pages.history import HistoryPage
from ui.pages.report import ReportPage
from ui.pages.settings import SettingsPage
from ui.pages.login import LoginPage
from ui.pages.users import UsersPage
from ui.pages.ship import ShipPage
from ui.pages.customers import CustomersPage
from ui.pages.analytics    import AnalyticsDashboard


class InventoryApp:

    def __init__(self):
        self.root = tk.Tk()
        self.root.title(APP_TITLE)
        self.root.configure(bg=COLORS["bg"])

        self._setup_window_size()

        self.db = None
        self.connected = False
        self.current_page = None
        self.sidebar_buttons = {}
        self._resize_job = None
        self._user_info_frame = None
        self._lbl_user_name = None

        self._setup_styles()
        self._build_layout()

        self.pages = {
            "dashboard": DashboardPage(self),
            "parts":     PartsPage(self),
            "products":  ProductsPage(self),
            "bom":       BomPage(self),
            "receive":   ReceivePage(self),
            "issue":     IssuePage(self),
            "produce":   ProducePage(self),
            "mrp":       MrpPage(self),
            "history":   HistoryPage(self),
            "report":    ReportPage(self),
            "settings":  SettingsPage(self),
            "users":     UsersPage(self),
            "ship":          ShipPage(self),
            "customers":     CustomersPage(self),
            "analytics":     AnalyticsDashboard(self),
        }
        self._login_page = LoginPage(self)

        self._auto_connect()

        if self.connected:
            self._show_login()
        else:
            self._show_page("settings")

        self.root.after(3000, lambda: check_for_updates(self.root, silent=True))

    # -----------------------------------------
    # 창 크기 초기 설정
    # -----------------------------------------
    def _setup_window_size(self):
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        usable_w = int(screen_w * 0.95)
        usable_h = int(screen_h * 0.95)
        win_w = max(WINDOW_MIN_WIDTH, usable_w)
        win_h = max(WINDOW_MIN_HEIGHT, usable_h)
        x = (screen_w - win_w) // 2
        y = (screen_h - win_h) // 2
        self.root.geometry(f"{win_w}x{win_h}+{x}+{y}")
        self.root.minsize(WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT)
        self.screen_w = screen_w
        self.screen_h = screen_h
        self.win_w = win_w
        self.win_h = win_h

    # -----------------------------------------
    # 스타일 설정
    # -----------------------------------------
    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
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
        style.configure("Vertical.TScrollbar", width=12)
        style.configure("Horizontal.TScrollbar", width=12)

    # -----------------------------------------
    # 반응형 계산 헬퍼
    # -----------------------------------------
    def _calc_sidebar_w(self, win_w):
        if win_w < 1000:
            return 60
        elif win_w < 1280:
            return 180
        elif win_w < 1600:
            return 210
        else:
            return SIDEBAR_WIDTH

    def _calc_logo_metrics(self, win_h):
        if win_h < 700:
            return 18, (8, 8)
        elif win_h < 900:
            return 22, (12, 12)
        else:
            return 26, (18, 20)

    def _calc_menu_metrics(self, win_h, num_items=None):
        if num_items is None:
            num_items = len(MENU_ITEMS)
        num = max(1, num_items)
        avail = max(1, win_h - 120 - 80)
        pady = max(4, min(13, avail // num - 4))
        font_size = FONT_SIZES["small"] if win_h <= 900 else FONT_SIZES["body"]
        return pady, font_size

    # -----------------------------------------
    # 레이아웃 구성
    # -----------------------------------------
    def _build_layout(self):
        sidebar_w = self._calc_sidebar_w(self.win_w)

        self.sidebar = tk.Frame(self.root, bg=COLORS["sidebar"], width=sidebar_w)
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y)
        self.sidebar.pack_propagate(False)

        logo_icon_size, logo_pady = self._calc_logo_metrics(self.win_h)

        self._logo_frame = tk.Frame(self.sidebar, bg=COLORS["sidebar"])
        self._logo_frame.pack(fill=tk.X, pady=logo_pady)

        self._logo_icon_lbl = tk.Label(
            self._logo_frame, text="\U0001f4e6", bg=COLORS["sidebar"],
            font=(FONT_FAMILY, logo_icon_size))
        self._logo_icon_lbl.pack()

        self._logo_name_lbl = tk.Label(
            self._logo_frame, text=APP_NAME, bg=COLORS["sidebar"],
            fg=COLORS["sidebar_text"],
            font=(FONT_FAMILY, FONT_SIZES["body_large"], "bold"))
        self._logo_name_lbl.pack()

        self._logo_ver_lbl = tk.Label(
            self._logo_frame, text=f"v{APP_VERSION}",
            bg=COLORS["sidebar"], fg=COLORS["text_secondary"],
            font=(FONT_FAMILY, FONT_SIZES["tiny"]))
        self._logo_ver_lbl.pack()

        # 메뉴 버튼 컨테이너 (동적으로 재구성)
        self._menu_frame = tk.Frame(self.sidebar, bg=COLORS["sidebar"])
        self._menu_frame.pack(fill=tk.X)

        # 하단 영역 (연결상태 + 사용자 정보)
        self._bottom_frame = tk.Frame(self.sidebar, bg=COLORS["sidebar"])
        self._bottom_frame.pack(side=tk.BOTTOM, fill=tk.X)

        # 사용자 정보 + 로그아웃 (로그인 후에만 표시)
        self._user_info_frame = tk.Frame(self._bottom_frame, bg=COLORS["sidebar"])
        self._lbl_user_name = tk.Label(
            self._user_info_frame, text="", bg=COLORS["sidebar"],
            fg=COLORS["sidebar_text"],
            font=(FONT_FAMILY, FONT_SIZES["tiny"], "bold"))
        self._lbl_user_name.pack(side=tk.TOP, padx=12, pady=(6, 2))
        tk.Button(
            self._user_info_frame, text="로그아웃",
            font=(FONT_FAMILY, FONT_SIZES["tiny"]),
            bg=COLORS["sidebar_hover"], fg=COLORS["sidebar_text"],
            relief="flat", cursor="hand2", padx=8, pady=3,
            command=self.logout,
        ).pack(side=tk.TOP, padx=12, pady=(0, 6))

        self.conn_label = tk.Label(
            self._bottom_frame, text="● 미연결",
            bg=COLORS["sidebar"], fg=COLORS["danger"],
            font=(FONT_FAMILY, FONT_SIZES["tiny"]),
        )
        self.conn_label.pack(fill=tk.X, padx=15, pady=8)

        self.content_frame = tk.Frame(self.root, bg=COLORS["bg"])
        self.content_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.root.bind("<Configure>", self._on_window_resize)

    # -----------------------------------------
    # 사이드바 메뉴 재구성 (로그인 후 권한별)
    # -----------------------------------------
    def _rebuild_sidebar_menus(self):
        """Session의 허용 메뉴만 사이드바에 렌더링"""
        for w in self._menu_frame.winfo_children():
            w.destroy()
        self.sidebar_buttons.clear()

        permitted = Session.permitted_menus()  # [menu_id, ...]
        # MENU_ITEMS 순서 유지하며 허용된 것만
        visible_items = [(pid, lbl) for pid, lbl in MENU_ITEMS if pid in permitted]

        sidebar_w = self._calc_sidebar_w(self.win_w)
        icon_only = sidebar_w < 100
        menu_btn_h, menu_font_size = self._calc_menu_metrics(self.win_h, len(visible_items))

        for page_id, label in visible_items:
            display = f"  {label.split()[0]}" if icon_only else f"  {label}"
            btn = tk.Label(
                self._menu_frame, text=display,
                bg=COLORS["sidebar"], fg=COLORS["sidebar_text"],
                font=(FONT_FAMILY, menu_font_size),
                anchor="w",
                padx=8 if icon_only else 12,
                pady=menu_btn_h,
                cursor="hand2",
            )
            btn.pack(fill=tk.X)
            btn.bind("<Enter>", lambda e, b=btn, pid=page_id:
                     b.configure(bg=COLORS["sidebar_hover"]) if pid != self.current_page else None)
            btn.bind("<Leave>", lambda e, b=btn, pid=page_id:
                     b.configure(bg=COLORS["sidebar"]) if pid != self.current_page else None)
            btn.bind("<Button-1>", lambda e, pid=page_id: self._show_page(pid))
            self.sidebar_buttons[page_id] = btn

    # -----------------------------------------
    # 로그인 화면 표시
    # -----------------------------------------
    def _show_login(self):
        """로그인 화면으로 전환 (사이드바 메뉴 숨김)"""
        # 메뉴 버튼 전부 제거
        for w in self._menu_frame.winfo_children():
            w.destroy()
        self.sidebar_buttons.clear()
        # 사용자 정보 숨김
        self._user_info_frame.pack_forget()
        # content_frame 초기화 후 LoginPage 렌더
        for w in self.content_frame.winfo_children():
            w.destroy()
        self._login_page.render()

    # -----------------------------------------
    # 로그인 성공 콜백 (LoginPage에서 호출)
    # -----------------------------------------
    def on_login_success(self):
        """로그인 성공 후 사이드바 재구성 및 대시보드 이동"""
        user = Session.current_user()
        name = user.get("이름", user.get("아이디", "")) if user else ""

        # 사용자 이름 표시
        self._lbl_user_name.configure(text=f"{name} 님")
        self._user_info_frame.pack(before=self.conn_label, fill=tk.X)

        # 메뉴 재구성
        self._rebuild_sidebar_menus()

        # 대시보드로 이동 (권한 있으면), 없으면 첫 번째 메뉴
        if Session.has_menu("dashboard"):
            self._show_page("dashboard")
        elif self.sidebar_buttons:
            first = next(iter(self.sidebar_buttons))
            self._show_page(first)

    # -----------------------------------------
    # 로그아웃
    # -----------------------------------------
    def logout(self):
        """로그아웃: 세션 초기화 후 로그인 화면"""
        Session.logout()
        self._show_login()

    # -----------------------------------------
    # 창 리사이즈 핸들러 (150ms 디바운스)
    # -----------------------------------------
    def _on_window_resize(self, event):
        if event.widget != self.root:
            return
        if self._resize_job:
            self.root.after_cancel(self._resize_job)
        self._resize_job = self.root.after(
            150, lambda: self._apply_resize(event.width, event.height))

    def _apply_resize(self, new_w, new_h):
        self.win_w = new_w
        self.win_h = new_h

        new_sidebar_w = self._calc_sidebar_w(new_w)
        self.sidebar.configure(width=new_sidebar_w)
        icon_only = new_sidebar_w < 100

        visible_items = [(pid, lbl) for pid, lbl in MENU_ITEMS if pid in self.sidebar_buttons]
        menu_btn_h, menu_font_size = self._calc_menu_metrics(new_h, len(visible_items))

        for page_id, label in visible_items:
            btn = self.sidebar_buttons.get(page_id)
            if btn is None:
                continue
            display = f"  {label.split()[0]}" if icon_only else f"  {label}"
            is_active = (page_id == self.current_page)
            btn.configure(
                text=display,
                pady=menu_btn_h,
                padx=8 if icon_only else 12,
                font=(FONT_FAMILY, menu_font_size),
                bg=COLORS["sidebar_active"] if is_active else COLORS["sidebar"],
            )

        logo_icon_size, logo_pady = self._calc_logo_metrics(new_h)
        self._logo_frame.pack_configure(pady=logo_pady)
        self._logo_icon_lbl.configure(font=(FONT_FAMILY, logo_icon_size))

        if icon_only:
            self._logo_name_lbl.pack_forget()
            self._logo_ver_lbl.pack_forget()
        else:
            if not self._logo_name_lbl.winfo_ismapped():
                self._logo_name_lbl.pack()
            if not self._logo_ver_lbl.winfo_ismapped():
                self._logo_ver_lbl.pack()

        if icon_only:
            self.conn_label.configure(
                text="●",
                fg=COLORS["success"] if self.connected else COLORS["danger"])
        else:
            if self.connected:
                self.conn_label.configure(text="● 연결됨", fg=COLORS["success"])
            else:
                self.conn_label.configure(text="● 미연결", fg=COLORS["danger"])

    # -----------------------------------------
    # 자동 연결
    # -----------------------------------------
    def _auto_connect(self):
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
                except Exception:
                    pass

    # -----------------------------------------
    # 페이지 전환
    # -----------------------------------------
    def _show_page(self, page_id):
        # 로그인 상태 확인
        if not Session.is_logged_in() and page_id != "settings":
            return

        # 권한 확인 (settings 는 항상 허용)
        if page_id != "settings" and not Session.has_menu(page_id):
            messagebox.showwarning("권한 없음", "해당 메뉴에 대한 접근 권한이 없습니다.")
            return

        for pid, btn in self.sidebar_buttons.items():
            if pid == page_id:
                btn.configure(bg=COLORS["sidebar_active"], fg="white")
            else:
                btn.configure(bg=COLORS["sidebar"], fg=COLORS["sidebar_text"])
        self.current_page = page_id

        for widget in self.content_frame.winfo_children():
            widget.destroy()

        if page_id != "settings" and not self.connected:
            self._show_not_connected()
            return

        try:
            page = self.pages.get(page_id)
            if page:
                page.render()
            else:
                tk.Label(self.content_frame,
                         text=f"페이지 '{page_id}'를 찾을 수 없습니다.",
                         bg=COLORS["bg"], fg=COLORS["danger"],
                         font=(FONT_FAMILY, FONT_SIZES["heading"])).pack(pady=50)
        except Exception as e:
            error_msg = str(e)
            tk.Label(self.content_frame,
                     text=f"페이지 로드 오류: {error_msg}",
                     bg=COLORS["bg"], fg=COLORS["danger"],
                     font=(FONT_FAMILY, FONT_SIZES["body"])).pack(pady=50)
            try:
                with open(get_log_path(), "a", encoding="utf-8") as f:
                    import traceback
                    f.write(f"Page load error ({page_id}): {traceback.format_exc()}\n")
            except Exception:
                pass

    def _show_not_connected(self):
        frame = tk.Frame(self.content_frame, bg=COLORS["bg"])
        frame.pack(expand=True)
        tk.Label(frame, text="\U0001f517", bg=COLORS["bg"],
                 font=(FONT_FAMILY, 48)).pack(pady=(0, 10))
        tk.Label(frame, text="구글 시트에 연결되지 않았습니다.",
                 bg=COLORS["bg"], fg=COLORS["text"],
                 font=(FONT_FAMILY, FONT_SIZES["heading"], "bold")).pack(pady=5)
        tk.Label(frame, text="설정 메뉴에서 구글 시트 연결을 먼저 진행해 주세요.",
                 bg=COLORS["bg"], fg=COLORS["text_secondary"],
                 font=(FONT_FAMILY, FONT_SIZES["body"])).pack(pady=5)
        tk.Button(frame, text="설정으로 이동",
                  font=(FONT_FAMILY, FONT_SIZES["body_large"], "bold"),
                  bg=COLORS["primary"], fg="white", padx=20, pady=10,
                  cursor="hand2",
                  command=lambda: self._show_page("settings")).pack(pady=20)

    # -----------------------------------------
    # 공통 UI 헬퍼
    # -----------------------------------------
    def _create_card(self, title=""):
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
        container = tk.Frame(self.content_frame, bg=COLORS["bg"])
        container.pack(fill=tk.BOTH, expand=True)

        v_scrollbar = ttk.Scrollbar(container, orient="vertical")
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        canvas = tk.Canvas(container, bg=COLORS["bg"], highlightthickness=0,
                            yscrollcommand=v_scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scrollbar.config(command=canvas.yview)

        scrollable = tk.Frame(canvas, bg=COLORS["bg"])
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor="nw")

        def _on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        scrollable.bind("<Configure>", _on_frame_configure)

        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(event):
            try:
                # Treeview나 Listbox 위에서는 해당 위젯이 자체 스크롤하므로 캔버스 스크롤 건너뜀
                widget = event.widget
                if isinstance(widget, (ttk.Treeview, tk.Listbox)):
                    return
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except Exception:
                pass

        def _bind_mw(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _unbind_mw(event):
            canvas.unbind_all("<MouseWheel>")

        canvas.bind("<Enter>", _bind_mw)
        canvas.bind("<Leave>", _unbind_mw)

        return scrollable

    def get_content_width(self):
        self.root.update_idletasks()
        w = self.content_frame.winfo_width()
        return w if w > 100 else self.win_w - self.sidebar.winfo_width()

    def center_dialog(self, dialog, width, height, parent=None):
        """다이얼로그를 부모 창 기준 중앙에 위치시킨다 (화면 경계 보정 포함)"""
        p = parent if parent else self.root
        p.update_idletasks()
        px = p.winfo_rootx()
        py = p.winfo_rooty()
        pw = p.winfo_width()
        ph = p.winfo_height()
        x = px + (pw - width) // 2
        y = py + (ph - height) // 2
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = max(0, min(x, sw - width))
        y = max(20, min(y, sh - height))
        dialog.geometry(f"{width}x{height}+{x}+{y}")

    def run(self):
        self.root.mainloop()


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
        except Exception:
            pass
        messagebox.showerror("시작 오류", f"앱 시작 중 오류가 발생했습니다:\n{e}")
