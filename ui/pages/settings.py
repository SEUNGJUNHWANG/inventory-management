"""
재고관리 시스템 - 설정 페이지 (구글 시트 연동, BOM 임포트)
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
from core.constants import COLORS, FONT_FAMILY, FONT_SIZES
from core.config import load_config, save_config


class SettingsPage:
    def __init__(self, app):
        self.app = app

    def render(self):
        scroll_frame = self.app._create_scrollable_frame()

        tk.Label(scroll_frame, text="⚙️ 설정", bg=COLORS["bg"],
                 fg=COLORS["text"], font=(FONT_FAMILY, FONT_SIZES["title"], "bold")).pack(
            anchor="w", padx=5, pady=(0, 15))

        # 구글 시트 연동 설정
        gs_card = tk.Frame(scroll_frame, bg=COLORS["card_bg"],
                           highlightbackground=COLORS["border"], highlightthickness=1)
        gs_card.pack(fill=tk.X, padx=5, pady=(0, 15))

        tk.Label(gs_card, text="📊 구글 시트 연동 설정", bg=COLORS["card_bg"],
                 fg=COLORS["text"], font=(FONT_FAMILY, FONT_SIZES["heading"], "bold")).pack(
            anchor="w", padx=20, pady=(15, 10))

        form = tk.Frame(gs_card, bg=COLORS["card_bg"])
        form.pack(fill=tk.X, padx=20, pady=5)

        config = load_config()

        # JSON 키 파일 경로
        tk.Label(form, text="서비스 계정 JSON 키 파일:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).grid(row=0, column=0, sticky="w", pady=5)
        self.json_path = tk.Entry(form, font=(FONT_FAMILY, FONT_SIZES["small"]), width=50)
        self.json_path.grid(row=1, column=0, padx=(0, 5), pady=2, sticky="ew")
        self.json_path.insert(0, config.get("json_key_path", ""))

        tk.Button(form, text="찾아보기", font=(FONT_FAMILY, 9),
                  command=self._browse_json).grid(row=1, column=1, padx=5)

        # 스프레드시트 URL
        tk.Label(form, text="구글 스프레드시트 URL:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).grid(row=2, column=0, sticky="w", pady=(10, 5))
        self.sheet_url = tk.Entry(form, font=(FONT_FAMILY, FONT_SIZES["small"]), width=50)
        self.sheet_url.grid(row=3, column=0, columnspan=2, pady=2, sticky="ew")
        self.sheet_url.insert(0, config.get("spreadsheet_url", ""))

        tk.Label(form, text="(비워두면 새 스프레드시트가 자동 생성됩니다)",
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"],
                 font=(FONT_FAMILY, 9)).grid(row=4, column=0, sticky="w")

        form.columnconfigure(0, weight=1)

        # 연결 테스트 버튼
        btn_frame = tk.Frame(gs_card, bg=COLORS["card_bg"])
        btn_frame.pack(fill=tk.X, padx=20, pady=15)

        tk.Button(btn_frame, text="🔗 연결 테스트 및 저장", font=(FONT_FAMILY, FONT_SIZES["body_large"], "bold"),
                  bg=COLORS["primary"], fg="white", padx=20, pady=8,
                  cursor="hand2", command=self._test_connection).pack(side=tk.LEFT)

        self.conn_status = tk.Label(btn_frame, text="", bg=COLORS["card_bg"],
                                    font=(FONT_FAMILY, FONT_SIZES["small"]))
        self.conn_status.pack(side=tk.LEFT, padx=15)

        # BOM 임포트 섹션
        import_card = tk.Frame(scroll_frame, bg=COLORS["card_bg"],
                               highlightbackground=COLORS["border"], highlightthickness=1)
        import_card.pack(fill=tk.X, padx=5, pady=(0, 15))

        tk.Label(import_card, text="📂 데이터 임포트", bg=COLORS["card_bg"],
                 fg=COLORS["text"], font=(FONT_FAMILY, FONT_SIZES["heading"], "bold")).pack(
            anchor="w", padx=20, pady=(15, 10))

        tk.Label(import_card, text="BOM 엑셀 파일 또는 부품 목록 엑셀 파일을 가져올 수 있습니다.",
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"],
                 font=(FONT_FAMILY, FONT_SIZES["small"])).pack(anchor="w", padx=20, pady=(0, 10))

        import_btn_frame = tk.Frame(import_card, bg=COLORS["card_bg"])
        import_btn_frame.pack(fill=tk.X, padx=20, pady=(0, 15))

        tk.Button(import_btn_frame, text="📂 BOM 엑셀 파일 선택 및 임포트",
                  font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                  bg=COLORS["success"], fg="white", padx=15, pady=8,
                  cursor="hand2", command=self._import_bom).pack(side=tk.LEFT, padx=(0, 10))

        self.import_status = tk.Label(import_card, text="", bg=COLORS["card_bg"],
                                      font=(FONT_FAMILY, FONT_SIZES["small"]))
        self.import_status.pack(anchor="w", padx=20, pady=(0, 15))

        # 앱 정보
        info_card = tk.Frame(scroll_frame, bg=COLORS["card_bg"],
                             highlightbackground=COLORS["border"], highlightthickness=1)
        info_card.pack(fill=tk.X, padx=5, pady=(0, 15))

        from core.constants import APP_VERSION, APP_NAME
        tk.Label(info_card, text=f"{APP_NAME} v{APP_VERSION}", bg=COLORS["card_bg"],
                 fg=COLORS["text_secondary"], font=(FONT_FAMILY, FONT_SIZES["small"])).pack(
            anchor="w", padx=20, pady=15)

    def _browse_json(self):
        filepath = filedialog.askopenfilename(
            title="서비스 계정 JSON 키 파일 선택",
            filetypes=[("JSON 파일", "*.json"), ("모든 파일", "*.*")],
        )
        if filepath:
            self.json_path.delete(0, tk.END)
            self.json_path.insert(0, filepath)

    def _test_connection(self):
        json_path = self.json_path.get().strip()
        sheet_url = self.sheet_url.get().strip()

        if not json_path:
            messagebox.showwarning("입력 오류", "JSON 키 파일 경로를 입력해 주세요.")
            return

        if not os.path.exists(json_path):
            messagebox.showerror("파일 오류", f"JSON 키 파일을 찾을 수 없습니다:\n{json_path}")
            return

        self.conn_status.configure(text="연결 테스트 중...", fg=COLORS["warning"])
        self.app.root.update()

        def test():
            try:
                from core.database import GoogleSheetsDB
                db = GoogleSheetsDB(json_path, sheet_url if sheet_url else None)

                # 설정 저장
                config = {
                    "json_key_path": json_path,
                    "spreadsheet_url": db.spreadsheet.url if hasattr(db, 'spreadsheet') else sheet_url,
                }
                save_config(config)

                # 앱의 DB 연결 교체
                self.app.db = db

                # URL 업데이트
                if not sheet_url and hasattr(db, 'spreadsheet'):
                    self.app.root.after(0, lambda: self._update_url(db.spreadsheet.url))

                self.app.root.after(0, lambda: self.conn_status.configure(
                    text="✅ 연결 성공!", fg=COLORS["success"]))
                self.app.root.after(0, lambda: messagebox.showinfo("성공", "구글 시트 연결에 성공했습니다!"))

            except Exception as e:
                err_msg = str(e)
                self.app.root.after(0, lambda: self.conn_status.configure(
                    text=f"❌ 연결 실패: {err_msg}", fg=COLORS["danger"]))
                self.app.root.after(0, lambda: messagebox.showerror("연결 실패", err_msg))

        threading.Thread(target=test, daemon=True).start()

    def _update_url(self, url):
        self.sheet_url.delete(0, tk.END)
        self.sheet_url.insert(0, url)

    def _import_bom(self):
        if not self.app.db:
            messagebox.showwarning("연결 필요", "먼저 구글 시트 연결을 설정해 주세요.")
            return

        filepath = filedialog.askopenfilename(
            title="BOM 엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx *.xls"), ("모든 파일", "*.*")],
        )
        if not filepath:
            return

        self.import_status.configure(text="BOM 데이터 임포트 중...", fg=COLORS["warning"])
        self.app.root.update()

        def do_import():
            try:
                from utils.excel_utils import import_bom_excel
                result = import_bom_excel(filepath, self.app.db)
                self.app.root.after(0, lambda: self.import_status.configure(
                    text=f"✅ {result}", fg=COLORS["success"]))
                self.app.root.after(0, lambda: messagebox.showinfo("임포트 완료", result))
            except Exception as e:
                err_msg = str(e)
                self.app.root.after(0, lambda: self.import_status.configure(
                    text=f"❌ 임포트 실패: {err_msg}", fg=COLORS["danger"]))
                self.app.root.after(0, lambda: messagebox.showerror("오류", err_msg))

        threading.Thread(target=do_import, daemon=True).start()
