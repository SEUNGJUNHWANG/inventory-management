"""
재고관리 시스템 - 부품 관리 페이지
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
from core.auth import Session
from core.constants import COLORS, FONT_FAMILY, FONT_SIZES, PARTS_COLUMNS
from ui.widget_utils import flash_btn


def _bind_tree_scroll(tree):
    """Treeview에 마우스 휠 스크롤 바인딩 (hover 기반)"""
    def _on_mw(e):
        try:
            tree.yview_scroll(int(-1 * (e.delta / 120)), "units")
        except Exception:
            pass
    tree.bind("<Enter>", lambda e: tree.bind_all("<MouseWheel>", _on_mw))
    tree.bind("<Leave>", lambda e: tree.unbind_all("<MouseWheel>"))

class PartsPage:
    def __init__(self, app):
        self.app = app
        self.parts_tree = None
        self.parts_search_var = None
        self.parts_menu = None
        self._action_overlay = None
        self._hovered_item = None

    def render(self):
        """부품 관리 페이지 렌더링"""
        scroll_frame = self.app._create_scrollable_frame()

        # 타이틀 & 버튼
        header = tk.Frame(scroll_frame, bg=COLORS["bg"])
        header.pack(fill=tk.X, padx=5, pady=(0, 10))
        tk.Label(header, text="🔩 부품 관리", bg=COLORS["bg"],
                 fg=COLORS["text"], font=(FONT_FAMILY, FONT_SIZES["title"], "bold")).pack(side=tk.LEFT)
        if Session.has_write("parts"):
            tk.Button(header, text="+ 부품 추가", font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                      bg=COLORS["primary"], fg="white", padx=15, pady=5,
                      cursor="hand2", command=self._add_part_dialog).pack(side=tk.RIGHT)
        if Session.has_write("parts"):
            tk.Button(header, text="📄 엑셀 대량 등록", font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                  bg=COLORS["info"], fg="white", padx=15, pady=5,
                  cursor="hand2", command=self._bulk_upload_parts).pack(side=tk.RIGHT, padx=(0, 8))
        tk.Button(header, text="📝 양식 다운로드", font=(FONT_FAMILY, FONT_SIZES["small"]),
                  bg="#e2e8f0", fg="#374151", padx=12, pady=5,
                  cursor="hand2", command=self._download_parts_template).pack(side=tk.RIGHT, padx=(0, 8))

        # 검색
        search_frame = tk.Frame(scroll_frame, bg=COLORS["bg"])
        search_frame.pack(fill=tk.X, padx=5, pady=(0, 10))
        tk.Label(search_frame, text="🔍 검색 (품번/부품명):", bg=COLORS["bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"])).pack(side=tk.LEFT)
        self.parts_search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=self.parts_search_var,
                                font=(FONT_FAMILY, FONT_SIZES["small"]), width=30)
        search_entry.pack(side=tk.LEFT, padx=5)
        search_entry.bind("<Return>", lambda e: self._load_parts_tree())
        tk.Button(search_frame, text="검색", font=(FONT_FAMILY, FONT_SIZES["tiny"]),
                  command=lambda: self._load_parts_tree()).pack(side=tk.LEFT)
        tk.Button(search_frame, text="새로고침", font=(FONT_FAMILY, FONT_SIZES["tiny"]),
                  command=lambda: (self.parts_search_var.set(""), self._load_parts_tree())).pack(side=tk.LEFT, padx=5)

        # 테이블
        card = tk.Frame(scroll_frame, bg=COLORS["card_bg"],
                        highlightbackground=COLORS["border"], highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True, padx=5)

        self.parts_tree = ttk.Treeview(card, columns=PARTS_COLUMNS, show="headings", height=20)
        for col in PARTS_COLUMNS:
            self.parts_tree.heading(col, text=col)
            self.parts_tree.column(col, width=100, anchor="center")
        self.parts_tree.column("품번", width=120)
        self.parts_tree.column("업체명", width=130)
        self.parts_tree.column("부품명", width=180)
        self.parts_tree.column("규격", width=130)
        self.parts_tree.column("단위", width=55)
        self.parts_tree.column("단가", width=90)
        self.parts_tree.column("현재재고", width=80)
        self.parts_tree.column("안전재고", width=80)
        self.parts_tree.column("MOQ", width=70)
        self.parts_tree.column("상태", width=70)
        self.parts_tree.column("비고", width=150)

        parts_scroll = ttk.Scrollbar(card, orient="vertical", command=self.parts_tree.yview)
        self.parts_tree.configure(yscrollcommand=parts_scroll.set)
        self.parts_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        parts_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 반응형: 창 크기 변경 시 비고 컬럼 너비 자동 조정
        def _on_tree_resize(event):
            total = self.parts_tree.winfo_width()
            fixed = 120+130+180+130+55+90+80+80+70+70+20  # 고정 컬럼 합계 + 스크롤바
            remaining = max(80, total - fixed)
            self.parts_tree.column("비고", width=remaining)
        self.parts_tree.bind("<Configure>", _on_tree_resize)
        _bind_tree_scroll(self.parts_tree)

        # 더블클릭 → 수정
        self.parts_tree.bind("<Double-1>", lambda e: self._edit_part_dialog())

        # 우클릭 메뉴 (기존 유지)
        self.parts_menu = tk.Menu(self.app.root, tearoff=0)
        if Session.has_write("parts"):
            self.parts_menu.add_command(label="✏️ 수정", command=self._edit_part_dialog)
        if Session.has_write("parts"):
            self.parts_menu.add_command(label="🗑️ 삭제", command=self._delete_part)
        self.parts_tree.bind("<Button-3>", self._parts_right_click)

        # ── 인라인 액션 오버레이 (행 호버 시 우측에 ✏️ 🗑️ 표시) ──
        self._action_overlay = tk.Frame(
            card, bg="#f1f5f9", relief="solid", bd=1,
            cursor="arrow"
        )
        btn_edit = tk.Label(
            self._action_overlay, text="✏️", bg="#f1f5f9",
            font=(FONT_FAMILY, 12), cursor="hand2", padx=6,
        )
        btn_edit.pack(side=tk.LEFT)
        btn_del = tk.Label(
            self._action_overlay, text="🗑️", bg="#f1f5f9",
            font=(FONT_FAMILY, 12), cursor="hand2", padx=6,
        )
        btn_del.pack(side=tk.LEFT)

        btn_edit.bind("<ButtonRelease-1>", lambda e: self._edit_part_dialog())
        btn_del.bind("<ButtonRelease-1>",  lambda e: self._delete_part())

        # 오버레이 위에 머물러도 숨기지 않음
        self._action_overlay.bind("<Enter>", lambda e: None)
        self._action_overlay.bind("<Leave>", lambda e: self._action_overlay.place_forget())
        btn_edit.bind("<Enter>", lambda e: btn_edit.configure(bg="#dbeafe"))
        btn_edit.bind("<Leave>", lambda e: btn_edit.configure(bg="#f1f5f9"))
        btn_del.bind("<Enter>",  lambda e: btn_del.configure(bg="#fee2e2"))
        btn_del.bind("<Leave>",  lambda e: btn_del.configure(bg="#f1f5f9"))

        self.parts_tree.bind("<Motion>",  self._on_tree_hover)
        self.parts_tree.bind("<Leave>",   lambda e: self._action_overlay.place_forget())

        self._load_parts_tree()

    def _load_parts_tree(self):
        def load():
            try:
                parts = self.app.db.get_all_parts()
                self.app.root.after(0, lambda: render(parts))
            except Exception as e:
                err_msg = str(e)
                self.app.root.after(0, lambda: messagebox.showerror("오류", err_msg))

        def render(parts):
            self.parts_tree.delete(*self.parts_tree.get_children())
            search = self.parts_search_var.get().strip().lower()
            for p in parts:
                if search and search not in str(p.get("품번", "")).lower() and search not in str(p.get("부품명", "")).lower():
                    continue
                current = int(p.get("현재재고", 0))
                safety = int(p.get("안전재고", 0))
                status = "정상"
                if safety > 0 and current <= safety:
                    status = "⚠️ 부족"
                elif safety > 0 and current <= safety * 1.2:
                    status = "주의"

                unit_price = p.get("단가", 0)
                moq = p.get("MOQ", 0)
                self.parts_tree.insert("", "end", values=(
                    p.get("품번", ""),
                    p.get("업체명", ""),
                    p.get("부품명", ""),
                    p.get("규격", ""),
                    p.get("단위", ""),
                    f"{int(unit_price):,}" if unit_price else "-",
                    current,
                    safety,
                    int(moq) if moq else 0,
                    status,
                    p.get("비고", ""),
                ))

        threading.Thread(target=load, daemon=True).start()

    def _on_tree_hover(self, event):
        """행 위에 마우스가 있을 때 우측에 액션 버튼 오버레이 표시"""
        if not self._action_overlay:
            return
        item = self.parts_tree.identify_row(event.y)
        if not item:
            self._action_overlay.place_forget()
            return
        bbox = self.parts_tree.bbox(item)
        if not bbox:
            self._action_overlay.place_forget()
            return
        _, row_y, _, row_h = bbox
        tree_x = self.parts_tree.winfo_x()
        tree_y = self.parts_tree.winfo_y()
        tree_w = self.parts_tree.winfo_width()
        overlay_w = 76
        scrollbar_w = 18
        x = tree_x + tree_w - overlay_w - scrollbar_w
        y = tree_y + row_y
        self._action_overlay.place(x=x, y=y, width=overlay_w, height=max(row_h, 24))
        self._action_overlay.lift()
        self._hovered_item = item
        # 오버레이 선택 동기화
        self.parts_tree.selection_set(item)

    def _parts_right_click(self, event):
        item = self.parts_tree.identify_row(event.y)
        if item:
            self.parts_tree.selection_set(item)
            self.parts_menu.post(event.x_root, event.y_root)

    def _add_part_dialog(self):
        dialog = tk.Toplevel(self.app.root)
        dialog.title("부품 추가")
        self.app.center_dialog(dialog, 400, 450)
        dialog.resizable(False, False)
        dialog.transient(self.app.root)
        dialog.grab_set()

        fields = {}
        labels = [("품번", ""), ("업체명", ""), ("부품명", ""), ("규격", ""), ("단위", "EA"),
                  ("단가", "0"), ("현재재고", "0"), ("안전재고", "0"), ("MOQ", "0"), ("비고", "")]

        for i, (label, default) in enumerate(labels):
            tk.Label(dialog, text=label + ":", font=(FONT_FAMILY, FONT_SIZES["small"])).grid(
                row=i, column=0, padx=10, pady=5, sticky="e")
            entry = tk.Entry(dialog, font=(FONT_FAMILY, FONT_SIZES["small"]), width=30)
            entry.grid(row=i, column=1, padx=10, pady=5)
            entry.insert(0, default)
            fields[label] = entry

        def save():
            try:
                self.app.db.add_part(
                    fields["품번"].get().strip(), fields["부품명"].get().strip(),
                    fields["규격"].get().strip(), fields["단위"].get().strip(),
                    int(fields["현재재고"].get()), int(fields["안전재고"].get()),
                    fields["비고"].get().strip(), fields["업체명"].get().strip(),
                    float(fields["단가"].get() or 0), int(fields["MOQ"].get() or 0),
                )
                messagebox.showinfo("성공", "부품이 추가되었습니다.")
                dialog.destroy()
                self._load_parts_tree()
            except Exception as e:
                messagebox.showerror("오류", str(e))

        save_btn = tk.Button(dialog, text="저장", font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                             bg=COLORS["primary"], fg="white", padx=20, pady=5,
                             command=save)
        save_btn.grid(row=len(labels), column=0, columnspan=2, pady=15)
        dialog.bind("<Return>", lambda e: flash_btn(save_btn, save))

    def _edit_part_dialog(self):
        selected = self.parts_tree.selection()
        if not selected:
            return
        values = self.parts_tree.item(selected[0])["values"]

        dialog = tk.Toplevel(self.app.root)
        dialog.title("부품 수정")
        self.app.center_dialog(dialog, 420, 640)
        dialog.resizable(False, False)
        dialog.transient(self.app.root)
        dialog.grab_set()

        fields = {}
        labels = [("품번", str(values[0])), ("업체명", str(values[1])), ("부품명", str(values[2])),
                  ("규격", str(values[3])), ("단위", str(values[4])),
                  ("단가", str(values[5]).replace(",", "").replace("-", "0")),
                  ("현재재고", str(values[6])), ("안전재고", str(values[7])),
                  ("MOQ", str(values[8])), ("비고", str(values[10]))]

        for i, (label, default) in enumerate(labels):
            tk.Label(dialog, text=label + ":", font=(FONT_FAMILY, FONT_SIZES["small"])).grid(
                row=i, column=0, padx=10, pady=5, sticky="e")
            entry = tk.Entry(dialog, font=(FONT_FAMILY, FONT_SIZES["small"]), width=30)
            entry.grid(row=i, column=1, padx=10, pady=5)
            entry.insert(0, default)
            if label == "품번":
                entry.configure(state="readonly")
            fields[label] = entry

        # ── 단가 변경사유 (선택 입력) ──
        sep_row = len(labels)
        tk.Frame(dialog, height=1, bg=COLORS.get("border", "#e2e8f0")).grid(
            row=sep_row, column=0, columnspan=2, sticky="ew", padx=10, pady=(4, 0))

        reason_row = sep_row + 1
        tk.Label(dialog, text="변경사유:",
                 font=(FONT_FAMILY, FONT_SIZES["small"]),
                 fg=COLORS.get("text_secondary", "#6b7280")).grid(
            row=reason_row, column=0, padx=10, pady=5, sticky="e")
        reason_entry = tk.Entry(dialog, font=(FONT_FAMILY, FONT_SIZES["small"]),
                                width=30, fg="#374151")
        reason_entry.grid(row=reason_row, column=1, padx=10, pady=5)
        tk.Label(dialog, text="(단가 변경 시에만 기록됩니다. 선택 입력)",
                 font=(FONT_FAMILY, 8), fg="#9ca3af").grid(
            row=reason_row + 1, column=0, columnspan=2, pady=(0, 4))

        def save():
            try:
                changed_by = ""
                try:
                    from core.auth import get_current_user
                    user = get_current_user()
                    changed_by = user.get("name", "") if user else ""
                except Exception:
                    pass

                self.app.db.update_part(
                    fields["품번"].get().strip(), fields["부품명"].get().strip(),
                    fields["규격"].get().strip(), fields["단위"].get().strip(),
                    int(fields["현재재고"].get()), int(fields["안전재고"].get()),
                    fields["비고"].get().strip(), fields["업체명"].get().strip(),
                    float(fields["단가"].get() or 0), int(fields["MOQ"].get() or 0),
                    changed_by=changed_by,
                    change_reason=reason_entry.get().strip(),
                )
                messagebox.showinfo("성공", "부품 정보가 수정되었습니다.")
                dialog.destroy()
                self._load_parts_tree()
            except Exception as e:
                messagebox.showerror("오류", str(e))

        save_btn = tk.Button(dialog, text="저장", font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                             bg=COLORS["primary"], fg="white", padx=20, pady=5,
                             command=save)
        save_btn.grid(row=reason_row + 2, column=0, columnspan=2, pady=15)
        dialog.bind("<Return>", lambda e: flash_btn(save_btn, save))

    def _delete_part(self):
        selected = self.parts_tree.selection()
        if not selected:
            return
        values = self.parts_tree.item(selected[0])["values"]
        part_id = str(values[0])
        part_name = str(values[1])

        if messagebox.askyesno("삭제 확인", f"'{part_name}({part_id})' 부품을 삭제하시겠습니까?"):
            try:
                self.app.db.delete_part(part_id)
                messagebox.showinfo("성공", "부품이 삭제되었습니다.")
                self._load_parts_tree()
            except Exception as e:
                messagebox.showerror("오류", str(e))

    def _download_parts_template(self):
        filepath = filedialog.asksaveasfilename(
            title="부품 대량 등록 양식 저장", defaultextension=".xlsx",
            initialfile="부품_대량등록_양식.xlsx", filetypes=[("Excel 파일", "*.xlsx")])
        if not filepath:
            return
        try:
            from utils.excel_utils import create_parts_template
            create_parts_template(filepath)
            messagebox.showinfo("완료",
                f"양식 파일이 저장되었습니다.\n{filepath}\n\n"
                "양식에 맞춰 부품 정보를 입력한 후,\n'엑셀 대량 등록' 버튼으로 업로드하세요.")
        except Exception as e:
            messagebox.showerror("오류", str(e))

    def _bulk_upload_parts(self):
        if not self.app.connected:
            messagebox.showwarning("연결 필요", "먼저 구글 시트에 연결해 주세요.")
            return

        filepath = filedialog.askopenfilename(
            title="부품 대량 등록 엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx"), ("모든 파일", "*.*")])
        if not filepath:
            return

        try:
            from utils.excel_utils import parse_parts_excel
            parts_to_upload = parse_parts_excel(filepath)
        except Exception as e:
            messagebox.showerror("파일 오류", f"엑셀 파일을 읽는 중 오류가 발생했습니다.\n{e}")
            return

        if not parts_to_upload:
            messagebox.showwarning("데이터 없음",
                "엑셀 파일에서 등록할 부품 데이터를 찾을 수 없습니다.\n"
                "품번과 부품명이 모두 입력되어 있는지 확인해 주세요.")
            return

        # 미리보기 대화상자
        preview = tk.Toplevel(self.app.root)
        preview.title("부품 대량 등록 미리보기")
        sw = self.app.root.winfo_screenwidth()
        sh = self.app.root.winfo_screenheight()
        win_w = max(1000, int(sw * 0.7))
        win_h = max(650, int(sh * 0.75))
        self.app.center_dialog(preview, win_w, win_h)
        preview.minsize(900, 600)
        preview.transient(self.app.root)
        preview.grab_set()

        tk.Label(preview, text=f"엑셀 파일에서 {len(parts_to_upload)}건의 부품을 찾았습니다.",
                 font=(FONT_FAMILY, 12, "bold")).pack(pady=(15, 5))
        tk.Label(preview, text="※ 품번이 기존에 있으면 업데이트, 없으면 신규 등록됩니다.",
                 font=(FONT_FAMILY, FONT_SIZES["small"]), fg="#6b7280").pack(pady=(0, 10))

        cols = ("품번", "부품명", "규격", "단위", "업체명", "현재재고", "안전재고", "비고")
        tree_frame = tk.Frame(preview)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=15)
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor="center")
        tree.column("부품명", width=180)
        tree.column("규격", width=180)

        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=tree_scroll.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        for p in parts_to_upload:
            tree.insert("", "end", values=(
                p["품번"], p["부품명"], p["규격"], p["단위"],
                p.get("업체명", ""), p["현재재고"], p["안전재고"], p["비고"]))

        status_label = tk.Label(preview, text="", font=(FONT_FAMILY, FONT_SIZES["small"]))
        status_label.pack(pady=5)

        btn_frame = tk.Frame(preview)
        btn_frame.pack(pady=(5, 15))

        def do_upload():
            upload_btn.configure(state="disabled")
            cancel_btn.configure(state="disabled")
            status_label.configure(text="업로드 중... (시간이 걸릴 수 있습니다)", fg="#d97706")
            preview.update()

            def upload_thread():
                try:
                    total = len(parts_to_upload)

                    def progress_cb(msg):
                        self.app.root.after(0, lambda m=msg: status_label.configure(text=m))

                    progress_cb(f"업로드 준비 중... (총 {total}건)")
                    new_count, update_count = self.app.db.bulk_add_or_update_parts(
                        parts_to_upload, progress_callback=progress_cb)

                    result_msg = f"✅ 완료! 신규 {new_count}건, 업데이트 {update_count}건"
                    self.app.root.after(0, lambda: status_label.configure(text=result_msg, fg="#059669"))
                    self.app.root.after(0, lambda: messagebox.showinfo("대량 등록 완료",
                        f"부품 대량 등록이 완료되었습니다.\n\n"
                        f"신규 등록: {new_count}건\n업데이트: {update_count}건\n합계: {total}건"))
                    self.app.root.after(0, lambda: cancel_btn.configure(state="normal", text="닫기"))
                    self.app.root.after(0, lambda: self._load_parts_tree())

                except Exception as e:
                    import traceback
                    err_msg = str(e)
                    tb_msg = traceback.format_exc()
                    try:
                        log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                                '..', '..', 'upload_error.log')
                        with open(log_path, 'w', encoding='utf-8') as f:
                            f.write(f'Error: {err_msg}\n\nTraceback:\n{tb_msg}')
                    except:
                        pass
                    self.app.root.after(0, lambda: status_label.configure(
                        text=f"❌ 업로드 실패: {err_msg}", fg="#dc2626"))
                    self.app.root.after(0, lambda: upload_btn.configure(state="normal"))
                    self.app.root.after(0, lambda: cancel_btn.configure(state="normal"))

            threading.Thread(target=upload_thread, daemon=True).start()

        upload_btn = tk.Button(btn_frame, text=f"✅ {len(parts_to_upload)}건 등록/업데이트 실행",
                               font=(FONT_FAMILY, FONT_SIZES["body"], "bold"),
                               bg="#059669", fg="white", padx=20, pady=8,
                               cursor="hand2", command=do_upload)
        upload_btn.pack(side=tk.LEFT, padx=10)

        cancel_btn = tk.Button(btn_frame, text="취소", font=(FONT_FAMILY, FONT_SIZES["body"]),
                               bg="#e2e8f0", fg="#374151", padx=20, pady=8,
                               cursor="hand2", command=preview.destroy)
        cancel_btn.pack(side=tk.LEFT, padx=10)
