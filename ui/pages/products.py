"""
재고관리 시스템 - 제품 관리 페이지
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
from core.constants import COLORS, FONT_FAMILY, FONT_SIZES, PRODUCTS_COLUMNS


class ProductsPage:
    def __init__(self, app):
        self.app = app
        self.products_tree = None
        self.products_menu = None

    def render(self):
        scroll_frame = self.app._create_scrollable_frame()

        header = tk.Frame(scroll_frame, bg=COLORS["bg"])
        header.pack(fill=tk.X, padx=5, pady=(0, 10))
        tk.Label(header, text="📦 제품 관리", bg=COLORS["bg"],
                 fg=COLORS["text"], font=(FONT_FAMILY, FONT_SIZES["title"], "bold")).pack(side=tk.LEFT)
        tk.Button(header, text="+ 제품 추가", font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                  bg=COLORS["primary"], fg="white", padx=15, pady=5,
                  cursor="hand2", command=self._add_product_dialog).pack(side=tk.RIGHT)

        card = tk.Frame(scroll_frame, bg=COLORS["card_bg"],
                        highlightbackground=COLORS["border"], highlightthickness=1)
        card.pack(fill=tk.BOTH, expand=True, padx=5)

        self.products_tree = ttk.Treeview(card, columns=PRODUCTS_COLUMNS, show="headings", height=20)
        for col in PRODUCTS_COLUMNS:
            self.products_tree.heading(col, text=col)
            self.products_tree.column(col, width=150, anchor="center")
        self.products_tree.column("제품명", width=250)
        self.products_tree.column("제품코드", width=120)
        self.products_tree.column("규격", width=150)
        self.products_tree.column("현재재고", width=100)

        prod_scroll = ttk.Scrollbar(card, orient="vertical", command=self.products_tree.yview)
        self.products_tree.configure(yscrollcommand=prod_scroll.set)
        self.products_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        prod_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 반응형: 비고 컬럼 너비 자동 조정
        def _on_tree_resize(event):
            total = self.products_tree.winfo_width()
            fixed = 120+250+150+100+20
            remaining = max(80, total - fixed)
            self.products_tree.column("비고", width=remaining)
        self.products_tree.bind("<Configure>", _on_tree_resize)

        self.products_menu = tk.Menu(self.app.root, tearoff=0)
        self.products_menu.add_command(label="수정", command=self._edit_product_dialog)
        self.products_menu.add_command(label="삭제", command=self._delete_product)
        self.products_tree.bind("<Button-3>", self._right_click)

        self._load_data()

    def _load_data(self):
        def load():
            try:
                products = self.app.db.get_all_products()
                self.app.root.after(0, lambda: render(products))
            except Exception as e:
                err_msg = str(e)
                self.app.root.after(0, lambda: messagebox.showerror("오류", err_msg))

        def render(products):
            self.products_tree.delete(*self.products_tree.get_children())
            for p in products:
                self.products_tree.insert("", "end", values=(
                    p.get("제품코드", ""), p.get("제품명", ""), p.get("규격", ""),
                    p.get("현재재고", 0), p.get("비고", ""),
                ))

        threading.Thread(target=load, daemon=True).start()

    def _right_click(self, event):
        item = self.products_tree.identify_row(event.y)
        if item:
            self.products_tree.selection_set(item)
            self.products_menu.post(event.x_root, event.y_root)

    def _add_product_dialog(self):
        dialog = tk.Toplevel(self.app.root)
        dialog.title("제품 추가")
        dialog.geometry("400x300")
        dialog.resizable(False, False)
        dialog.transient(self.app.root)
        dialog.grab_set()

        fields = {}
        labels = [("제품코드", ""), ("제품명", ""), ("규격", ""), ("현재재고", "0"), ("비고", "")]

        for i, (label, default) in enumerate(labels):
            tk.Label(dialog, text=label + ":", font=(FONT_FAMILY, FONT_SIZES["small"])).grid(
                row=i, column=0, padx=10, pady=5, sticky="e")
            entry = tk.Entry(dialog, font=(FONT_FAMILY, FONT_SIZES["small"]), width=30)
            entry.grid(row=i, column=1, padx=10, pady=5)
            entry.insert(0, default)
            fields[label] = entry

        def save():
            try:
                self.app.db.add_product(
                    fields["제품코드"].get().strip(), fields["제품명"].get().strip(),
                    fields["규격"].get().strip(), int(fields["현재재고"].get()),
                    fields["비고"].get().strip(),
                )
                messagebox.showinfo("성공", "제품이 추가되었습니다.")
                dialog.destroy()
                self._load_data()
            except Exception as e:
                messagebox.showerror("오류", str(e))

        tk.Button(dialog, text="저장", font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                  bg=COLORS["primary"], fg="white", padx=20, pady=5,
                  command=save).grid(row=len(labels), column=0, columnspan=2, pady=15)

    def _edit_product_dialog(self):
        selected = self.products_tree.selection()
        if not selected:
            return
        values = self.products_tree.item(selected[0])["values"]

        dialog = tk.Toplevel(self.app.root)
        dialog.title("제품 수정")
        dialog.geometry("400x300")
        dialog.resizable(False, False)
        dialog.transient(self.app.root)
        dialog.grab_set()

        fields = {}
        labels = [("제품코드", str(values[0])), ("제품명", str(values[1])), ("규격", str(values[2])),
                  ("현재재고", str(values[3])), ("비고", str(values[4]))]

        for i, (label, default) in enumerate(labels):
            tk.Label(dialog, text=label + ":", font=(FONT_FAMILY, FONT_SIZES["small"])).grid(
                row=i, column=0, padx=10, pady=5, sticky="e")
            entry = tk.Entry(dialog, font=(FONT_FAMILY, FONT_SIZES["small"]), width=30)
            entry.grid(row=i, column=1, padx=10, pady=5)
            entry.insert(0, default)
            if label == "제품코드":
                entry.configure(state="readonly")
            fields[label] = entry

        def save():
            try:
                self.app.db.update_product(
                    fields["제품코드"].get().strip(), fields["제품명"].get().strip(),
                    fields["규격"].get().strip(), int(fields["현재재고"].get()),
                    fields["비고"].get().strip(),
                )
                messagebox.showinfo("성공", "제품 정보가 수정되었습니다.")
                dialog.destroy()
                self._load_data()
            except Exception as e:
                messagebox.showerror("오류", str(e))

        tk.Button(dialog, text="저장", font=(FONT_FAMILY, FONT_SIZES["small"], "bold"),
                  bg=COLORS["primary"], fg="white", padx=20, pady=5,
                  command=save).grid(row=len(labels), column=0, columnspan=2, pady=15)

    def _delete_product(self):
        selected = self.products_tree.selection()
        if not selected:
            return
        values = self.products_tree.item(selected[0])["values"]
        if messagebox.askyesno("삭제 확인", f"'{values[1]}({values[0]})' 제품을 삭제하시겠습니까?"):
            try:
                self.app.db.delete_product(str(values[0]))
                messagebox.showinfo("성공", "제품이 삭제되었습니다.")
                self._load_data()
            except Exception as e:
                messagebox.showerror("오류", str(e))
