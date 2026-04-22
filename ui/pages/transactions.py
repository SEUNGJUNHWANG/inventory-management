"""
재고관리 시스템 - 입고 / 출고 / 생산 페이지
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
from core.constants import COLORS, FONT_FAMILY, FONT_SIZES


class ReceivePage:
    """부품 입고 페이지"""
    def __init__(self, app):
        self.app = app

    def render(self):
        card = self.app._create_card("📥 부품 입고")

        tk.Label(card, text="바코드 스캐너로 품번을 스캔하거나 직접 입력하세요.",
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"],
                 font=(FONT_FAMILY, FONT_SIZES["small"])).pack(anchor="w", padx=20, pady=(0, 10))

        form = tk.Frame(card, bg=COLORS["card_bg"])
        form.pack(fill=tk.X, padx=20, pady=5)

        tk.Label(form, text="품번:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.part_id = tk.Entry(form, font=(FONT_FAMILY, FONT_SIZES["body"]), width=25)
        self.part_id.grid(row=0, column=1, padx=5, pady=5)
        self.part_id.focus_set()

        self.part_info = tk.Label(form, text="", bg=COLORS["card_bg"],
                                  fg=COLORS["text_secondary"], font=(FONT_FAMILY, FONT_SIZES["small"]))
        self.part_info.grid(row=0, column=2, padx=10)

        self.part_id.bind("<KeyRelease>", self._on_part_id_change)
        self.part_id.bind("<Return>", lambda e: self.qty.focus_set())

        tk.Label(form, text="수량:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.qty = tk.Entry(form, font=(FONT_FAMILY, FONT_SIZES["body"]), width=15)
        self.qty.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        tk.Label(form, text="비고:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).grid(row=2, column=0, sticky="e", padx=5, pady=5)
        self.note = tk.Entry(form, font=(FONT_FAMILY, FONT_SIZES["small"]), width=40)
        self.note.grid(row=2, column=1, columnspan=2, padx=5, pady=5)

        btn_frame = tk.Frame(card, bg=COLORS["card_bg"])
        btn_frame.pack(fill=tk.X, padx=20, pady=15)
        tk.Button(btn_frame, text="입고 처리", font=(FONT_FAMILY, FONT_SIZES["body_large"], "bold"),
                  bg=COLORS["success"], fg="white", padx=25, pady=8,
                  cursor="hand2", command=self._do_receive).pack(side=tk.LEFT)

        self.result_label = tk.Label(card, text="", bg=COLORS["card_bg"],
                                     font=(FONT_FAMILY, FONT_SIZES["body_large"]), wraplength=600, justify="left")
        self.result_label.pack(anchor="w", padx=20, pady=(0, 15))

    def _on_part_id_change(self, event=None):
        part_id = self.part_id.get().strip()
        if len(part_id) >= 2:
            def lookup():
                part = self.app.db.get_part_by_id(part_id)
                if part:
                    self.app.root.after(0, lambda: self.part_info.configure(
                        text=f"→ {part['부품명']} (현재재고: {part['현재재고']}개)", fg=COLORS["primary"]))
                else:
                    self.app.root.after(0, lambda: self.part_info.configure(
                        text="미등록 품번", fg=COLORS["danger"]))
            threading.Thread(target=lookup, daemon=True).start()

    def _do_receive(self):
        part_id = self.part_id.get().strip()
        qty_str = self.qty.get().strip()
        note = self.note.get().strip()

        if not part_id or not qty_str:
            messagebox.showwarning("입력 오류", "품번과 수량을 입력해 주세요.")
            return
        try:
            qty = int(qty_str)
            if qty <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("입력 오류", "수량은 양의 정수를 입력해 주세요.")
            return

        def process():
            success, msg = self.app.db.receive_part(part_id, qty, note)
            self.app.root.after(0, lambda: self._show_result(success, msg))

        threading.Thread(target=process, daemon=True).start()

    def _show_result(self, success, msg):
        color = COLORS["success"] if success else COLORS["danger"]
        self.result_label.configure(text=msg, fg=color)
        if success:
            self.part_id.delete(0, tk.END)
            self.qty.delete(0, tk.END)
            self.note.delete(0, tk.END)
            self.part_id.focus_set()
            self.part_info.configure(text="")


class IssuePage:
    """부품 출고 페이지"""
    def __init__(self, app):
        self.app = app

    def render(self):
        card = self.app._create_card("📤 부품 출고 (개별)")

        tk.Label(card, text="바코드 스캐너로 품번을 스캔하거나 직접 입력하세요.",
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"],
                 font=(FONT_FAMILY, FONT_SIZES["small"])).pack(anchor="w", padx=20, pady=(0, 10))

        form = tk.Frame(card, bg=COLORS["card_bg"])
        form.pack(fill=tk.X, padx=20, pady=5)

        tk.Label(form, text="품번:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.part_id = tk.Entry(form, font=(FONT_FAMILY, FONT_SIZES["body"]), width=25)
        self.part_id.grid(row=0, column=1, padx=5, pady=5)
        self.part_id.focus_set()

        self.part_info = tk.Label(form, text="", bg=COLORS["card_bg"],
                                  fg=COLORS["text_secondary"], font=(FONT_FAMILY, FONT_SIZES["small"]))
        self.part_info.grid(row=0, column=2, padx=10)

        self.part_id.bind("<KeyRelease>", self._on_part_id_change)
        self.part_id.bind("<Return>", lambda e: self.qty.focus_set())

        tk.Label(form, text="수량:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.qty = tk.Entry(form, font=(FONT_FAMILY, FONT_SIZES["body"]), width=15)
        self.qty.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        tk.Label(form, text="비고:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).grid(row=2, column=0, sticky="e", padx=5, pady=5)
        self.note = tk.Entry(form, font=(FONT_FAMILY, FONT_SIZES["small"]), width=40)
        self.note.grid(row=2, column=1, columnspan=2, padx=5, pady=5)

        btn_frame = tk.Frame(card, bg=COLORS["card_bg"])
        btn_frame.pack(fill=tk.X, padx=20, pady=15)
        tk.Button(btn_frame, text="출고 처리", font=(FONT_FAMILY, FONT_SIZES["body_large"], "bold"),
                  bg=COLORS["warning"], fg="white", padx=25, pady=8,
                  cursor="hand2", command=self._do_issue).pack(side=tk.LEFT)

        self.result_label = tk.Label(card, text="", bg=COLORS["card_bg"],
                                     font=(FONT_FAMILY, FONT_SIZES["body_large"]), wraplength=600, justify="left")
        self.result_label.pack(anchor="w", padx=20, pady=(0, 15))

    def _on_part_id_change(self, event=None):
        part_id = self.part_id.get().strip()
        if len(part_id) >= 2:
            def lookup():
                part = self.app.db.get_part_by_id(part_id)
                if part:
                    self.app.root.after(0, lambda: self.part_info.configure(
                        text=f"→ {part['부품명']} (현재재고: {part['현재재고']}개)", fg=COLORS["primary"]))
                else:
                    self.app.root.after(0, lambda: self.part_info.configure(
                        text="미등록 품번", fg=COLORS["danger"]))
            threading.Thread(target=lookup, daemon=True).start()

    def _do_issue(self):
        part_id = self.part_id.get().strip()
        qty_str = self.qty.get().strip()
        note = self.note.get().strip()

        if not part_id or not qty_str:
            messagebox.showwarning("입력 오류", "품번과 수량을 입력해 주세요.")
            return
        try:
            qty = int(qty_str)
            if qty <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("입력 오류", "수량은 양의 정수를 입력해 주세요.")
            return

        def process():
            success, msg = self.app.db.issue_part(part_id, qty, note)
            self.app.root.after(0, lambda: self._show_result(success, msg))

        threading.Thread(target=process, daemon=True).start()

    def _show_result(self, success, msg):
        color = COLORS["success"] if success else COLORS["danger"]
        self.result_label.configure(text=msg, fg=color)
        if success:
            self.part_id.delete(0, tk.END)
            self.qty.delete(0, tk.END)
            self.note.delete(0, tk.END)
            self.part_id.focus_set()
            self.part_info.configure(text="")


class ProducePage:
    """제품 생산 (BOM 기반 자동 출고) 페이지"""
    def __init__(self, app):
        self.app = app

    def render(self):
        card = self.app._create_card("🏭 제품 생산 (BOM 기반 자동 출고)")

        tk.Label(card, text="제품 코드와 생산 수량을 입력하면, BOM에 따라 필요한 부품이 자동으로 출고됩니다.",
                 bg=COLORS["card_bg"], fg=COLORS["text_secondary"],
                 font=(FONT_FAMILY, FONT_SIZES["small"])).pack(anchor="w", padx=20, pady=(0, 10))

        form = tk.Frame(card, bg=COLORS["card_bg"])
        form.pack(fill=tk.X, padx=20, pady=5)

        tk.Label(form, text="제품코드:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.product_id = tk.Entry(form, font=(FONT_FAMILY, FONT_SIZES["body"]), width=25)
        self.product_id.grid(row=0, column=1, padx=5, pady=5)
        self.product_id.focus_set()

        self.product_info = tk.Label(form, text="", bg=COLORS["card_bg"],
                                     fg=COLORS["text_secondary"], font=(FONT_FAMILY, FONT_SIZES["small"]))
        self.product_info.grid(row=0, column=2, padx=10)

        self.product_id.bind("<KeyRelease>", self._on_product_id_change)
        self.product_id.bind("<Return>", lambda e: self.qty.focus_set())

        tk.Label(form, text="생산수량:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.qty = tk.Entry(form, font=(FONT_FAMILY, FONT_SIZES["body"]), width=15)
        self.qty.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        tk.Label(form, text="비고:", bg=COLORS["card_bg"],
                 font=(FONT_FAMILY, FONT_SIZES["small"], "bold")).grid(row=2, column=0, sticky="e", padx=5, pady=5)
        self.note = tk.Entry(form, font=(FONT_FAMILY, FONT_SIZES["small"]), width=40)
        self.note.grid(row=2, column=1, columnspan=2, padx=5, pady=5)

        btn_frame = tk.Frame(card, bg=COLORS["card_bg"])
        btn_frame.pack(fill=tk.X, padx=20, pady=15)
        tk.Button(btn_frame, text="🏭 생산 처리", font=(FONT_FAMILY, FONT_SIZES["body_large"], "bold"),
                  bg=COLORS["danger"], fg="white", padx=25, pady=8,
                  cursor="hand2", command=self._do_produce).pack(side=tk.LEFT)

        self.result_text = tk.Text(card, font=(FONT_FAMILY, FONT_SIZES["small"]), height=10, width=70,
                                   state="disabled", wrap="word")
        self.result_text.pack(padx=20, pady=(0, 15))

    def _on_product_id_change(self, event=None):
        product_id = self.product_id.get().strip()
        if len(product_id) >= 2:
            def lookup():
                product = self.app.db.get_product_by_id(product_id)
                if product:
                    bom = self.app.db.get_bom_for_product(product_id)
                    info = f"→ {product['제품명']} (현재재고: {product['현재재고']}개, BOM: {len(bom)}개 부품)"
                    self.app.root.after(0, lambda: self.product_info.configure(
                        text=info, fg=COLORS["primary"]))
                else:
                    self.app.root.after(0, lambda: self.product_info.configure(
                        text="미등록 제품코드", fg=COLORS["danger"]))
            threading.Thread(target=lookup, daemon=True).start()

    def _do_produce(self):
        product_id = self.product_id.get().strip()
        qty_str = self.qty.get().strip()
        note = self.note.get().strip()

        if not product_id or not qty_str:
            messagebox.showwarning("입력 오류", "제품코드와 생산수량을 입력해 주세요.")
            return
        try:
            qty = int(qty_str)
            if qty <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("입력 오류", "수량은 양의 정수를 입력해 주세요.")
            return

        if not messagebox.askyesno("생산 확인",
                f"제품코드 '{product_id}' {qty}개 생산을 진행하시겠습니까?\nBOM에 따라 부품이 자동 출고됩니다."):
            return

        def process():
            success, msg, warnings = self.app.db.produce_product(product_id, qty, note)
            self.app.root.after(0, lambda: self._show_result(success, msg, warnings))

        threading.Thread(target=process, daemon=True).start()

    def _show_result(self, success, msg, warnings):
        self.result_text.configure(state="normal")
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert("1.0", msg)
        if warnings:
            self.result_text.insert(tk.END, "\n\n" + "\n".join(warnings))
        self.result_text.configure(state="disabled")

        if success:
            self.product_id.delete(0, tk.END)
            self.qty.delete(0, tk.END)
            self.note.delete(0, tk.END)
            self.product_id.focus_set()
            self.product_info.configure(text="")
            if warnings:
                messagebox.showwarning("안전재고 경고", "\n".join(warnings))
