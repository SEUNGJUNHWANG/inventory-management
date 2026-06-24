# -*- coding: utf-8 -*-
"""
재고관리 시스템 - 안전 런처
import 오류를 포함한 모든 시작 오류를 startup_error.log 에 기록합니다.
"""
import sys
import os
import traceback

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_PATH = os.path.join(BASE_DIR, "startup_error.log")

if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)


def show_error(msg):
    """tkinter 메시지박스로 오류 표시 (tkinter 자체가 없으면 조용히 건너뜀)"""
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("시작 오류", msg)
        root.destroy()
    except Exception:
        pass


def write_log(text):
    try:
        with open(LOG_PATH, "w", encoding="utf-8") as f:
            f.write(text)
    except Exception:
        pass


try:
    # main.py 의 모든 import + InventoryApp 실행을 한 번에 감쌈
    import runpy
    runpy.run_path(os.path.join(BASE_DIR, "main.py"), run_name="__main__")

except Exception:
    tb = traceback.format_exc()
    write_log(tb)
    short = tb.strip().split("\n")[-1]  # 마지막 오류 줄만 발췌
    show_error(f"앱 시작 중 오류가 발생했습니다.\n\n{short}\n\n자세한 내용: startup_error.log")
    sys.exit(1)
