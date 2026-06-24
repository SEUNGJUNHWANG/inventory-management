# -*- coding: utf-8 -*-
"""
공용 UI 유틸리티
"""


def flash_btn(btn, cmd, ms=130):
    """Enter 키 등으로 버튼을 호출할 때 시각적 눌림 효과 후 명령 실행

    Args:
        btn : tk.Button 인스턴스
        cmd : 실행할 함수 (인수 없음)
        ms  : 눌림 유지 시간 (밀리초, 기본 130)
    """
    try:
        orig = btn.cget("relief")
        btn.configure(relief="sunken")
        btn.after(ms, lambda: (btn.configure(relief=orig), cmd()))
    except Exception:
        cmd()
