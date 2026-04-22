"""
재고관리 시스템 - 공통 헬퍼 함수
UI 스레드 처리, 에러 로깅 등 공통 유틸리티
"""

import threading
import traceback
import os
from datetime import datetime
from core.config import get_log_path


def run_in_thread(func, callback=None, error_callback=None):
    """
    백그라운드 스레드에서 함수 실행
    - func: 실행할 함수
    - callback: 성공 시 호출 (결과를 인자로 받음)
    - error_callback: 실패 시 호출 (에러 메시지를 인자로 받음)
    """
    def wrapper():
        try:
            result = func()
            if callback:
                callback(result)
        except Exception as e:
            log_error(e)
            if error_callback:
                error_callback(str(e))

    t = threading.Thread(target=wrapper, daemon=True)
    t.start()
    return t


def log_error(exception, context=""):
    """에러를 로그 파일에 기록"""
    try:
        log_path = get_log_path()
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        tb = traceback.format_exc()
        entry = f"\n{'='*60}\n[{timestamp}] {context}\nError: {exception}\n\nTraceback:\n{tb}\n"
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(entry)
    except:
        pass


def safe_int(value, default=0):
    """안전한 정수 변환"""
    try:
        return int(value)
    except (ValueError, TypeError):
        return default


def safe_float(value, default=0.0):
    """안전한 실수 변환"""
    try:
        return float(value)
    except (ValueError, TypeError):
        return default
