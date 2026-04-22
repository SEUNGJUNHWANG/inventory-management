"""
재고관리 시스템 - 설정 파일 관리
config.json 파일을 읽고 쓰는 유틸리티
"""

import sys
import os
import json


def get_app_dir():
    """앱 루트 디렉토리 반환 (EXE 빌드 시에도 정상 동작)"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_config_path():
    """config.json 파일 경로 반환"""
    return os.path.join(get_app_dir(), "config.json")


def load_config() -> dict:
    """설정 파일 로드"""
    path = get_config_path()
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return {}
    return {}


def save_config(cfg: dict):
    """설정 파일 저장"""
    path = get_config_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def get_log_path():
    """에러 로그 파일 경로 반환"""
    return os.path.join(get_app_dir(), "error.log")
