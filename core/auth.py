# -*- coding: utf-8 -*-
"""
재고관리 시스템 - 인증 및 세션 관리 모듈
"""

import hashlib


def hash_password(password: str) -> str:
    """비밀번호를 SHA-256으로 해싱"""
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def verify_password(password: str, hashed: str) -> bool:
    """입력 비밀번호와 저장된 해시 비교"""
    return hash_password(password) == hashed


class Session:
    """
    현재 로그인된 사용자 세션 (앱 종료 시 자동 소멸)
    user_data = {
        "아이디":   "admin",
        "이름":     "황승준",
        "메뉴권한": ["dashboard", "parts", "products", ...]
    }
    """
    _user: dict | None = None

    @classmethod
    def login(cls, user_data: dict):
        cls._user = user_data

    @classmethod
    def logout(cls):
        cls._user = None

    @classmethod
    def is_logged_in(cls) -> bool:
        return cls._user is not None

    @classmethod
    def current_user(cls) -> dict | None:
        return cls._user

    @classmethod
    def user_name(cls) -> str:
        return cls._user.get("이름", "") if cls._user else ""

    @classmethod
    def user_id(cls) -> str:
        return cls._user.get("아이디", "") if cls._user else ""

    @classmethod
    def has_menu(cls, menu_id: str) -> bool:
        """해당 메뉴 접근 권한이 있는지 확인"""
        if not cls._user:
            return False
        return menu_id in cls._user.get("메뉴권한", [])

    @classmethod
    def can_manage_users(cls) -> bool:
        """사용자 관리 권한 (users 메뉴 포함 여부)"""
        return cls.has_menu("users")

    @classmethod
    def permitted_menus(cls) -> list:
        """접근 가능한 메뉴 ID 목록 (조회 권한)"""
        if not cls._user:
            return []
        all_perms = cls._user.get("메뉴권한", [])
        return [m for m in all_perms if not m.endswith("_w")]

    @classmethod
    def has_write(cls, menu_id: str) -> bool:
        """해당 메뉴의 쓰기(입력/수정/삭제/처리) 권한 확인.
        관리자(users 권한)는 항상 True.
        """
        if not cls._user:
            return False
        if "users" in cls._user.get("메뉴권한", []):
            return True   # 관리자는 모든 쓰기 허용
        return f"{menu_id}_w" in cls._user.get("메뉴권한", [])

    @classmethod
    def permitted_writes(cls) -> list:
        """쓰기 권한이 있는 메뉴 ID 목록 (_w 접미사 제거한 형태)"""
        if not cls._user:
            return []
        if "users" in cls._user.get("메뉴권한", []):
            from core.constants import WRITE_MENU_IDS
            return [m.replace("_w", "") for m in WRITE_MENU_IDS]
        all_perms = cls._user.get("메뉴권한", [])
        return [m.replace("_w", "") for m in all_perms if m.endswith("_w")]

    @classmethod
    def is_admin(cls) -> bool:
        """관리자(users 권한) 여부"""
        return cls.has_menu("users")
