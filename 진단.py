# -*- coding: utf-8 -*-
"""
실행 오류 진단 스크립트
이 파일을 python 진단.py 로 실행하면 오류 원인을 찾아줍니다.
"""
import sys, os, traceback

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

errors = []

def test(label, fn):
    try:
        fn()
        print(f"  ✅ {label}")
    except Exception as e:
        msg = traceback.format_exc()
        errors.append((label, msg))
        print(f"  ❌ {label}\n     → {e}")

print("=" * 50)
print("  재고관리 시스템 - 실행 오류 진단")
print("=" * 50)

print("\n[1] 기본 모듈 임포트")
test("core.constants",  lambda: __import__("core.constants"))
test("core.config",     lambda: __import__("core.config"))
test("core.database",   lambda: __import__("core.database"))
test("core.updater",    lambda: __import__("core.updater"))

print("\n[2] UI 모듈 임포트")
test("ui.pages.dashboard",    lambda: __import__("ui.pages.dashboard"))
test("ui.pages.parts",        lambda: __import__("ui.pages.parts"))
test("ui.pages.products",     lambda: __import__("ui.pages.products"))
test("ui.pages.bom",          lambda: __import__("ui.pages.bom"))
test("ui.pages.transactions", lambda: __import__("ui.pages.transactions"))
test("ui.pages.mrp",          lambda: __import__("ui.pages.mrp"))
test("ui.pages.history",      lambda: __import__("ui.pages.history"))
test("ui.pages.report",       lambda: __import__("ui.pages.report"))
test("ui.pages.settings",     lambda: __import__("ui.pages.settings"))

print("\n[3] 설정 파일 확인")
def check_config():
    from core.config import load_config
    cfg = load_config()
    json_path = cfg.get("json_key_path", "")
    sheet_url = cfg.get("spreadsheet_url", "")
    if not json_path:
        raise ValueError("json_key_path 설정 없음")
    if not os.path.exists(json_path):
        raise FileNotFoundError(f"JSON 키 파일 없음: {json_path}")
    if not sheet_url:
        raise ValueError("spreadsheet_url 설정 없음")
    print(f"     JSON 키: {json_path}")
    print(f"     시트 URL: {sheet_url[:40]}...")
test("config.json 확인", check_config)

print("\n[4] 구글 시트 연결 테스트")
def check_db():
    from core.config import load_config
    from core.database import GoogleSheetsDB
    cfg = load_config()
    json_path = cfg.get("json_key_path", "")
    sheet_url = cfg.get("spreadsheet_url", "")
    db = GoogleSheetsDB(json_path, sheet_url)
    print(f"     연결 성공!")
test("GoogleSheetsDB 연결", check_db)

print("\n" + "=" * 50)
if errors:
    print(f"  총 {len(errors)}개 오류 발견")
    print("\n[상세 오류 내용]")
    for label, tb in errors:
        print(f"\n--- {label} ---")
        print(tb)
else:
    print("  모든 항목 정상! (오류 없음)")
print("=" * 50)
input("\n아무 키나 누르면 종료...")
