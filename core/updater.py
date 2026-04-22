# -*- coding: utf-8 -*-
"""
재고관리 시스템 - 자동 업데이트 모듈
GitHub Releases에서 최신 버전을 확인하고 자동으로 업데이트합니다.
"""

import os
import sys
import threading
import subprocess
import tempfile

try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False

try:
    import tkinter as tk
    from tkinter import messagebox, ttk
except ImportError:
    pass

from core.constants import APP_VERSION, GITHUB_OWNER, GITHUB_REPO


# ─────────────────────────────────────────
# 버전 비교 유틸리티
# ─────────────────────────────────────────
def _parse_version(v_str: str) -> tuple:
    """'v2.1.0' 또는 '2.1.0' → (2, 1, 0)"""
    v_str = v_str.lstrip("v").strip()
    try:
        parts = [int(x) for x in v_str.split(".")]
        while len(parts) < 3:
            parts.append(0)
        return tuple(parts)
    except Exception:
        return (0, 0, 0)


def _is_newer(remote_version: str, local_version: str) -> bool:
    """remote 버전이 local보다 새로운지 비교"""
    return _parse_version(remote_version) > _parse_version(local_version)


# ─────────────────────────────────────────
# GitHub Releases API 조회
# ─────────────────────────────────────────
def _fetch_latest_release() -> dict | None:
    """GitHub Releases API에서 최신 릴리즈 정보 조회"""
    if not REQUESTS_AVAILABLE:
        return None

    api_url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
    try:
        resp = requests.get(api_url, timeout=10, headers={"Accept": "application/vnd.github+json"})
        if resp.status_code == 200:
            return resp.json()
        return None
    except Exception:
        return None


def _find_exe_asset(release_data: dict) -> str | None:
    """릴리즈 에셋에서 .exe 파일 다운로드 URL 추출"""
    for asset in release_data.get("assets", []):
        name = asset.get("name", "")
        if name.endswith(".exe"):
            return asset.get("browser_download_url")
    return None


# ─────────────────────────────────────────
# 업데이트 다운로드 및 교체
# ─────────────────────────────────────────
def _download_and_replace(download_url: str, root: tk.Tk):
    """새 EXE 다운로드 후 현재 EXE와 교체하고 재시작"""

    # ── 진행 창 표시 ──
    progress_win = tk.Toplevel(root)
    progress_win.title("업데이트 다운로드 중...")
    progress_win.geometry("400x120")
    progress_win.resizable(False, False)
    progress_win.grab_set()

    tk.Label(progress_win, text="새 버전을 다운로드 중입니다. 잠시 기다려 주세요...",
             font=("맑은 고딕", 10), pady=15).pack()

    bar = ttk.Progressbar(progress_win, mode="indeterminate", length=350)
    bar.pack(pady=5)
    bar.start(10)

    status_label = tk.Label(progress_win, text="", font=("맑은 고딕", 9), fg="#64748b")
    status_label.pack()

    def do_download():
        try:
            status_label.configure(text="다운로드 중...")
            progress_win.update()

            # 현재 EXE 경로
            current_exe = sys.executable if getattr(sys, "frozen", False) else None
            if not current_exe:
                root.after(0, lambda: messagebox.showwarning(
                    "알림", "개발 환경에서는 자동 업데이트를 실행할 수 없습니다."))
                root.after(0, progress_win.destroy)
                return

            exe_dir = os.path.dirname(current_exe)
            new_exe_path = os.path.join(exe_dir, "_update_new.exe")
            bat_path = os.path.join(exe_dir, "_do_update.bat")

            # 다운로드
            resp = requests.get(download_url, stream=True, timeout=120)
            resp.raise_for_status()

            total = int(resp.headers.get("content-length", 0))
            downloaded = 0
            with open(new_exe_path, "wb") as f:
                for chunk in resp.iter_content(chunk_size=65536):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if total > 0:
                            pct = int(downloaded / total * 100)
                            root.after(0, lambda p=pct: status_label.configure(
                                text=f"다운로드 중... {p}%"))

            # 교체 배치 스크립트 생성 (ANSI 인코딩으로 저장)
            bat_content = (
                "@echo off\r\n"
                "timeout /t 2 /nobreak >nul\r\n"
                f'move /y "{new_exe_path}" "{current_exe}"\r\n'
                f'start "" "{current_exe}"\r\n'
                'del "%~f0"\r\n'
            )
            with open(bat_path, "w", encoding="ansi") as f:
                f.write(bat_content)

            # 교체 스크립트 실행 후 앱 종료
            root.after(0, lambda: _launch_updater_and_quit(bat_path, root, progress_win))

        except Exception as e:
            err = str(e)
            root.after(0, progress_win.destroy)
            root.after(0, lambda: messagebox.showerror(
                "업데이트 오류", f"다운로드 중 오류가 발생했습니다:\n{err}"))

    threading.Thread(target=do_download, daemon=True).start()


def _launch_updater_and_quit(bat_path: str, root: tk.Tk, progress_win: tk.Toplevel):
    """업데이트 배치 실행 후 앱 종료"""
    try:
        progress_win.destroy()
    except Exception:
        pass
    messagebox.showinfo(
        "업데이트 준비 완료",
        "업데이트 파일 다운로드가 완료되었습니다.\n\n"
        "앱을 종료하고 새 버전으로 자동 교체합니다.\n"
        "잠시 후 앱이 다시 시작됩니다."
    )
    subprocess.Popen(
        ["cmd", "/c", bat_path],
        creationflags=subprocess.CREATE_NO_WINDOW
        if hasattr(subprocess, "CREATE_NO_WINDOW") else 0,
        close_fds=True,
    )
    root.after(500, root.quit)


# ─────────────────────────────────────────
# 공개 인터페이스: 업데이트 확인 (백그라운드)
# ─────────────────────────────────────────
def check_for_updates(root: tk.Tk, silent: bool = True):
    """
    앱 시작 시 백그라운드에서 업데이트를 확인합니다.

    Args:
        root:   tkinter 루트 창
        silent: True이면 업데이트 없을 때 아무것도 표시 안 함 (기본값)
    """
    if not REQUESTS_AVAILABLE:
        return
    # 개발 환경에서는 업데이트 체크 생략 (선택적)
    # if not getattr(sys, "frozen", False):
    #     return

    def _worker():
        release = _fetch_latest_release()
        if release is None:
            return

        latest_tag = release.get("tag_name", "")
        if not latest_tag:
            return

        if not _is_newer(latest_tag, APP_VERSION):
            if not silent:
                root.after(0, lambda: messagebox.showinfo(
                    "업데이트 확인", f"현재 최신 버전입니다. (v{APP_VERSION})"))
            return

        latest_ver = latest_tag.lstrip("v")
        download_url = _find_exe_asset(release)
        release_notes = release.get("body", "").strip()
        notes_preview = release_notes[:300] + "..." if len(release_notes) > 300 else release_notes

        def _prompt():
            msg = (
                f"새 버전 v{latest_ver}이(가) 출시되었습니다!\n"
                f"현재 버전: v{APP_VERSION}\n\n"
            )
            if notes_preview:
                msg += f"[업데이트 내용]\n{notes_preview}\n\n"
            msg += "지금 업데이트하시겠습니까?"

            if messagebox.askyesno("업데이트 알림", msg, default="yes"):
                if download_url:
                    _download_and_replace(download_url, root)
                else:
                    messagebox.showwarning(
                        "업데이트",
                        f"다운로드 파일을 찾을 수 없습니다.\n"
                        f"GitHub 페이지에서 직접 다운로드해 주세요:\n"
                        f"https://github.com/{GITHUB_OWNER}/{GITHUB_REPO}/releases"
                    )

        root.after(0, _prompt)

    threading.Thread(target=_worker, daemon=True).start()
