"""HWP + AI 연동 유틸리티

한글 오토메이션(HwpController)과 외부 AI 서버(rewrite_server.py)를 묶어서
문서를 통째로 다듬는 helper 함수들.

- 이 모듈은 반드시 **Windows Python**에서 실행돼야 한다.
- AI 서버는 WSL 등 다른 환경에 떠 있어도 HTTP로 접근만 되면 된다.
"""

from __future__ import annotations

import os
from typing import Literal

import requests
import win32gui

from .hwp_controller import HwpController

# 필요에 따라 보스 환경에 맞게 수정 가능
AI_SERVER_REWRITE = (
    "http://127.0.0.1:5005/rewrite"  # WSL에서 돌고 있는 rewrite_server.py
)

Mode = Literal["rewrite", "summarize", "extend"]


def _call_ai_server(text: str, mode: Mode = "rewrite") -> str:
    """외부 AI 서버에 텍스트 재작성을 요청한다."""
    if not text.strip():
        return text

    payload = {"mode": mode, "text": text}
    resp = requests.post(AI_SERVER_REWRITE, json=payload, timeout=120)
    resp.raise_for_status()
    data = resp.json()
    return data.get("text", text) or text


# ---------- 1. 파일 경로 기반: 우리가 직접 여는 문서 ----------


def rewrite_document_at_path(path: str, mode: Mode = "rewrite") -> None:
    """지정한 HWP 파일을 열어서 전체 내용을 AI로 재작성 후 저장하는 v1 플로우."""
    abs_path = os.path.abspath(path)
    if not os.path.exists(abs_path):
        print(f"[ERROR] 파일을 찾을 수 없습니다: {abs_path}")
        return

    hwp = HwpController()
    if not hwp.connect(visible=True):
        print("[ERROR] 한글(HWP)에 연결하지 못했습니다.")
        return

    # 1) 문서 열기
    print(f"[INFO] 문서를 엽니다: {abs_path}")
    ok = hwp.open_document(abs_path)
    if not ok:
        print("[ERROR] 문서를 열지 못했습니다.")
        return

    # 2) 전체 텍스트 가져오기
    original = hwp.get_text()
    if not original:
        print("[WARN] 문서 텍스트를 가져오지 못했습니다.")
        return
    print(f"[INFO] 원본 문서 길이: {len(original)} 글자")

    # 3) AI 서버 호출
    try:
        rewritten = _call_ai_server(original, mode=mode)
    except Exception as e:
        print(f"[ERROR] AI 서버 호출 중 오류: {e}")
        return
    print(f"[INFO] 재작성된 문서 길이: {len(rewritten)} 글자")

    # 4) 문서 전체 교체
    try:
        hwp.select_all()
        hwp.insert_text(rewritten)
        print("[INFO] 문서 전체가 재작성되었습니다.")
    except Exception as e:
        print(f"[ERROR] 문서 교체 중 오류: {e}")
        return

    # 5) 저장
    if hwp.save_document(abs_path):
        print(f"[INFO] 문서를 저장했습니다: {abs_path}")
    else:
        print(f"[WARN] 문서 저장에 실패했습니다: {abs_path}")


# ---------- 2. 현재 열려 있는 한글 창에 붙어서 재작성 ----------


def _find_active_hwp_hwnd() -> int | None:
    """현재 열려 있는 한글(HWP) 창 중 하나의 HWND를 찾는다.

    기준:
    - CLASS에 'Hwp'가 들어가거나
    - TITLE에 '한글'이 들어가는 창
    """
    candidates: list[int] = []

    def enum_handler(hwnd, _):
        title = win32gui.GetWindowText(hwnd)
        cls = win32gui.GetClassName(hwnd)
        if (cls is not None and "Hwp" in cls) or (
            title is not None and "한글" in title
        ):
            print(
                f"[DEBUG] 후보 HWP 창 발견: HWND={hwnd} CLASS='{cls}' TITLE='{title}'"
            )
            candidates.append(hwnd)

    win32gui.EnumWindows(enum_handler, None)

    if not candidates:
        return None

    # 일단 첫 번째 후보를 사용 (필요하면 나중에 더 똑똑하게 고를 수 있음)
    return candidates[0]


def rewrite_active_hwp_window(mode: Mode = "rewrite") -> None:
    """지금 열려 있는 한글 문서(창)에 붙어서 전체 내용을 재작성한다.

    사용 전제:
    - 한글에서 편집하고 싶은 문서를 이미 열어둔 상태
    - 이 스크립트는 그 창에 붙어서 get_text → AI → 전체 교체를 수행
    """
    hwnd = _find_active_hwp_hwnd()
    if hwnd is None:
        print("[ERROR] 'Hwp' 클래스/제목을 가진 한글 창을 찾지 못했습니다.")
        return

    hwp = HwpController()
    # connect()로 COM 초기화 후, 특정 HWND에 붙기 시도
    if not hwp.connect(visible=True):
        print("[ERROR] 한글(HWP)에 연결하지 못했습니다.")
        return

    ok_conn, msg = hwp.connect_to_hwp_instance(hwnd)
    print(f"[INFO] HWND={hwnd} 한글 창에 연결 시도 결과: {msg}")
    if not ok_conn:
        print("[ERROR] 지정한 한글 창에 붙지 못했습니다.")
        return

    # 전체 텍스트 가져오기
    original = hwp.get_text()
    if not original:
        print("[WARN] 문서 텍스트를 가져오지 못했습니다.")
        return

    print(f"[INFO] (활성 창) 원본 문서 길이: {len(original)} 글자")

    # AI 호출
    try:
        rewritten = _call_ai_server(original, mode=mode)
    except Exception as e:
        print(f"[ERROR] AI 서버 호출 중 오류: {e}")
        return

    print(f"[INFO] (활성 창) 재작성된 문서 길이: {len(rewritten)} 글자")

    # 교체
    try:
        hwp.select_all()
        hwp.insert_text(rewritten)
        print("[INFO] (활성 창) 문서 전체가 재작성되었습니다.")
    except Exception as e:
        print(f"[ERROR] (활성 창) 문서 교체 중 오류: {e}")
        return


if __name__ == "__main__":
    # 간단 테스트용
    # 1) 파일 경로 기반:
    # rewrite_document_at_path(r"F:\dev\hwp-mcp-test1.hwp", mode="rewrite")
    #
    # 2) 현재 열려 있는 한글 창 기반:
    rewrite_active_hwp_window(mode="rewrite")
