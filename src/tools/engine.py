"""Engine layer for HwpInlineAI

이 모듈은 hwp-mcp의 HwpController와 AI 서버를 감싸서
UI(Python GUI, 웹, C# 등)가 쓰기 쉬운 함수들로 제공한다.

현재 목표:
- 하나의 문서 세션을 유지 (connect_document)
- 전체 문서 텍스트 가져오기 / 재작성하기

나중 목표:
- 선택 영역 텍스트 가져오기 / 재작성
- diff / undo 관리
"""

from __future__ import annotations

import os
from typing import Literal, Optional

import requests

from .hwp_controller import HwpController

AI_SERVER_REWRITE = "http://127.0.0.1:5005/rewrite"
Mode = Literal["rewrite", "summarize", "extend"]

# 세션 상태 (단일 문서 기준)
_current_hwp: Optional[HwpController] = None
_current_path: Optional[str] = None


def _call_ai_server(text: str, mode: Mode = "rewrite") -> str:
    if not text.strip():
        return text

    payload = {"mode": mode, "text": text}
    resp = requests.post(AI_SERVER_REWRITE, json=payload, timeout=120)
    resp.raise_for_status()
    data = resp.json()
    return data.get("text", text) or text


def connect_document(path: str, visible: bool = True) -> None:
    """한글에 연결하고 지정한 HWP 문서를 연다.

    성공하면 _current_hwp / _current_path에 세션 상태를 저장한다.
    """
    global _current_hwp, _current_path

    abs_path = os.path.abspath(path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(abs_path)

    hwp = HwpController()
    if not hwp.connect(visible=visible):
        raise RuntimeError("한글(HWP)에 연결하지 못했습니다.")

    ok = hwp.open_document(abs_path)
    if not ok:
        raise RuntimeError(f"문서를 열지 못했습니다: {abs_path}")

    _current_hwp = hwp
    _current_path = abs_path


def ensure_connected() -> HwpController:
    """현재 문서 세션이 있는지 확인하고 HwpController를 반환.

    UI 코드에서 각 작업 전에 호출해서 연결 여부를 보장한다.
    """
    global _current_hwp
    if _current_hwp is None:
        raise RuntimeError("현재 연결된 문서가 없습니다. 먼저 파일을 선택/연결해주세요.")
    return _current_hwp


def get_current_document_path() -> Optional[str]:
    return _current_path


def get_current_text() -> str:
    """현재 연결된 문서의 전체 텍스트를 반환.

    HwpController.get_text() 안에서:
    - 우선 GetTextFile("TEXT", "")를 시도하고
    - 실패 시 클립보드 fallback까지 처리하도록 해두었기 때문에
    여기서는 단순 위임만 한다.
    """
    hwp = ensure_connected()
    text = hwp.get_text()
    return text or ""


def rewrite_current_document(mode: Mode = "rewrite") -> None:
    """현재 연결된 문서 전체를 AI로 재작성해서 덮어쓴다."""
    hwp = ensure_connected()
    original = hwp.get_text()
    if not original:
        print("[WARN] 문서 텍스트를 가져오지 못했습니다.")
        return

    print(f"[ENGINE] 원본 문서 길이: {len(original)} 글자")
    try:
        rewritten = _call_ai_server(original, mode=mode)
    except Exception as e:
        print(f"[ENGINE] AI 서버 호출 중 오류: {e}")
        return

    print(f"[ENGINE] 재작성된 문서 길이: {len(rewritten)} 글자")

    try:
        hwp.select_all()
        hwp.insert_text(rewritten)
        print("[ENGINE] 문서 전체가 재작성되었습니다.")
    except Exception as e:
        print(f"[ENGINE] 문서 교체 중 오류: {e}")
        return

    # 저장은 UI에서 할지 여기서 할지 결정 가능. 지금은 즉시 저장.
    if _current_path:
        ok = hwp.save_document(_current_path)
        if ok:
            print(f"[ENGINE] 문서를 저장했습니다: {_current_path}")
        else:
            print(f"[ENGINE] 문서 저장 실패: {_current_path}")
