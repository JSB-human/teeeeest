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
from .hwp_table_tools import HwpTableTools, parse_table_data

AI_SERVER_REWRITE = "http://127.0.0.1:5005/rewrite"
Mode = Literal["rewrite", "summarize", "extend"]

# 세션 상태 (단일 문서 기준)
_current_hwp: Optional[HwpController] = None
_current_path: Optional[str] = None


# -------- 내부 유틸 --------


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


def get_cursor_position_meta() -> dict | None:
    """현재 커서 위치 메타데이터를 반환한다.

    Returns:
        dict | None: {"list_id": int, "para_id": int, "char_pos": int}
    """
    hwp = ensure_connected()
    pos = hwp.get_cursor_pos()
    print("DEBUG engine.get_cursor_position_meta ->", pos)
    return pos


# -------- 표(Table) 관련 고수준 헬퍼 --------


def fill_current_table_from_json(data_str: str, has_header: bool = False) -> str:
    """현재 커서가 위치한 표에 JSON 문자열로 전달된 데이터를 채운다.

    이 함수는 `HwpTableTools`와 `parse_table_data`를 래핑해서,
    UI나 AI 레이어에서 "표 데이터만" 넘기면 되도록 단순화한 헬퍼이다.

    Args:
        data_str: JSON 형식의 2차원 배열 문자열.
            예시)
            ```json
            [
              ["이름", "나이", "직업"],
              ["홍길동", 30, "개발자"],
              ["김영희", 28, "디자이너"]
            ]
            ```
        has_header: 첫 번째 행을 헤더로 처리할지 여부.

    Returns:
        HwpTableTools.fill_table_with_data()가 반환하는 결과 메시지 문자열.

    사용 예 (엔진을 직접 쓸 때):

    ```python
    from hwp_mcp.src.tools import engine

    # (1) 먼저 문서를 연결하고, 한글에서 채울 표 안에 커서를 둔다.
    engine.connect_document(r"C:\path\to\doc.hwp")

    # (2) JSON 문자열을 준비한다. (AI가 생성해줄 수도 있음)
    json_str = """
    [
      ["이름", "나이", "직업"],
      ["홍길동", 30, "개발자"],
      ["김영희", 28, "디자이너"]
    ]
    """.strip()

    # (3) 현재 커서가 위치한 표의 (1,1) 셀부터 순서대로 채운다.
    msg = engine.fill_current_table_from_json(json_str, has_header=True)
    print(msg)  # "표 데이터 입력 완료" 등
    ```
    """
    hwp = ensure_connected()

    # JSON 문자열을 2차원 리스트로 파싱
    data_list = parse_table_data(data_str)
    if not data_list:
        return "Error: 표 데이터 파싱 실패 또는 비어있는 데이터"

    table_tools = HwpTableTools(hwp)

    # 현재 전략: 커서가 이미 해당 표 안에 위치해 있다고 가정하고,
    # 표의 (1,1)에서부터 데이터를 채운다.
    # 나중에 필요하면 start_row/start_col을 파라미터로 확장 가능.
    return table_tools.fill_table_with_data(
        data_list=data_list,
        start_row=1,
        start_col=1,
        has_header=has_header,
    )


# -------- 선택 영역 기반 v0 (클립보드 이용) --------


def get_selection_text_via_clipboard() -> str:
    """현재 한글에서 사용자가 선택한 영역의 텍스트를 클립보드로부터 가져온다.

    전제:
    - 사용자가 한글 문서에서 이미 드래그/선택을 해둔 상태
    동작:
    - 엔진이 한글 창을 활성화
    - Ctrl+C 키 입력을 보내서 선택된 내용을 클립보드에 복사
    - 클립보드에서 텍스트를 읽어 반환
    """
    import win32clipboard
    import win32gui
    import win32api
    import win32con
    import time

    hwp = ensure_connected()

    # 한글 창 활성화
    try:
        hwnd = hwp.hwp.XHwpWindows.Item(0).WindowHandle
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.1)
    except Exception as e_hwnd:
        print(f"[ENGINE] 선택 영역용 윈도우 활성화 실패(무시하고 진행): {e_hwnd}")

    # Ctrl+C 키 이벤트 전송
    try:
        win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
        win32api.keybd_event(ord("C"), 0, 0, 0)
        time.sleep(0.05)
        win32api.keybd_event(ord("C"), 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.1)
    except Exception as e_keys:
        print(f"[ENGINE] Ctrl+C 키 이벤트 전송 실패: {e_keys}")

    # 클립보드에서 텍스트 읽기
    try:
        win32clipboard.OpenClipboard()
        try:
            text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        finally:
            win32clipboard.CloseClipboard()
        return text or ""
    except Exception as e_clip:
        print(f"[ENGINE] 클립보드에서 선택 영역 텍스트 읽기 실패: {e_clip}")
        return ""


def apply_text_to_selection_via_clipboard(new_text: str) -> None:
    """현재 선택된 영역에 new_text를 덮어쓴다 (클립보드 기반 v0).

    방식:
    - new_text를 클립보드에 넣고
    - 한글 창 활성화 후 Ctrl+V 보내서 선택 영역을 덮어쓰기
    """
    import win32clipboard
    import win32gui
    import win32api
    import win32con
    import time

    if not new_text:
        print("[ENGINE] 적용할 텍스트가 비어 있습니다.")
        return

    hwp = ensure_connected()

    # 클립보드에 새 텍스트 넣기
    try:
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, new_text)
        finally:
            win32clipboard.CloseClipboard()
    except Exception as e_clip:
        print(f"[ENGINE] 클립보드에 텍스트 설정 실패: {e_clip}")
        return

    # 한글 창 활성화
    try:
        hwnd = hwp.hwp.XHwpWindows.Item(0).WindowHandle
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.1)
    except Exception as e_hwnd:
        print(f"[ENGINE] 선택 영역 적용용 윈도우 활성화 실패(무시하고 진행): {e_hwnd}")

    # Ctrl+V 키 이벤트 전송 (선택 영역 덮어쓰기)
    try:
        win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
        win32api.keybd_event(ord("V"), 0, 0, 0)
        time.sleep(0.05)
        win32api.keybd_event(ord("V"), 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
    except Exception as e_keys:
        print(f"[ENGINE] Ctrl+V 키 이벤트 전송 실패: {e_keys}")
