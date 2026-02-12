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
import json
from datetime import datetime, timedelta
from typing import Any, Dict, Literal, Optional

import requests

from .hwp_controller import HwpController
from .hwp_table_tools import HwpTableTools, parse_table_data
from state.session_store import SessionStore
from services.diff_service import build_text_diff_summary, build_table_diff_summary

AI_SERVER_REWRITE = "http://127.0.0.1:5005/rewrite"
AI_SERVER_PLAN_TABLE = "http://127.0.0.1:5005/plan_table"
Mode = Literal["rewrite", "summarize", "extend", "table"]

# 세션 상태 (단일 문서 기준)
_current_hwp: Optional[HwpController] = None
_current_path: Optional[str] = None
_last_table_patch: Optional[Dict[str, Any]] = None

_session_store = SessionStore()
_active_preview_changeset_id: Optional[str] = None


def _normalize_patch_to_cells(patch: Dict[str, Any]) -> list[dict[str, Any]]:
    """patch를 [{row,col,old,new}] 형태 셀 변경 리스트로 정규화한다."""
    mode = patch.get("mode")
    cells: list[dict[str, Any]] = []

    if mode == "update_cells":
        for cell in patch.get("cells", []) or []:
            cells.append(
                {
                    "row": int(cell.get("row", 0)) + 1,
                    "col": int(cell.get("col", 0)) + 1,
                    "new": str(cell.get("value", "")),
                }
            )
        return cells

    if mode == "rewrite_table":
        table_data = patch.get("table", []) or []
        for r_idx, row in enumerate(table_data, start=1):
            for c_idx, new_val in enumerate(row, start=1):
                cells.append({"row": r_idx, "col": c_idx, "new": str(new_val)})
        return cells

    if mode == "update_column":
        col = int(patch.get("column", 0)) + 1
        values = patch.get("values", []) or []
        for i, val in enumerate(values, start=1):
            cells.append({"row": i, "col": col, "new": str(val)})
        return cells

    return cells


# -------- 내부 유틸 --------


def _call_ai_server(text: str, mode: Mode = "rewrite") -> str:
    if not text.strip():
        return text

    payload = {"mode": mode, "text": text}
    print(f"[ENGINE] → /rewrite payload: {payload!r}")
    resp = requests.post(AI_SERVER_REWRITE, json=payload, timeout=120)
    resp.raise_for_status()
    data = resp.json()
    print(f"[ENGINE] ← /rewrite response: {data!r}")
    return data.get("text", text) or text


def _call_table_planner(selection_text: str, instruction: str) -> dict:
    """표 관련 작업에 대해 어떤 패치를 적용할지 /plan_table에 요청한다.

    rewrite_server.py의 /plan_table 엔드포인트에 요청을 보내고,
    {"mode": ..., ...} 형태의 패치 딕셔너리를 반환한다.

    실패 시 기본값 {"mode": "rewrite_table", "table": []}을 반환한다.
    """
    payload = {
        "selection_text": selection_text,
        "instruction": instruction,
    }
    try:
        print(f"[ENGINE] → /plan_table payload: {payload!r}")
        resp = requests.post(AI_SERVER_PLAN_TABLE, json=payload, timeout=120)
        resp.raise_for_status()
        data = resp.json()
        print(f"[ENGINE] ← /plan_table response: {data!r}")
        # data 자체가 patch
        return data
    except Exception as e:
        print(f"[ENGINE] 표 플래너 호출 실패, rewrite_table 기본 패치로 폴백: {e}")
        return {"mode": "rewrite_table", "table": []}


def make_table_json_from_text(
    source_text: str,
    instr: str | None = None,
    table_rewrite: bool = False,
) -> str:
    """선택 텍스트 + 사용자 요청을 기반으로 표 JSON만 생성하는 AI 호출 헬퍼.

    - 일반 모드: 서술형 텍스트를 새 표로 정리할 때 사용
    - table_rewrite=True: 이미 존재하는 표(탭/줄바꿈 형태)를 다시 작성할 때 사용

    서버는 일반적인 "rewrite" LLM 엔드포인트라고 가정하고,
    프롬프트로 출력 형식을 강하게 제한한다.

    반환값은 JSON 배열 문자열만 (예: [["열1","열2"],["값1","값2"]]).
    """
    if not source_text.strip():
        return "[]"

    user_instr = instr.strip() if instr else ""

    # 선택 텍스트에서 대략적인 행/열 수를 추정 (table_rewrite용 힌트)
    approx_rows = 0
    approx_cols = 0
    if table_rewrite:
        lines = [l for l in source_text.splitlines() if l.strip()]
        approx_rows = len(lines)
        if lines:
            # 탭/쉼표 중 더 많이 나온 구분자로 열 수 추정
            first = lines[0]
            comma_count = first.count(",")
            tab_count = first.count("\t")
            if comma_count == 0 and tab_count == 0:
                approx_cols = 1
            else:
                if tab_count > comma_count:
                    approx_cols = len([c for c in first.split("\t")])
                else:
                    approx_cols = len([c for c in first.split(",")])

    if table_rewrite:
        # 이미 존재하는 표를 다시 작성하는 모드
        system_prompt = (
            "너는 한국어 문서의 '기존 표'를 다시 작성하는 어시스턴트야. "
            "입력 텍스트는 탭/줄바꿈으로 표현된 표라고 가정해. "
            "열 개수와 행 개수를 가능한 한 그대로 유지해야 한다."
        )

        default_request = (
            "위 표를 더 자연스럽고 명확하게 다듬어줘. "
            "첫 행은 헤더(열 이름)로 유지하고, 각 행의 구조(열 수)와 행 개수는 유지해. "
            "가능한 한 각 행의 의미(예: 날짜, 버전)는 유지하되 '내용' 같은 설명은 자연스럽게 다시 써줘."
        )

        shape_hint = []
        if approx_rows:
            shape_hint.append(
                f"- 이 표에는 현재 대략 {approx_rows}행이 있다. 출력 JSON에서도 행 수는 가능하면 {approx_rows}개로 맞춰라."
            )
        if approx_cols:
            shape_hint.append(
                f"- 첫 행(헤더)의 열 개수는 대략 {approx_cols}개다. 출력 JSON의 모든 행은 열 수를 {approx_cols}개로 맞춰라."
            )

        rules = [
            "- 출력은 JSON 배열 하나여야 한다.",
            '- 예: [["열1","열2"],["값1","값2"]]',
            "- 기존 표의 행/열 구조를 가능한 한 유지하라.",
            "- 열 이름만 한 줄만 내놓는 것은 금지다. 최소 1개 이상의 데이터 행이 있어야 한다.",
            "- 코드블록(```), 주석, 설명 문장을 절대 포함하지 마.",
        ] + shape_hint

    else:
        # 새 표를 생성하는 모드
        system_prompt = (
            "너는 한국어 문서를 표로 정리하는 어시스턴트야. "
            "반드시 JSON 배열만 출력해야 하고, 설명 문장이나 코드블록을 추가하면 안 된다."
        )

        default_request = (
            "위 텍스트를 표로 정리해줘. "
            "첫 행은 헤더(열 이름)로, 그 아래는 데이터 행으로 만들어. 최소 1개 이상의 데이터 행을 만들어야 한다."
        )

        rules = [
            "- 출력은 JSON 배열 하나여야 한다.",
            '- 예: [["열1","열2"],["값1","값2"]]',
            "- 열 이름만 한 줄만 출력하는 것은 금지다. 데이터 행이 최소 1개 이상 있어야 한다.",
            "- 코드블록(```), 주석, 설명 문장을 절대 포함하지 마.",
        ]

    full_prompt = "\n".join(
        [
            system_prompt,
            "\n[원문]",
            source_text,
            "\n[요청]",
            user_instr or default_request,
            "\n[출력 형식 규칙]",
            *rules,
        ]
    )

    raw = _call_ai_server(full_prompt, mode="table")

    # 혹시 모델이 실수로 주변에 텍스트를 넣었을 경우를 대비해
    # 가장 바깥의 '['부터 마지막 ']'까지만 잘라낸다.
    stripped = raw.strip()
    start = stripped.find("[")
    end = stripped.rfind("]")
    if start == -1 or end == -1 or end <= start:
        # 파싱 불가 시 빈 배열로 폴백
        return "[]"
    return stripped[start : end + 1]


def text_table_increment_dates(
    source_text: str,
    days: int = 1,
    header_rows: int = 1,
    date_col: int = 0,
) -> str:
    """선택된 표 텍스트의 날짜 열만 코드로 +days 만큼 이동한 JSON 표로 변환.

    전제:
    - source_text는 탭(\t)과 줄바꿈으로 구분된 표 텍스트 (HWP에서 표 복사 형태)
    - header_rows: 상단 몇 행을 헤더로 보고, 그 행들은 날짜 변환을 시도하되 실패해도 무시
    - date_col: 날짜가 위치한 열의 0 기반 인덱스 (기본: 첫 번째 열)

    지원 날짜 포맷 예시:
    - "2026. 01. 15"
    - "2026.01.15"
    - "2026-01-15"
    - "2026/01/15"

    파싱 실패한 셀은 그대로 둔다.
    반환값은 JSON 문자열([[row1],[row2],...])이다.
    """
    if not source_text.strip():
        return "[]"

    # 줄 단위로 나눈 뒤, 각 줄을 탭 또는 쉼표로 분리 (text_to_table_json과 유사한 로직)
    rows: list[list[str]] = []
    for line in source_text.splitlines():
        line = line.rstrip("\r\n")
        if not line.strip():
            continue
        # 탭이 있으면 탭 기준, 없으면 쉼표 기준, 둘 다 없으면 단일 셀
        if "\t" in line:
            cells = [c.strip() for c in line.split("\t")]
        elif "," in line:
            cells = [c.strip() for c in line.split(",")]
        else:
            cells = [line.strip()]
        rows.append(cells)

    if not rows:
        return "[]"

    # 날짜 파싱 포맷 후보
    date_formats = [
        "%Y. %m. %d",
        "%Y.%m.%d",
        "%Y-%m-%d",
        "%Y/%m/%d",
    ]

    def try_parse_date(s: str) -> datetime | None:
        s_stripped = s.strip()
        for fmt in date_formats:
            try:
                return datetime.strptime(s_stripped, fmt)
            except ValueError:
                continue
        return None

    # 날짜 증가 적용
    for row_idx, row in enumerate(rows):
        if date_col < 0 or date_col >= len(row):
            continue
        cell = row[date_col]
        dt = try_parse_date(cell)
        if dt is None:
            # 파싱 실패 시 그대로 둔다 (헤더 행 등)
            continue
        # header_rows보다 큰 행(또는 포함)부터 전부 증가시키고 싶다면 조건 조정 가능
        # 지금은 헤더 포함 전체 행에 대해 파싱 성공 시 증가 적용
        new_dt = dt + timedelta(days=days)
        # 원래 형식이 점(.)을 쓰는지, 대시/슬래시를 쓰는지에 따라 포맷 선택
        if "/" in cell:
            out = new_dt.strftime("%Y/%m/%d")
        elif "-" in cell:
            out = new_dt.strftime("%Y-%m-%d")
        else:
            # 기본은 "YYYY. MM. DD"
            out = new_dt.strftime("%Y. %m. %d")
        row[date_col] = out

    return json.dumps(rows, ensure_ascii=False)


def apply_table_patch(patch: dict, row_start: int = 1, col_start: int = 1) -> str:
    """/plan_table에서 받은 패치 JSON을 실제 HWP 표에 반영한다.

    지원하는 patch 형식:
    - {"mode": "rewrite_table", "table": [[...],[...],...]}
    - {"mode": "update_column", "column": 0, "values": [...]}
    - {"mode": "update_cells", "cells": [{"row":0,"col":2,"value":"..."}, ...]}

    row_start / col_start 는 선택된 표의 좌상단 셀(1-based)을 나타낸다.
    지금은 1,1 기준으로 사용하고, 나중에 선택된 표의 실제 시작 좌표로 확장할 수 있다.
    """
    hwp = ensure_connected()
    tools = HwpTableTools(hwp)

    mode = patch.get("mode")

    if mode == "rewrite_table":
        table = patch.get("table") or []
        if not table:
            return "Error: 빈 table 패치"
        data_list = [[str(cell) for cell in row] for row in table]
        return tools.fill_table_with_data(
            data_list=data_list,
            start_row=row_start,
            start_col=col_start,
            has_header=True,
        )

    if mode == "update_column":
        col = int(patch.get("column", 0))
        values = patch.get("values") or []
        if not values:
            return "Error: values 없음"
        for i, value in enumerate(values):
            r = row_start + i
            c = col_start + col
            tools.set_cell_text(row=r, col=c, text=str(value))
        return "열 업데이트 완료"

    if mode == "update_cells":
        cells = patch.get("cells") or []
        for cell in cells:
            r = int(cell.get("row", 0))
            c = int(cell.get("col", 0))
            v = str(cell.get("value", ""))
            tools.set_cell_text(row=row_start + r, col=col_start + c, text=v)
        return "셀 업데이트 완료"

    return f"Error: 알 수 없는 mode {mode}"


def apply_planned_table_action(selection_text: str, instruction: str) -> str:
    """LLM 플래너(/plan_table)가 결정한 패치를 현재 표에 적용한다."""
    patch = _call_table_planner(selection_text, instruction)
    return apply_table_patch(patch, row_start=1, col_start=1)


def preview_current_table_modification(instruction: str) -> str:
    """현재 표를 AI 계획 기반으로 diff 스타일 미리보기한다."""
    global _last_table_patch
    hwp = ensure_connected()

    selection_text = hwp.get_current_table_as_text()
    if not selection_text:
        return "Error: 표 안 텍스트를 읽지 못했습니다. 커서를 표 안에 두세요."

    patch = _call_table_planner(selection_text, instruction)
    normalized = _normalize_patch_to_cells(patch)
    if not normalized:
        return f"Error: 지원하지 않는 patch mode ({patch.get('mode')})"

    changed = 0
    preview_cells: list[dict[str, Any]] = []

    # 같은 좌표가 중복되면 마지막 값만 사용
    dedup: Dict[tuple[int, int], str] = {}
    for item in normalized:
        dedup[(int(item["row"]), int(item["col"]))] = str(item["new"])

    # 중요: 프리뷰 단계에서는 문서를 절대 수정하지 않는다.
    # 셀별 old/new만 계산해서 메모리에 보관하고, 적용 버튼에서만 반영한다.
    for (r, c), new_val in dedup.items():
        old_val = hwp.get_table_cell_text(r, c)
        if old_val == new_val:
            continue
        preview_cells.append({"row": r, "col": c, "old": old_val, "new": new_val})
        changed += 1

    _last_table_patch = {
        "mode": "cell_list",
        "cells": preview_cells,
    }
    return f"미리보기 완료: {changed}개 셀 (문서 미변경, 적용 시 반영)"


def finalize_table_modification() -> str:
    """미리보기 patch를 확정 적용한다."""
    global _last_table_patch
    hwp = ensure_connected()
    if not _last_table_patch:
        return "Error: 확정할 미리보기가 없습니다."

    cells = _last_table_patch.get("cells", []) or []
    for cell in cells:
        r = int(cell.get("row", 1))
        c = int(cell.get("col", 1))
        new_val = str(cell.get("new", ""))
        hwp.fill_table_cell(r, c, new_val)
    _last_table_patch = None
    return "적용 완료"


def cancel_table_modification(undo_count: int = 120) -> str:
    """미리보기 patch를 취소(undo)한다."""
    global _last_table_patch
    _ = undo_count
    # 프리뷰에서 문서를 건드리지 않으므로 취소는 상태만 비우면 된다.
    _last_table_patch = None
    return "취소 완료"


def get_last_table_preview_cells(limit: int = 20) -> list[dict[str, Any]]:
    """최근 표 미리보기 변경 목록을 반환한다."""
    if not _last_table_patch:
        return []
    cells = _last_table_patch.get("cells", []) or []
    return cells[: max(0, limit)]


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
        raise RuntimeError(
            "현재 연결된 문서가 없습니다. 먼저 파일을 선택/연결해주세요."
        )
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
    json_str ="""
    [["이름", "나이", "직업"], ["홍길동", 30, "개발자"], ["김영희", 28, "디자이너"]]
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


def create_and_fill_table_from_json(data_str: str, has_header: bool = False) -> str:
    """현재 커서 위치에 JSON 데이터로 표를 새로 만들고 채운다.

    케이스 B: 아직 표가 없고, AI가 표 데이터를 JSON으로 넘겨주는 경우에 사용.

    동작:
    1. `data_str`를 2차원 리스트로 파싱
    2. 행/열 수를 계산해서 `insert_table(rows, cols)`로 표 생성
    3. `fill_table_with_data(...)`로 표에 데이터 채우기

    Args:
        data_str: JSON 형식의 2차원 배열 문자열
        has_header: 첫 번째 행을 헤더로 처리할지 여부

    Returns:
        생성/채우기 결과를 합친 메시지 문자열
    """
    hwp = ensure_connected()

    data_list = parse_table_data(data_str)
    if not data_list:
        return "Error: 표 데이터 파싱 실패 또는 비어있는 데이터"

    # 행/열 수 계산
    rows = len(data_list)
    cols = max((len(r) for r in data_list), default=0)
    if rows == 0 or cols == 0:
        return "Error: 유효한 행/열 정보를 계산할 수 없습니다."

    table_tools = HwpTableTools(hwp)

    # 1) 표 생성
    create_msg = table_tools.insert_table(rows=rows, cols=cols)
    if "Error" in create_msg:
        return create_msg

    # 2) 데이터 채우기 (표가 방금 생성되었으므로 (1,1)부터 채움)
    fill_msg = table_tools.fill_table_with_data(
        data_list=data_list,
        start_row=1,
        start_col=1,
        has_header=has_header,
    )

    return f"{create_msg} / {fill_msg}"


def text_to_table_json(text: str) -> str:
    """자유 형식 텍스트(간단 CSV/TSV)를 표 JSON 문자열로 변환한다.

    규칙:
    - 이미 JSON 배열 형태로 시작하면(text.lstrip().startswith("[")) 그대로 반환
    - 아니면 줄 단위로 나눈 뒤, 각 줄을 쉼표(,) 또는 탭(\t) 기준으로 split
    - 빈 줄은 무시

    예시 입력:
        이름,나이,직업\n홍길동,30,개발자\n김영희,28,디자이너

    결과 JSON:
        [["이름","나이","직업"],["홍길동","30","개발자"],...]
    """
    stripped = text.strip()
    if not stripped:
        return "[]"

    # 이미 JSON 형태면 그대로 사용 (AI가 JSON으로 응답해준 경우 등)
    if stripped.startswith("["):
        return stripped

    rows = []
    for line in stripped.splitlines():
        line = line.strip()
        if not line:
            continue
        # 쉼표나 탭 중 더 많이 나온 구분자로 분리
        comma_count = line.count(",")
        tab_count = line.count("\t")
        if comma_count == 0 and tab_count == 0:
            # 구분자가 없으면 전체를 하나의 셀로 취급
            rows.append([line])
        else:
            if tab_count > comma_count:
                cells = [c.strip() for c in line.split("\t")]
            else:
                cells = [c.strip() for c in line.split(",")]
            rows.append(cells)

    return json.dumps(rows, ensure_ascii=False)


def smart_fill_table_from_json(data_str: str, has_header: bool = True) -> str:
    """커서 위치가 표 안이면 기존 표를 덮어쓰고, 표 밖이면 새 표를 생성해서 채운다.

    - 표 안인 경우(기존 템플릿 표 재작성):
      - data_str의 첫 번째 행은 "헤더 정의"로만 사용한다고 가정
      - 실제 표의 첫 행(템플릿 헤더)은 그대로 두고,
        JSON의 두 번째 행부터 데이터를 2행부터 채운다.

    - 표 밖인 경우(새 표 생성):
      - data_str 전체를 사용해서 표를 새로 만들고 채운다.
      - has_header=True이면 JSON의 첫 행을 헤더(굵게)로 처리.

    Args:
        data_str: JSON 형식의 2차원 배열 문자열
        has_header: 새 표 생성 시 첫 행을 헤더로 처리할지 여부

    Returns:
        작업 결과 메시지 문자열

    사용 예:

    ```python
    from src.tools import engine

    engine.connect_document(r"C:\path\to\doc.hwp")

    json_str ="""
    [["이름", "나이", "직업"], ["홍길동", 30, "개발자"], ["김영희", 28, "디자이너"]]
    """.strip()

    # 커서가 표 안이면: 템플릿 헤더는 유지, 2행부터 데이터 채움
    # 커서가 표 밖이면: 새 표를 만들고 1행은 헤더로 처리
    msg = engine.smart_fill_table_from_json(json_str, has_header=True)
    print(msg)
    ```
    """
    hwp = ensure_connected()
    data_list = parse_table_data(data_str)
    if not data_list:
        return "Error: 표 데이터 파싱 실패 또는 비어있는 데이터"

    table_tools = HwpTableTools(hwp)

    # 1) 커서가 표 안인지 확인
    try:
        in_table = hwp.is_cursor_in_table()
    except Exception as e:
        print(f"[ENGINE] is_cursor_in_table 체크 실패(표 밖으로 간주): {e}")
        in_table = False

    if in_table:
        # --- 케이스 A: 기존 표 재작성 모드 ---
        # JSON의 첫 행은 헤더 정의로만 쓰고, 실제 표 헤더(1행)는 그대로 둔다.
        if len(data_list) <= 1:
            return "Error: 데이터 행이 부족합니다 (헤더 외에 데이터가 필요합니다)."

        header = data_list[0]  # 필요하면 나중에 검증용으로 사용할 수 있음
        body_rows = data_list[1:]

        # 기존 표의 2행부터 데이터를 채운다.
        return table_tools.fill_table_with_data(
            data_list=body_rows,
            start_row=2,  # 2행부터 시작 → 1행(헤더)은 그대로 둠
            start_col=1,
            has_header=False,
        )
    else:
        # --- 케이스 B: 새 표 생성 모드 ---
        # 전체 데이터를 사용해서 표 생성 + 채우기.
        rows = len(data_list)
        cols = max((len(r) for r in data_list), default=0)
        if rows == 0 or cols == 0:
            return "Error: 유효한 행/열 정보를 계산할 수 없습니다."

        create_msg = table_tools.insert_table(rows=rows, cols=cols)
        if "Error" in create_msg:
            # 일부 환경에서 커서가 표 안인데 is_cursor_in_table()이 False로 판정되거나,
            # 머리글/바닥글 등 표 삽입이 불가능한 위치일 수 있다.
            # 이 경우 "기존 표 재작성" 전략으로 폴백한다.
            body_rows = data_list[1:] if len(data_list) > 1 else data_list
            return table_tools.fill_table_with_data(
                data_list=body_rows,
                start_row=2 if len(data_list) > 1 else 1,
                start_col=1,
                has_header=False,
            )

        fill_msg = table_tools.fill_table_with_data(
            data_list=data_list,
            start_row=1,
            start_col=1,
            has_header=has_header,
        )

        return f"{create_msg} / {fill_msg}"


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


def apply_text_to_selection_diff(old_text: str, new_text: str) -> None:
    """현재 선택 영역에 Diff 미리보기 텍스트를 삽입한다.

    현재 구현은 HwpController.insert_diff_text를 사용하며,
    먼저 선택 영역을 삭제한 뒤 diff 텍스트를 주입한다.
    """
    import win32api
    import win32con
    import time

    hwp = ensure_connected()

    # 선택 영역 삭제
    win32api.keybd_event(win32con.VK_DELETE, 0, 0, 0)
    win32api.keybd_event(win32con.VK_DELETE, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.05)

    # Diff 텍스트 삽입
    hwp.insert_diff_text(old_text, new_text)


# ------------------------------
# ChangeSet workflow (Phase 1)
# ------------------------------

def create_selection_changeset(instruction: str) -> str:
    selection_text = get_selection_text_via_clipboard()
    if not selection_text:
        raise RuntimeError("No selected text")

    prompt = selection_text if not instruction else f"{selection_text}\n요청: {instruction}"
    rewritten = _call_ai_server(prompt, mode="rewrite")

    diff = build_text_diff_summary(selection_text, rewritten)

    cs = _session_store.create(
        kind="text",
        prompt=instruction or "rewrite",
        before=selection_text,
        after=rewritten,
        diff=diff,
    )
    return cs.id


def preview_selection_changeset(changeset_id: str) -> str:
    global _active_preview_changeset_id
    cs = _session_store.get(changeset_id)
    if not cs:
        raise RuntimeError(f"ChangeSet not found: {changeset_id}")
    if cs.kind != "text":
        raise RuntimeError("Only text changeset supported")

    apply_text_to_selection_diff(str(cs.before), str(cs.after))
    _session_store.update_status(changeset_id, "previewed")
    _active_preview_changeset_id = changeset_id
    return "Text diff preview ready"


def create_table_changeset(instruction: str) -> str:
    hwp = ensure_connected()
    selection_text = hwp.get_current_table_as_text()
    if not selection_text:
        raise RuntimeError("Failed to read current table text")

    patch = _call_table_planner(selection_text, instruction)
    normalized = _normalize_patch_to_cells(patch)

    preview_cells: list[dict[str, Any]] = []
    dedup: Dict[tuple[int, int], str] = {}
    for item in normalized:
        dedup[(int(item["row"]), int(item["col"]))] = str(item["new"])

    for (r, c), new_val in dedup.items():
        old_val = hwp.get_table_cell_text(r, c)
        if old_val == new_val:
            continue
        preview_cells.append({"row": r, "col": c, "old": old_val, "new": new_val})

    table_diff = build_table_diff_summary(preview_cells)
    table_diff["table_cells"] = preview_cells

    cs = _session_store.create(
        kind="table",
        prompt=instruction or "table patch",
        before={"selection_text": selection_text},
        after={"cells": preview_cells},
        diff=table_diff,
    )
    return cs.id


def preview_table_changeset(changeset_id: str) -> str:
    cs = _session_store.get(changeset_id)
    if not cs:
        raise RuntimeError(f"ChangeSet not found: {changeset_id}")
    if cs.kind != "table":
        raise RuntimeError("Only table changeset supported")

    _session_store.update_status(changeset_id, "previewed")
    cells = (cs.diff or {}).get("table_cells", [])
    return f"Table preview ready: {len(cells)} cells"


def approve_changeset(changeset_id: str) -> str:
    cs = _session_store.get(changeset_id)
    if not cs:
        raise RuntimeError(f"ChangeSet not found: {changeset_id}")

    hwp = ensure_connected()

    if cs.kind == "text":
        hwp.hwp.Run("Undo")
        hwp.insert_text(str(cs.after))
        _session_store.update_status(changeset_id, "applied")
        return "텍스트 변경 적용 완료"

    if cs.kind == "table":
        cells = (cs.diff or {}).get("table_cells", [])
        for cell in cells:
            r = int(cell.get("row", 1))
            c = int(cell.get("col", 1))
            new_val = str(cell.get("new", ""))
            hwp.fill_table_cell(r, c, new_val)
        _session_store.update_status(changeset_id, "applied")
        return f"표 변경 적용 완료 ({len(cells)}개 셀)"

    raise RuntimeError(f"Unsupported changeset kind: {cs.kind}")


def reject_changeset(changeset_id: str) -> str:
    cs = _session_store.get(changeset_id)
    if not cs:
        raise RuntimeError(f"ChangeSet not found: {changeset_id}")

    hwp = ensure_connected()

    if cs.kind == "text":
        hwp.hwp.Run("Undo")
        hwp.hwp.Run("Undo")
        _session_store.update_status(changeset_id, "rejected")
        return "텍스트 변경 거절(원복) 완료"

    if cs.kind == "table":
        _session_store.update_status(changeset_id, "rejected")
        return "표 변경 거절 완료"

    raise RuntimeError(f"Unsupported changeset kind: {cs.kind}")


def get_changeset_diff_summary(changeset_id: str) -> Dict[str, Any]:
    cs = _session_store.get(changeset_id)
    if not cs:
        raise RuntimeError(f"ChangeSet not found: {changeset_id}")
    return cs.diff or {}
