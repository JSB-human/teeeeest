"""Minimal AI rewrite server for HwpInlineAI

역할:
- 텍스트를 받아서 Gemini로 rewrite / summarize / extend 해주는 HTTP 서버
- 불필요한 엔드포인트 없이 /rewrite 하나만 제공

의존성:
- flask
- google-generativeai

환경변수:
- GEMINI_API_KEY   : Gemini API 키 (없으면 에코/원문 그대로 반환)
- INLINEAI_MODEL   : (선택) 모델 이름, 기본값 'gemini-2.5-flash'
- INLINEAI_HOST    : (선택) 바인딩 호스트, 기본값 '127.0.0.1'
- INLINEAI_PORT    : (선택) 포트 번호, 기본값 5005
"""

import os
from typing import Literal

from flask import Flask, request, jsonify
import json

app = Flask(__name__)

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
DEFAULT_MODEL = os.environ.get("INLINEAI_MODEL", "gemini-2.5-flash")
HOST = os.environ.get("INLINEAI_HOST", "127.0.0.1")
PORT = int(os.environ.get("INLINEAI_PORT", "5005"))

# google-genai 설정 (최신 SDK)
try:
    from google import genai
    from google.genai import types

    if GEMINI_API_KEY:
        client = genai.Client(api_key=GEMINI_API_KEY)
    else:
        client = None
        print("[WARN] GEMINI_API_KEY가 설정되지 않았습니다. 서버는 원문을 그대로 반환합니다.")
except Exception as e:
    client = None
    print(f"[WARN] google-genai를 불러올 수 없습니다: {e}. 서버는 원문을 그대로 반환합니다.")


Mode = Literal["rewrite", "summarize", "extend", "table"]


def _build_instruction(mode: Mode) -> str:
    # ... (생략되지 않음, 기존 로직 유지)
    if mode == "summarize":
        return "아래 한국어 글을 간결하게 요약해줘. 반드시 요약된 본문만 출력하고, 설명이나 불릿, 추가 설명은 쓰지 마."
    elif mode == "extend":
        return "아래 한국어 글의 흐름을 유지하면서 내용을 자연스럽게 더 확장해줘. 반드시 수정된 본문만 출력하고, 수정 이유나 설명은 쓰지 마."
    elif mode == "table":
        return ""
    else:
        return "아래 한국어 글을 의미는 유지하면서 자연스럽고 매끄럽게 다듬어줘. 반드시 다듬어진 본문만 출력하고, 수정 이유나 설명, 머리말/꼬리말은 쓰지 마."


def gemini_rewrite(text: str, mode: Mode = "rewrite") -> str:
    if not text.strip() or client is None:
        return text

    if mode == "table":
        prompt = text
    else:
        instruction = _build_instruction(mode)
        prompt = f"{instruction}\n\n--- 원문 ---\n{text}\n\n--- 수정된 글 ---"

    try:
        response = client.models.generate_content(
            model=DEFAULT_MODEL,
            contents=prompt
        )
        return response.text.strip() or text
    except Exception as e:
        print(f"[ERROR] Gemini 호출 중 오류: {e}")
        return text


@app.route("/health", methods=["GET"])
def health():
    """헬스 체크용 엔드포인트."""
    return jsonify(
        {
            "status": "ok",
            "model": DEFAULT_MODEL,
            "has_key": bool(GEMINI_API_KEY),
        }
    )


@app.route("/rewrite", methods=["POST"])
def rewrite():
    data = request.get_json(force=True) or {}
    text = str(data.get("text", ""))
    mode = str(data.get("mode", "rewrite"))

    if not text.strip():
        return jsonify({"text": text})

    if mode not in ("rewrite", "summarize", "extend", "table"):
        mode = "rewrite"

    try:
        rewritten = gemini_rewrite(text, mode)  # type: ignore[arg-type]
    except Exception as e:
        print("[ERROR] Gemini 호출 중 오류:", e)
        rewritten = text

    return jsonify({"text": rewritten})


@app.route("/plan_table", methods=["POST"])
def plan_table():
    """표 관련 작업에 대해 어떤 패치를 적용할지 계획을 세우는 엔드포인트.

    입력:
        {
          "selection_text": "...",   # 드래그된 표 텍스트 (탭/개행 기반)
          "instruction": "..."      # 사용자가 입력한 자연어 지시문
        }

    출력 (아래 셋 중 하나 형식의 JSON):

    1) 전체 표 재작성:
        {"mode": "rewrite_table", "table": [[...],[...],...]}

    2) 특정 열만 수정:
        {"mode": "update_column", "column": 0, "values": [...]}  # column은 0부터 시작

    3) 개별 셀만 수정:
        {"mode": "update_cells", "cells": [{"row":0,"col":2,"value":"..."}, ...]}
    """
    data = request.get_json(force=True) or {}
    selection = str(data.get("selection_text", ""))
    instruction = str(data.get("instruction", ""))

    if not selection.strip():
        return jsonify({"mode": "rewrite_table", "table": []})

    # LLM에게 기대하는 패치 형식 설명
    patch_spec = {
        "rewrite_table": {
            "description": "표 전체를 다시 작성할 때 사용",
            "schema": {"mode": "rewrite_table", "table": "2차원 배열"},
        },
        "update_column": {
            "description": "특정 열만 수정할 때 사용",
            "schema": {"mode": "update_column", "column": "0부터 시작 인덱스", "values": "행마다 하나씩"},
        },
        "update_cells": {
            "description": "특정 셀 몇 개만 수정할 때 사용",
            "schema": {"mode": "update_cells", "cells": "[{row,col,value}, ...]"},
        },
    }

    spec_str = json.dumps(patch_spec, ensure_ascii=False, indent=2)

    plan_prompt = f"""
너는 한글(HWP) 문서 안의 표를 다루는 도구 오케스트레이터야.

[가능한 패치 타입]
{spec_str}

[현재 선택된 표]
{selection}

[사용자 지시]
{instruction}

위 정보를 바탕으로, 아래 셋 중 하나 형태의 JSON만 출력해라:

1) 전체 표 재작성:
{{"mode": "rewrite_table", "table": [[...],[...],...]}}

2) 특정 열만 수정:
{{"mode": "update_column", "column": 0, "values": [...]}}

3) 특정 셀만 수정:
{{"mode": "update_cells", "cells": [{{"row":0,"col":2,"value":"..."}}, ...]}}

추가 설명, 코드블록, 자연어 문장을 절대 붙이지 마.
"""

    # 키/라이브러리가 없으면 기본값: 전체 재작성 (빈 패치)
    if client is None:
        return jsonify({"mode": "rewrite_table", "table": []})

    try:
        response = client.models.generate_content(
            model=DEFAULT_MODEL,
            contents=plan_prompt
        )
        text = response.text.strip()
        
        # JSON만 추출
        start = text.find("{")
        end = text.rfind("}")
        if start == -1 or end == -1 or end <= start:
            raise ValueError("no json in response")
        json_str = text[start : end + 1]
        patch = json.loads(json_str)
        return jsonify(patch)
    except Exception as e:
        print("[ERROR] plan_table LLM 오류 또는 파싱 실패:", e)
        # 실패 시 안전한 기본값
        return jsonify({"mode": "rewrite_table", "table": []})


if __name__ == "__main__":
    print(f"[INFO] AI rewrite 서버 시작: http://{HOST}:{PORT}  (model={DEFAULT_MODEL})")
    if not GEMINI_API_KEY:
        print("[WARN] GEMINI_API_KEY 미설정 - 입력 텍스트를 그대로 반환합니다.")
    app.run(host=HOST, port=PORT)
