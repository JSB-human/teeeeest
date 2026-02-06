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

app = Flask(__name__)

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
DEFAULT_MODEL = os.environ.get("INLINEAI_MODEL", "gemini-2.5-flash")
HOST = os.environ.get("INLINEAI_HOST", "127.0.0.1")
PORT = int(os.environ.get("INLINEAI_PORT", "5005"))

# google-generativeai 설정
try:
    import google.generativeai as genai

    if GEMINI_API_KEY:
        genai.configure(api_key=GEMINI_API_KEY)
    else:
        print("[WARN] GEMINI_API_KEY가 설정되지 않았습니다. 서버는 원문을 그대로 반환합니다.")
except Exception as e:  # ImportError 포함
    genai = None
    print(f"[WARN] google-generativeai를 불러올 수 없습니다: {e}. 서버는 원문을 그대로 반환합니다.")


Mode = Literal["rewrite", "summarize", "extend"]


def _build_instruction(mode: Mode) -> str:
    if mode == "summarize":
        return (
            "아래 한국어 글을 간결하게 요약해줘. "
            "반드시 요약된 본문만 출력하고, 설명이나 불릿, 추가 설명은 쓰지 마."
        )
    elif mode == "extend":
        return (
            "아래 한국어 글의 흐름을 유지하면서 내용을 자연스럽게 더 확장해줘. "
            "반드시 수정된 본문만 출력하고, 수정 이유나 설명은 쓰지 마."
        )
    else:
        return (
            "아래 한국어 글을 의미는 유지하면서 자연스럽고 매끄럽게 다듬어줘. "
            "반드시 다듬어진 본문만 출력하고, 수정 이유나 설명, 머리말/꼬리말은 쓰지 마."
        )


def gemini_rewrite(text: str, mode: Mode = "rewrite") -> str:
    """Gemini로 문장을 다듬거나 요약/확장.

    - 키/라이브러리가 없으면 원문 그대로 반환
    """
    if not text.strip():
        return text

    if not genai or not GEMINI_API_KEY:
        # 설정이 안 되어 있으면 원문 그대로
        return text

    instruction = _build_instruction(mode)
    prompt = f"{instruction}\n\n--- 원문 ---\n{text}\n\n--- 수정된 글 ---"

    model = genai.GenerativeModel(DEFAULT_MODEL)
    response = model.generate_content(prompt)

    # Gemini 응답에서 텍스트 추출
    try:
        rewritten = response.text or ""
    except Exception:
        # 라이브러리 버전에 따라 구조가 다를 수 있어 fallback
        rewritten = "".join(
            part.text
            for part in getattr(response, "candidates", [])
            if getattr(part, "text", None)
        )

    return rewritten.strip() or text


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

    if mode not in ("rewrite", "summarize", "extend"):
        mode = "rewrite"

    try:
        rewritten = gemini_rewrite(text, mode)  # type: ignore[arg-type]
    except Exception as e:
        print("[ERROR] Gemini 호출 중 오류:", e)
        rewritten = text

    return jsonify({"text": rewritten})


if __name__ == "__main__":
    print(f"[INFO] AI rewrite 서버 시작: http://{HOST}:{PORT}  (model={DEFAULT_MODEL})")
    if not GEMINI_API_KEY:
        print("[WARN] GEMINI_API_KEY 미설정 - 입력 텍스트를 그대로 반환합니다.")
    app.run(host=HOST, port=PORT)
