# HwpInlineAI 승인/거절 + Diff 아키텍처 설계서

작성일: 2026-02-12
목표: 대화형 문서 작성 + 안전한 승인 워크플로우 + 텍스트/표 Diff를 일관된 구조로 제공

---

## 1) 제품 목표 (Product Goal)

- 사용자는 자연어로 문서 작성을 지시한다.
- AI는 바로 반영하지 않고 항상 **변경안(Proposal)** 을 만든다.
- 사용자는 변경안을 **승인/거절** 한다.
- 변경안은 텍스트/표 모두 **Diff 시각화** 된다.
- 모든 변경은 추적 가능해야 한다(감사 로그).

---

## 2) 현재 코드 기준 진단

현 상태에서 이미 있는 기반:
- `ai/rewrite_server.py` : `/rewrite`, `/plan_table` 제공
- `src/tools/engine.py` : 표 preview/apply/cancel, 선택영역 치환 로직 존재
- `ui_app.py` : 승인/거절 버튼 및 기본 흐름 존재

문제점/보완 필요:
1. 텍스트/표 변경 상태 저장 구조가 분산되어 있음
2. preview와 apply가 공통 타입으로 관리되지 않음
3. reject 시 복구 전략이 Undo 의존적이라 케이스별 불안정 가능
4. 변경 이력(누가/언제/무엇을 승인했는지) 영속 저장 없음
5. 모델(Gemini) 종속성이 강함

---

## 3) 제안 아키텍처

### 3.1 핵심 개념: ChangeSet 단일화

모든 변경안을 아래 공통 스키마로 관리:

```json
{
  "id": "uuid",
  "kind": "text|table|document",
  "scope": {
    "document_path": "...",
    "selection_meta": {"para_id": 0, "char_pos": 0},
    "table_anchor": {"row": 1, "col": 1}
  },
  "prompt": "사용자 요청",
  "model": "gpt-4.1|gemini-2.5-flash|...",
  "before": "원본 또는 원본 구조",
  "after": "수정안 또는 수정 구조",
  "diff": {
    "text_spans": [],
    "table_cells": []
  },
  "status": "draft|previewed|approved|rejected|applied|failed",
  "created_at": "ISO",
  "updated_at": "ISO"
}
```

### 3.2 레이어 분리

1) **LLM Adapter Layer**
- 역할: 모델별 API 호출 표준화
- 파일 제안: `src/ai/adapters/{base.py, gemini.py, openai.py, anthropic.py}`

2) **Proposal Service Layer**
- 역할: 프롬프트+원본 -> ChangeSet 생성
- 파일 제안: `src/services/proposal_service.py`

3) **Diff Service Layer**
- 역할: 텍스트/표 diff 계산
- 파일 제안: `src/services/diff_service.py`

4) **Apply Service Layer**
- 역할: 승인된 ChangeSet을 HWP에 반영
- 파일 제안: `src/services/apply_service.py`

5) **Session Store Layer**
- 역할: 변경안 상태 저장 + 복구용 스냅샷 보관
- 파일 제안: `src/state/session_store.py`

6) **Audit Log Layer**
- 역할: 승인/거절/실패 이벤트 기록
- 파일 제안: `src/state/audit_log.py`

---

## 4) 상태 머신 (필수)

```text
draft -> previewed -> approved -> applied
                    \-> rejected
previewed -> failed (렌더/계산 오류)
approved -> failed (적용 실패)
```

규칙:
- `applied` 이전에는 문서 영구 반영 금지
- `rejected`는 항상 문서 원상 복구 보장
- 모든 상태 전이는 감사 로그에 기록

---

## 5) Diff 전략

### 5.1 텍스트 Diff
- Python `difflib.SequenceMatcher` 기반 단어/문장 단위 span 생성
- UI 표시: 삭제(빨강), 추가(초록)
- 적용 방식:
  - preview: 임시 시각화만
  - approve: 최종 텍스트 clean replace

### 5.2 표 Diff
- 셀 단위 `{row,col,old,new}`
- UI 패널에 변경 셀 목록 제공
- 적용 방식:
  - preview: 문서 무변경 + 패널 표시 우선 (권장)
  - approve: 변경 셀만 patch 적용

---

## 6) 파일 구조 제안

```text
F:/dev/hwp-mcp/
  ai/
    rewrite_server.py
  src/
    ai/
      adapters/
        base.py
        gemini.py
        openai.py
        anthropic.py
    services/
      proposal_service.py
      diff_service.py
      apply_service.py
    state/
      session_store.py
      audit_log.py
    tools/
      engine.py            # orchestration만 남기고 서비스 호출
      hwp_controller.py
      hwp_table_tools.py
  ui_app.py
```

---

## 7) 구현 순서 (실행 플랜)

### Phase 1 (빠른 안정화)
1. `ChangeSet` 데이터 클래스 도입
2. `engine.py`의 text/table preview 결과를 ChangeSet으로 저장
3. approve/reject 공통 처리 함수 통합

### Phase 2 (Diff 고도화)
4. `diff_service.py` 구현 (text/table)
5. UI 우측 패널에 diff 요약 목록 추가
6. reject 복구를 Undo 의존에서 snapshot 기반으로 개선

### Phase 3 (모델 확장)
7. `LLM Adapter` 추상화
8. `rewrite_server.py`를 provider 선택형으로 확장
9. env 기반 모델 라우팅

### Phase 4 (운영성)
10. `audit_log.jsonl` 기록
11. 에러 리커버리/재시도
12. 통합 테스트 (텍스트/표/승인/거절)

---

## 8) 즉시 적용 가능한 최소 변경안 (MVP)

- `engine.py`에 아래 함수 추가:
  - `create_changeset_for_selection(prompt)`
  - `preview_changeset(changeset_id)`
  - `approve_changeset(changeset_id)`
  - `reject_changeset(changeset_id)`
- `_last_table_patch`, `_last_ai_result` 같은 흩어진 상태를 `SessionStore` 하나로 통합
- `ui_app.py` 버튼은 changeset_id 기준으로 동작하도록 변경

---

## 9) 권장 운영 규칙

- 기본값은 항상 `preview_required = true`
- 문서 전체 재작성은 confirm 이중 체크
- 표 변경은 “변경 셀 개수”가 임계치 이상이면 추가 확인
- 모든 apply는 파일 백업(예: `.bak`) 후 실행

---

## 10) 다음 작업 제안 (바로 코딩 가능)

1. `src/state/session_store.py` 생성
2. `src/services/diff_service.py` 생성
3. `engine.py`를 ChangeSet 중심으로 리팩터
4. `ui_app.py` 승인/거절 이벤트를 changeset_id 기반으로 변경

---

원하면 다음 단계로 위 1~4를 실제 코드로 바로 적용하겠습니다.
