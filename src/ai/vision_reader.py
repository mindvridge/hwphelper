"""LLM 비전으로 HWP 문서 페이지를 분석하여 표 구조를 인식한다.

페이지 이미지를 비전 가능 LLM에 전송하여:
1. 표 위치/구조/셀 내용 추출
2. 셀 유형 분류 (라벨/빈칸/예시/안내문)
3. 시각적 서식 정보 (색상, 굵기 등)

사용법::

    reader = VisionReader(llm_router)
    result = await reader.read_page_tables(page_image)
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any

import structlog

if TYPE_CHECKING:
    from .llm_router import LLMRouter
    from ..hwp_engine.page_renderer import PageImage

logger = structlog.get_logger()


@dataclass
class VisionCell:
    """비전으로 인식된 셀."""

    row: int
    col: int
    row_span: int = 1
    col_span: int = 1
    text: str = ""
    is_label: bool = False
    is_empty: bool = False
    is_guide: bool = False  # ※ 안내문
    color: str = "black"    # "black", "blue", "red" 등


@dataclass
class VisionTable:
    """비전으로 인식된 표."""

    table_idx: int
    rows: int
    cols: int
    cells: list[VisionCell] = field(default_factory=list)
    description: str = ""   # 표 전체 설명


@dataclass
class VisionPageResult:
    """한 페이지의 비전 분석 결과."""

    page_num: int
    tables: list[VisionTable] = field(default_factory=list)
    page_description: str = ""
    raw_response: str = ""


@dataclass
class VerificationResult:
    """셀 쓰기 검증 결과."""

    table_idx: int
    row: int
    col: int
    expected_text: str
    actual_text: str = ""
    match: bool = False
    confidence: float = 0.0
    issue: str = ""  # "ok", "wrong_cell", "truncated", "empty", "formatting_lost"


_SYSTEM_PROMPT = """당신은 한국어 정부과제 사업계획서(HWP 문서)를 분석하는 전문가입니다.
주어진 문서 페이지 이미지를 분석하여 표 구조와 셀 내용을 정확하게 추출하세요.

출력 형식 (JSON):
{
  "page_description": "페이지 전체 설명",
  "tables": [
    {
      "table_idx": 0,
      "description": "표 설명",
      "rows": 행수,
      "cols": 열수,
      "cells": [
        {
          "row": 0, "col": 0,
          "row_span": 1, "col_span": 1,
          "text": "셀 텍스트",
          "is_label": true/false,
          "is_empty": true/false,
          "is_guide": true/false,
          "color": "black"
        }
      ]
    }
  ]
}

규칙:
- is_label: 항목명(라벨) 셀이면 true
- is_empty: 텍스트가 없거나 빈칸이면 true
- is_guide: ※로 시작하는 안내문이면 true
- color: 텍스트 색상 (black, blue, red, gray 등)
- 병합된 셀은 row_span, col_span으로 표시
- 한국어 텍스트를 정확하게 인식하세요
"""

_VERIFY_PROMPT = """이 문서 페이지 이미지를 보고, 지정된 표의 특정 셀에 올바른 내용이 작성되었는지 확인하세요.

확인할 내용:
{checks}

출력 형식 (JSON):
{{
  "results": [
    {{
      "row": 0, "col": 0,
      "expected": "기대 텍스트",
      "actual": "실제로 보이는 텍스트",
      "match": true/false,
      "confidence": 0.95,
      "issue": "ok" 또는 "wrong_cell"/"truncated"/"empty"/"formatting_lost"
    }}
  ]
}}
"""


class VisionReader:
    """LLM 비전으로 문서 페이지를 분석한다."""

    def __init__(self, llm_router: LLMRouter, model_id: str | None = None) -> None:
        self._router = llm_router
        self._model_id = model_id or self._find_vision_model()

    def _find_vision_model(self) -> str:
        """비전 지원 모델을 자동 탐색한다."""
        for mid, cfg in self._router._models.items():
            if cfg.get("supports_vision"):
                return mid
        return self._router._default_model

    async def read_page_tables(self, page: PageImage) -> VisionPageResult:
        """페이지 이미지에서 모든 표를 인식한다."""
        messages = [
            {"role": "user", "content": [
                {"type": "image_url", "image_url": {
                    "url": f"data:{page.mime_type};base64,{page.base64_data}",
                }},
                {"type": "text", "text": _SYSTEM_PROMPT + "\n\n이 페이지의 모든 표를 분석하여 JSON으로 출력하세요."},
            ]},
        ]

        try:
            # 독립 httpx 클라이언트로 직접 호출 (공유 클라이언트 데드락 방지)
            response_text = await self._direct_call(messages)
            return self._parse_page_result(page.page_num, response_text)
        except Exception as exc:
            logger.warning("비전 페이지 분석 실패", page=page.page_num, error=str(exc))
            return VisionPageResult(page_num=page.page_num)

    async def _direct_call(self, messages: list[dict]) -> str:
        """동기 httpx를 별도 스레드에서 실행하여 이벤트 루프 블로킹을 방지한다."""
        import asyncio
        import concurrent.futures

        loop = asyncio.get_running_loop()
        with concurrent.futures.ThreadPoolExecutor(max_workers=1) as pool:
            return await loop.run_in_executor(pool, self._sync_call, messages)

    def _sync_call(self, messages: list[dict]) -> str:
        """동기 httpx로 LLM API를 호출한다."""
        import httpx
        import os

        cfg = self._router._models.get(self._model_id, {})
        provider = cfg.get("provider", "")
        model_name = cfg.get("model", "")
        api_key = os.environ.get(cfg.get("api_key_env", ""), "") or cfg.get("default_api_key", "")
        base_url = os.environ.get(cfg.get("base_url_env", ""), "") or cfg.get("default_base_url", "")

        if provider == "anthropic":
            converted_msgs = self._router._convert_content_for_anthropic(messages[0]["content"])
            body = {
                "model": model_name,
                "messages": [{"role": "user", "content": converted_msgs}],
                "max_tokens": 4096,
            }
            headers = {
                "Content-Type": "application/json",
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
            }
            url = "https://api.anthropic.com/v1/messages"
        else:
            body = {
                "model": model_name,
                "messages": messages,
                "max_tokens": 4096,
            }
            headers = {"Content-Type": "application/json"}
            if api_key:
                headers["Authorization"] = f"Bearer {api_key}"
            url = f"{base_url.rstrip('/')}/chat/completions"

        with httpx.Client(timeout=120) as client:
            resp = client.post(url, json=body, headers=headers)
            resp.raise_for_status()
            data = resp.json()

        if provider == "anthropic":
            return data.get("content", [{}])[0].get("text", "")
        else:
            return data["choices"][0]["message"]["content"]

    async def verify_writes(
        self,
        page: PageImage,
        checks: list[dict[str, Any]],
    ) -> list[VerificationResult]:
        """페이지 이미지에서 셀 쓰기 결과를 검증한다.

        checks: [{"table_idx": 0, "row": 0, "col": 3, "expected": "텍스트"}, ...]
        """
        checks_text = json.dumps(checks, ensure_ascii=False, indent=2)
        prompt = _VERIFY_PROMPT.format(checks=checks_text)

        messages = [
            {"role": "system", "content": "당신은 문서 검증 전문가입니다. 이미지와 기대값을 비교하세요."},
            {"role": "user", "content": [
                {"type": "image", "source": {
                    "type": "base64",
                    "media_type": page.mime_type,
                    "data": page.base64_data,
                }},
                {"type": "text", "text": prompt},
            ]},
        ]

        try:
            from .llm_router import LLMResponse
            response = await self._router.chat(
                messages=messages,
                model_id=self._model_id,
                max_tokens=4096,
            )
            if not isinstance(response, LLMResponse):
                return []

            return self._parse_verify_result(response.content, checks)
        except Exception as exc:
            logger.warning("비전 검증 실패", error=str(exc))
            return []

    def _sync_read_page(self, page: PageImage) -> VisionPageResult:
        """동기 방식으로 페이지를 분석한다 (스레드에서 호출용)."""
        # 간소화된 프롬프트로 응답 시간 단축
        compact_prompt = (
            "이 한국어 문서 페이지의 표를 분석하세요. JSON으로 출력:\n"
            '{"tables":[{"table_idx":0,"rows":행수,"cols":열수,"description":"표설명",'
            '"cells":[{"row":0,"col":0,"text":"셀텍스트","is_label":true,"is_empty":false,'
            '"is_guide":false,"color":"black"}]}]}\n'
            "규칙: is_label=항목명, is_guide=※안내문, color=텍스트색상(black/blue). "
            "셀 텍스트는 처음 30자만."
        )
        messages = [
            {"role": "user", "content": [
                {"type": "image_url", "image_url": {
                    "url": f"data:{page.mime_type};base64,{page.base64_data}",
                }},
                {"type": "text", "text": compact_prompt},
            ]},
        ]
        try:
            response_text = self._sync_call(messages)
            return self._parse_page_result(page.page_num, response_text)
        except Exception as exc:
            logger.warning("비전 동기 분석 실패", error=str(exc))
            return VisionPageResult(page_num=page.page_num)

    async def read_all_pages(self, pages: list[PageImage]) -> list[VisionPageResult]:
        """모든 페이지를 순차적으로 분석한다."""
        results: list[VisionPageResult] = []
        for page in pages:
            result = await self.read_page_tables(page)
            results.append(result)
        return results

    def _parse_page_result(self, page_num: int, raw: str) -> VisionPageResult:
        """LLM 응답에서 JSON을 파싱한다."""
        try:
            # JSON 블록 추출
            json_str = self._extract_json(raw)
            data = json.loads(json_str)

            tables: list[VisionTable] = []
            for t in data.get("tables", []):
                cells = [
                    VisionCell(
                        row=c.get("row", 0),
                        col=c.get("col", 0),
                        row_span=c.get("row_span", 1),
                        col_span=c.get("col_span", 1),
                        text=c.get("text", ""),
                        is_label=c.get("is_label", False),
                        is_empty=c.get("is_empty", False),
                        is_guide=c.get("is_guide", False),
                        color=c.get("color", "black"),
                    )
                    for c in t.get("cells", [])
                ]
                tables.append(VisionTable(
                    table_idx=t.get("table_idx", 0),
                    rows=t.get("rows", 0),
                    cols=t.get("cols", 0),
                    cells=cells,
                    description=t.get("description", ""),
                ))

            return VisionPageResult(
                page_num=page_num,
                tables=tables,
                page_description=data.get("page_description", ""),
                raw_response=raw,
            )
        except Exception as exc:
            logger.debug("비전 응답 파싱 실패", error=str(exc))
            return VisionPageResult(page_num=page_num, raw_response=raw)

    def _parse_verify_result(
        self, raw: str, checks: list[dict]
    ) -> list[VerificationResult]:
        """검증 응답을 파싱한다."""
        try:
            json_str = self._extract_json(raw)
            data = json.loads(json_str)

            results: list[VerificationResult] = []
            for r in data.get("results", []):
                check = next(
                    (c for c in checks if c["row"] == r.get("row") and c["col"] == r.get("col")),
                    {},
                )
                results.append(VerificationResult(
                    table_idx=check.get("table_idx", 0),
                    row=r.get("row", 0),
                    col=r.get("col", 0),
                    expected_text=r.get("expected", ""),
                    actual_text=r.get("actual", ""),
                    match=r.get("match", False),
                    confidence=r.get("confidence", 0.0),
                    issue=r.get("issue", "unknown"),
                ))
            return results
        except Exception:
            return []

    @staticmethod
    def _extract_json(text: str) -> str:
        """텍스트에서 JSON 블록을 추출한다."""
        # ```json ... ``` 블록
        if "```json" in text:
            start = text.index("```json") + 7
            end = text.index("```", start)
            return text[start:end].strip()
        if "```" in text:
            start = text.index("```") + 3
            end = text.index("```", start)
            return text[start:end].strip()
        # 중괄호로 시작하는 JSON
        for i, ch in enumerate(text):
            if ch == "{":
                depth = 0
                for j in range(i, len(text)):
                    if text[j] == "{":
                        depth += 1
                    elif text[j] == "}":
                        depth -= 1
                    if depth == 0:
                        return text[i : j + 1]
        return text
