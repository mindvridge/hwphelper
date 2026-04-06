"""프롬프트 빌더 — 셀 콘텐츠 생성용 프롬프트를 구성.

LLM에게 정부과제 계획서에 적합한 콘텐츠를 생성하도록
표 구조, 셀 컨텍스트, 사업 정보를 포함한 프롬프트를 구성한다.
"""

from __future__ import annotations

import json
from typing import Any


CELL_SYSTEM_PROMPT = """당신은 한국 정부지원사업 계획서 작성 전문 컨설턴트입니다.
HWP 문서의 표 셀에 들어갈 전문적인 내용을 작성합니다.

## 작성 규칙
1. 라벨(항목명)을 정확히 파악하여 해당 셀에 적합한 내용을 작성합니다
2. 구체적 수치, 일정, 기술명을 포함하여 평가위원이 신뢰할 수 있게 작성합니다
3. "~한다", "~이다" 체로 간결하게 작성합니다 (의지형 "~하겠습니다" 금지)
4. 셀 크기에 맞는 분량: 제목(1줄), 내용(3~10줄), 개요(5~15줄)
5. 금액은 "00,000천원" 또는 "0억 0,000만원" 형식
6. 날짜는 "2026.01 ~ 2026.12" 형식
7. 마크다운(**, ##, -), HTML 태그, 따옴표를 절대 사용하지 않습니다
8. 셀에 들어갈 내용만 반환합니다. 부연 설명, 옵션 제시를 하지 않습니다

## 항목별 분량 가이드
- 사업명/과제명: 20자 내외 한 줄
- 대표자/담당자: 이름만
- 연락처/이메일: 형식에 맞게
- 사업 개요/요약: 3~5줄
- 기술 설명/핵심기술: 5~10줄
- 시장 분석: 5~8줄 (TAM/SAM/SOM 포함)
- 사업화 전략: 5~8줄
- 개발 일정: 단계별 월 단위
- 예산/비용: 숫자 + 산출근거"""


class PromptBuilder:
    """셀 콘텐츠 생성용 프롬프트 빌더."""

    def __init__(self, system_prompt: str = CELL_SYSTEM_PROMPT) -> None:
        self._system_prompt = system_prompt

    @property
    def system_prompt(self) -> str:
        return self._system_prompt

    def build_cell_prompt(
        self,
        cell_schema: dict[str, Any],
        program_name: str = "",
        company_info: str = "",
        rag_context: str = "",
    ) -> str:
        """단일 셀 채우기용 프롬프트를 생성한다."""
        parts: list[str] = []

        # 맥락 정보
        if program_name:
            parts.append(f"## 정부과제: {program_name}")
        if company_info:
            parts.append(f"## 기업/기관 정보\n{company_info}")

        # 셀 위치와 컨텍스트
        row, col = cell_schema.get("row", 0), cell_schema.get("col", 0)
        parts.append(f"## 대상 셀: ({row}행, {col}열)")

        context = cell_schema.get("context", {})
        if context:
            ctx_lines: list[str] = []
            if "row_label" in context:
                ctx_lines.append(f"- 행 라벨: {context['row_label']}")
            if "col_header" in context:
                ctx_lines.append(f"- 열 헤더: {context['col_header']}")
            if "table_header" in context:
                ctx_lines.append(f"- 표 헤더: {context['table_header']}")
            if "same_row_content" in context:
                ctx_lines.append(f"- 같은 행 내용: {context['same_row_content']}")
            if ctx_lines:
                parts.append("## 셀 주변 정보\n" + "\n".join(ctx_lines))

        # RAG 컨텍스트
        if rag_context:
            parts.append(f"## 참고 문서\n{rag_context}")

        parts.append("위 정보를 바탕으로 이 셀에 들어갈 내용을 작성해주세요:")

        return "\n\n".join(parts)

    def build_batch_prompt(
        self,
        cells: list[dict[str, Any]],
        program_name: str = "",
        company_info: str = "",
    ) -> str:
        """여러 셀을 한 번에 채우기 위한 프롬프트를 생성한다."""
        parts: list[str] = []

        if program_name:
            parts.append(f"## 정부과제: {program_name}")
        if company_info:
            parts.append(f"## 기업/기관 정보\n{company_info}")

        parts.append(f"## 채워야 할 셀 ({len(cells)}개)")
        for cell in cells:
            row, col = cell.get("row", 0), cell.get("col", 0)
            context = cell.get("context", {})
            label = context.get("row_label", "") or context.get("col_header", "")
            parts.append(f"- ({row},{col}): {label or '(컨텍스트 없음)'}")

        parts.append(
            '\n다음 JSON 형식으로 응답해주세요:\n'
            '{"cells": [{"row": 0, "col": 1, "content": "내용"}, ...]}'
        )

        return "\n\n".join(parts)
