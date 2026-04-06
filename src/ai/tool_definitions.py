"""LLM 도구(함수) 정의 — ChatAgent가 LLM에 전달하는 도구 스키마.

이 정의는 프로바이더 중립 형식이며, LLMRouter가 Anthropic/OpenAI 형식으로 변환한다.
"""

from __future__ import annotations

from typing import Any

# 문서 수정을 수행하는 도구 (실행 전 스냅샷 저장 대상)
DOCUMENT_MODIFYING_TOOLS: set[str] = {
    "write_cell",
    "fill_field",
    "fill_all_empty_cells",
}

IMAGE_TOOLS: set[str] = {
    "generate_image",
}

HWP_TOOLS: list[dict[str, Any]] = [
    {
        "name": "analyze_document",
        "description": "현재 열린 HWP 문서의 표 구조를 분석하고 빈 셀을 감지합니다.",
        "parameters": {
            "type": "object",
            "properties": {},
            "required": [],
        },
    },
    {
        "name": "read_table",
        "description": "특정 표의 전체 내용을 읽어옵니다. 셀별 텍스트, 타입(라벨/빈셀/기입력), 서식 정보를 포함합니다.",
        "parameters": {
            "type": "object",
            "properties": {
                "table_idx": {
                    "type": "integer",
                    "description": "표 번호 (0부터 시작)",
                },
            },
            "required": ["table_idx"],
        },
    },
    {
        "name": "read_cell",
        "description": "특정 셀의 텍스트와 서식 정보를 읽어옵니다.",
        "parameters": {
            "type": "object",
            "properties": {
                "table_idx": {"type": "integer", "description": "표 번호"},
                "row": {"type": "integer", "description": "행 번호"},
                "col": {"type": "integer", "description": "열 번호"},
            },
            "required": ["table_idx", "row", "col"],
        },
    },
    {
        "name": "write_cell",
        "description": "특정 셀에 텍스트를 삽입합니다. 기존 서식(글꼴, 크기, 색상 등)이 유지됩니다.",
        "parameters": {
            "type": "object",
            "properties": {
                "table_idx": {"type": "integer", "description": "표 번호"},
                "row": {"type": "integer", "description": "행 번호"},
                "col": {"type": "integer", "description": "열 번호"},
                "text": {"type": "string", "description": "삽입할 텍스트"},
            },
            "required": ["table_idx", "row", "col", "text"],
        },
    },
    {
        "name": "fill_field",
        "description": "누름틀 필드에 텍스트를 삽입합니다.",
        "parameters": {
            "type": "object",
            "properties": {
                "field_name": {"type": "string", "description": "필드 이름"},
                "text": {"type": "string", "description": "삽입할 텍스트"},
            },
            "required": ["field_name", "text"],
        },
    },
    {
        "name": "fill_all_empty_cells",
        "description": "문서의 모든 빈 셀을 AI가 자동으로 채웁니다. 사업 정보를 제공하면 더 정확한 내용이 생성됩니다.",
        "parameters": {
            "type": "object",
            "properties": {
                "program_name": {
                    "type": "string",
                    "description": "정부과제명 (예: 예비창업패키지, TIPS, 데이터바우처)",
                },
                "company_name": {"type": "string", "description": "기업/기관명"},
                "company_desc": {"type": "string", "description": "기업 소개 (한 줄)"},
                "business_idea": {"type": "string", "description": "사업 아이디어 요약"},
            },
            "required": ["program_name", "company_name"],
        },
    },
    {
        "name": "validate_format",
        "description": "문서의 서식 규정(글꼴, 크기, 자간, 줄간격 등) 준수 여부를 검증합니다.",
        "parameters": {
            "type": "object",
            "properties": {
                "program_name": {
                    "type": "string",
                    "description": "정부과제명 (서식 규정이 과제별로 다름)",
                },
                "auto_fix": {
                    "type": "boolean",
                    "description": "위반 사항 자동 수정 여부",
                    "default": False,
                },
            },
            "required": ["program_name"],
        },
    },
    {
        "name": "undo",
        "description": "마지막 작업을 되돌립니다. 셀 쓰기, 자동 채우기 등의 편집을 취소할 수 있습니다.",
        "parameters": {
            "type": "object",
            "properties": {},
            "required": [],
        },
    },
    {
        "name": "save_document",
        "description": "현재 문서를 저장합니다. HWP, HWPX, PDF 형식을 지원합니다.",
        "parameters": {
            "type": "object",
            "properties": {
                "format": {
                    "type": "string",
                    "enum": ["hwp", "hwpx", "pdf"],
                    "description": "저장 형식",
                    "default": "hwp",
                },
            },
            "required": [],
        },
    },
    {
        "name": "get_document_info",
        "description": "현재 문서의 기본 정보(파일명, 표 개수, 편집 히스토리)를 반환합니다.",
        "parameters": {
            "type": "object",
            "properties": {},
            "required": [],
        },
    },
    {
        "name": "generate_image",
        "description": "AI로 이미지를 생성합니다. 사업계획서에 넣을 다이어그램, 로고, 컨셉 이미지 등을 만들 수 있습니다.",
        "parameters": {
            "type": "object",
            "properties": {
                "prompt": {
                    "type": "string",
                    "description": "생성할 이미지에 대한 상세 설명 (영어 권장)",
                },
                "size": {
                    "type": "string",
                    "enum": ["1024x1024", "1024x1792", "1792x1024"],
                    "description": "이미지 크기",
                    "default": "1024x1024",
                },
            },
            "required": ["prompt"],
        },
    },
]


def get_tools_for_provider(provider: str) -> list[dict[str, Any]]:
    """프로바이더에 맞는 도구 형식으로 변환한다."""
    if provider == "anthropic":
        return [
            {
                "name": t["name"],
                "description": t["description"],
                "input_schema": t["parameters"],
            }
            for t in HWP_TOOLS
        ]
    else:
        # OpenAI 형식
        return [
            {
                "type": "function",
                "function": {
                    "name": t["name"],
                    "description": t["description"],
                    "parameters": t["parameters"],
                },
            }
            for t in HWP_TOOLS
        ]
