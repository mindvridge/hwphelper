"""셀 생성기 테스트."""

from __future__ import annotations

from unittest.mock import AsyncMock, MagicMock

import pytest

from src.ai.cell_generator import CellGenerator
from src.ai.llm_router import LLMResponse, LLMRouter
from src.hwp_engine.cell_writer import CellFill


@pytest.fixture
def mock_router() -> LLMRouter:
    router = MagicMock(spec=LLMRouter)
    router.default_model = "test"
    return router


@pytest.fixture
def generator(mock_router) -> CellGenerator:
    return CellGenerator(llm_router=mock_router)


class TestGenerateSingleCell:
    """단일 셀 생성 테스트."""

    @pytest.mark.asyncio
    async def test_basic_generation(self, generator: CellGenerator, mock_router) -> None:
        mock_router.chat = AsyncMock(return_value=LLMResponse(content="AI 기반 문서 자동화", model="test"))

        result = await generator.generate_single_cell(
            cell_schema={"row": 0, "col": 1, "context": {"row_label": "사업명"}},
            program_name="예비창업패키지",
        )

        assert result == "AI 기반 문서 자동화"
        mock_router.chat.assert_called_once()

    @pytest.mark.asyncio
    async def test_strips_quotes(self, generator: CellGenerator, mock_router) -> None:
        mock_router.chat = AsyncMock(return_value=LLMResponse(content='"AI 문서 자동화"', model="test"))

        result = await generator.generate_single_cell(
            cell_schema={"row": 0, "col": 1, "context": {}},
        )

        assert result == "AI 문서 자동화"


class TestGenerateBatch:
    """배치 생성 테스트."""

    @pytest.mark.asyncio
    async def test_batch_parse(self, generator: CellGenerator, mock_router) -> None:
        json_resp = '{"cells": [{"row": 0, "col": 1, "content": "값1"}, {"row": 1, "col": 1, "content": "값2"}]}'
        mock_router.chat = AsyncMock(return_value=LLMResponse(content=json_resp, model="test"))

        fills = await generator.generate_batch(
            cells=[
                {"row": 0, "col": 1, "context": {}},
                {"row": 1, "col": 1, "context": {}},
            ],
        )

        assert len(fills) == 2
        assert fills[0].text == "값1"
        assert fills[1].row == 1

    @pytest.mark.asyncio
    async def test_batch_with_code_block(self, generator: CellGenerator, mock_router) -> None:
        resp = '```json\n{"cells": [{"row": 0, "col": 0, "content": "테스트"}]}\n```'
        mock_router.chat = AsyncMock(return_value=LLMResponse(content=resp, model="test"))

        fills = await generator.generate_batch(cells=[{"row": 0, "col": 0, "context": {}}])

        assert len(fills) == 1
        assert fills[0].text == "테스트"


class TestParseBatchResponse:
    """배치 응답 파싱 테스트."""

    def test_valid_json(self) -> None:
        content = '{"cells": [{"row": 0, "col": 1, "content": "hello"}]}'
        fills = CellGenerator._parse_batch_response(content)
        assert len(fills) == 1
        assert isinstance(fills[0], CellFill)

    def test_invalid_json(self) -> None:
        fills = CellGenerator._parse_batch_response("not json at all")
        assert fills == []

    def test_json_in_code_block(self) -> None:
        content = '```json\n{"cells": [{"row": 1, "col": 2, "content": "val"}]}\n```'
        fills = CellGenerator._parse_batch_response(content)
        assert len(fills) == 1
        assert fills[0].row == 1


class TestGenerateAll:
    """전체 셀 생성 테스트."""

    @pytest.mark.asyncio
    async def test_generate_all(self, generator: CellGenerator, mock_router) -> None:
        mock_router.chat = AsyncMock(return_value=LLMResponse(content="생성된 내용", model="test"))

        schema = {
            "tables": [
                {
                    "table_idx": 0,
                    "cells": [
                        {"row": 0, "col": 0, "needs_fill": False, "context": {}},
                        {"row": 0, "col": 1, "needs_fill": True, "context": {"row_label": "사업명"}},
                        {"row": 1, "col": 1, "needs_fill": True, "context": {"row_label": "기관명"}},
                    ],
                }
            ]
        }

        fills = await generator.generate_all(schema, concurrency=1)

        assert len(fills) == 2
        assert all(f.text == "생성된 내용" for f in fills)

    @pytest.mark.asyncio
    async def test_generate_all_empty(self, generator: CellGenerator) -> None:
        fills = await generator.generate_all({"tables": []})
        assert fills == []

    @pytest.mark.asyncio
    async def test_progress_callback(self, generator: CellGenerator, mock_router) -> None:
        mock_router.chat = AsyncMock(return_value=LLMResponse(content="OK", model="test"))

        progress_calls = []

        async def on_progress(current, total, cell):
            progress_calls.append((current, total))

        schema = {
            "tables": [{
                "table_idx": 0,
                "cells": [{"row": 0, "col": 0, "needs_fill": True, "context": {}}],
            }]
        }

        await generator.generate_all(schema, on_progress=on_progress, concurrency=1)

        assert len(progress_calls) == 1
        assert progress_calls[0] == (1, 1)
