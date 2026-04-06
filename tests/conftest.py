"""공통 테스트 픽스처.

단위 테스트용 데이터 픽스처와 통합 테스트용 COM/LLM 픽스처를 제공한다.
COM 관련 픽스처는 한/글 미설치 시 자동으로 skip된다.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from src.hwp_engine.com_controller import HAS_HWP
from src.hwp_engine.table_reader import Cell, CellStyle, CellType, Table


# ------------------------------------------------------------------
# 스타일 픽스처
# ------------------------------------------------------------------


@pytest.fixture
def sample_style() -> CellStyle:
    """기본 셀 스타일."""
    return CellStyle(
        font_name="맑은 고딕",
        font_size=10.0,
        bold=False,
        italic=False,
        char_spacing=0.0,
        line_spacing=160.0,
        alignment="left",
        text_color="0x00000000",
    )


@pytest.fixture
def bold_style() -> CellStyle:
    """볼드 셀 스타일."""
    return CellStyle(
        font_name="맑은 고딕",
        font_size=10.0,
        bold=True,
        italic=False,
        char_spacing=0.0,
        line_spacing=160.0,
        alignment="center",
        text_color="0x00000000",
    )


# ------------------------------------------------------------------
# 표 데이터 픽스처
# ------------------------------------------------------------------


@pytest.fixture
def sample_table(sample_style: CellStyle, bold_style: CellStyle) -> Table:
    """2행 3열 샘플 표.

    | 사업명 | (빈 셀)        | 비고     |
    | 기관명 | (내용 입력)     | 연락처   |
    """
    return Table(
        table_idx=0,
        rows=2,
        cols=3,
        cells=[
            Cell(row=0, col=0, text="사업명", cell_type=CellType.UNKNOWN, style=bold_style),
            Cell(row=0, col=1, text="", cell_type=CellType.UNKNOWN, style=sample_style),
            Cell(row=0, col=2, text="비고", cell_type=CellType.UNKNOWN, style=bold_style),
            Cell(row=1, col=0, text="기관명", cell_type=CellType.UNKNOWN, style=bold_style),
            Cell(row=1, col=1, text="(내용 입력)", cell_type=CellType.UNKNOWN, style=sample_style),
            Cell(row=1, col=2, text="연락처", cell_type=CellType.UNKNOWN, style=bold_style),
        ],
    )


@pytest.fixture
def classified_table(sample_table: Table) -> Table:
    """분류 완료된 샘플 표."""
    from src.hwp_engine.cell_classifier import CellClassifier

    classifier = CellClassifier()
    return classifier.classify_table(sample_table)


@pytest.fixture
def large_table(sample_style: CellStyle, bold_style: CellStyle) -> Table:
    """4행 4열 큰 표.

    | 구분   | 항목     | 내용   | 비고     |
    | 1     | 목표     |        | 참고     |
    | 2     | 전략     |        |          |
    | 합계   |         |        |          |
    """
    return Table(
        table_idx=1,
        rows=4,
        cols=4,
        cells=[
            Cell(row=0, col=0, text="구분", style=bold_style),
            Cell(row=0, col=1, text="항목", style=bold_style),
            Cell(row=0, col=2, text="내용", style=bold_style),
            Cell(row=0, col=3, text="비고", style=bold_style),
            Cell(row=1, col=0, text="1", style=sample_style),
            Cell(row=1, col=1, text="목표", style=sample_style),
            Cell(row=1, col=2, text="", style=sample_style),
            Cell(row=1, col=3, text="참고", style=sample_style),
            Cell(row=2, col=0, text="2", style=sample_style),
            Cell(row=2, col=1, text="전략", style=sample_style),
            Cell(row=2, col=2, text="", style=sample_style),
            Cell(row=2, col=3, text="", style=sample_style),
            Cell(row=3, col=0, text="합계", style=bold_style),
            Cell(row=3, col=1, text="", style=sample_style),
            Cell(row=3, col=2, text="", style=sample_style),
            Cell(row=3, col=3, text="", style=sample_style),
        ],
    )


# ------------------------------------------------------------------
# COM 픽스처 (한/글 설치 필요)
# ------------------------------------------------------------------


@pytest.fixture
def hwp_controller():
    """한/글 COM 연결. 미설치 시 스킵."""
    if not HAS_HWP:
        pytest.skip("한/글이 설치되어 있지 않습니다.")

    from src.hwp_engine.com_controller import HwpController

    ctrl = HwpController(visible=False)
    ctrl.connect()
    yield ctrl
    ctrl.quit()


@pytest.fixture
def sample_template(tmp_path: Path, hwp_controller):
    """테스트용 HWP 템플릿 (2x3 표) 생성.

    한/글 COM으로 빈 문서에 표를 생성한다.
    """
    hwp = hwp_controller.hwp
    filepath = str(tmp_path / "test_template.hwp")

    # 새 문서
    hwp.HAction.Run("FileNew")

    # 2행 3열 표 삽입
    hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
    hwp.HParameterSet.HTableCreation.Rows = 2
    hwp.HParameterSet.HTableCreation.Cols = 3
    hwp.HParameterSet.HTableCreation.WidthType = 2  # 페이지 너비
    hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)

    # 라벨 채우기
    labels = [("사업명", 0, 0), ("", 0, 1), ("비고", 0, 2),
              ("기관명", 1, 0), ("", 1, 1), ("연락처", 1, 2)]
    for text, row, col in labels:
        if text:
            try:
                hwp.ShapeObjTableSelCell(0, row, col)
                hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                hwp.HParameterSet.HInsertText.Text = text
                hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
            except Exception:
                pass

    hwp_controller.save_as(filepath)
    return filepath


# ------------------------------------------------------------------
# LLM 픽스처 (Mock)
# ------------------------------------------------------------------


@pytest.fixture
def mock_llm_router():
    """Mock LLM 라우터 — 고정 응답을 반환한다."""
    from src.ai.llm_router import LLMResponse, LLMRouter

    router = MagicMock(spec=LLMRouter)
    router.default_model = "test-model"
    router._models = {"test-model": {"provider": "openai", "model": "test"}}
    router.chat = AsyncMock(return_value=LLMResponse(
        content="테스트 AI 응답입니다.",
        tool_calls=[],
        model="test-model",
    ))
    router.list_models.return_value = []
    return router


@pytest.fixture
def mock_doc_manager(tmp_path: Path):
    """Mock DocumentManager."""
    from src.hwp_engine.document_manager import DocumentManager

    mgr = MagicMock(spec=DocumentManager)
    mgr.active_sessions = []
    return mgr


@pytest.fixture
def mock_chat_agent(mock_llm_router, mock_doc_manager):
    """테스트용 ChatAgent (mock LLM)."""
    from src.ai.chat_agent import ChatAgent

    agent = ChatAgent(llm_router=mock_llm_router, doc_manager=mock_doc_manager)
    return agent


# ------------------------------------------------------------------
# 문서 스키마 픽스처
# ------------------------------------------------------------------


@pytest.fixture
def sample_schema() -> dict[str, Any]:
    """분석 완료된 문서 스키마."""
    return {
        "document_name": "test_template.hwp",
        "total_tables": 2,
        "total_cells_to_fill": 5,
        "tables": [
            {
                "table_idx": 0,
                "rows": 2,
                "cols": 3,
                "cells_to_fill": 2,
                "cells": [
                    {"row": 0, "col": 0, "text": "사업명", "cell_type": "label", "needs_fill": False},
                    {"row": 0, "col": 1, "text": "", "cell_type": "empty", "needs_fill": True,
                     "context": {"row_label": "사업명"}},
                    {"row": 0, "col": 2, "text": "비고", "cell_type": "label", "needs_fill": False},
                    {"row": 1, "col": 0, "text": "기관명", "cell_type": "label", "needs_fill": False},
                    {"row": 1, "col": 1, "text": "(내용 입력)", "cell_type": "placeholder", "needs_fill": True,
                     "context": {"row_label": "기관명"}},
                    {"row": 1, "col": 2, "text": "연락처", "cell_type": "label", "needs_fill": False},
                ],
            },
            {
                "table_idx": 1,
                "rows": 4,
                "cols": 4,
                "cells_to_fill": 3,
                "cells": [
                    {"row": 0, "col": 0, "text": "구분", "cell_type": "label", "needs_fill": False},
                    {"row": 0, "col": 1, "text": "항목", "cell_type": "label", "needs_fill": False},
                    {"row": 0, "col": 2, "text": "내용", "cell_type": "label", "needs_fill": False},
                    {"row": 0, "col": 3, "text": "비고", "cell_type": "label", "needs_fill": False},
                    {"row": 1, "col": 2, "text": "", "cell_type": "empty", "needs_fill": True,
                     "context": {"row_label": "목표", "col_header": "내용"}},
                    {"row": 2, "col": 2, "text": "", "cell_type": "empty", "needs_fill": True,
                     "context": {"row_label": "전략", "col_header": "내용"}},
                    {"row": 2, "col": 3, "text": "", "cell_type": "empty", "needs_fill": True,
                     "context": {"col_header": "비고"}},
                ],
            },
        ],
    }


# ------------------------------------------------------------------
# FastAPI 테스트 앱 픽스처
# ------------------------------------------------------------------


@pytest.fixture
def test_app():
    """테스트용 FastAPI 앱."""
    from src.server import create_app

    app = create_app()
    app.state.llm_router = MagicMock()
    app.state.llm_router.default_model = "test-model"
    app.state.llm_router.list_models.return_value = []
    app.state.doc_manager = MagicMock()
    app.state.doc_manager.active_sessions = []
    app.state.chat_agent = MagicMock()
    app.state.format_checker = MagicMock()
    return app
