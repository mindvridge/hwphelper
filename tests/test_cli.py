"""CLI 테스트.

Typer CliRunner로 각 명령어의 동작을 검증한다.
CLI 함수 내부에서 지역 import를 사용하므로 원본 모듈 경로로 패치한다.
"""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
from typer.testing import CliRunner

from src.cli import app

runner = CliRunner()


class TestCLIHelp:
    """도움말 및 기본 동작."""

    def test_no_args_shows_help(self) -> None:
        result = runner.invoke(app, [])
        assert result.exit_code == 0
        assert "HWP-AI AutoFill" in result.output

    def test_help_flag(self) -> None:
        result = runner.invoke(app, ["--help"])
        assert result.exit_code == 0
        assert "serve" in result.output
        assert "analyze" in result.output
        assert "generate" in result.output
        assert "validate" in result.output
        assert "models" in result.output
        assert "ingest" in result.output
        assert "list-fields" in result.output
        assert "add-fields" in result.output


class TestModelsCommand:
    """models 명령어 테스트."""

    def test_models_list(self) -> None:
        mock_model = MagicMock()
        mock_model.id = "claude-sonnet"
        mock_model.provider = "anthropic"
        mock_model.model = "claude-sonnet-4"
        mock_model.available = True
        mock_model.description = "Claude Sonnet"
        mock_model.estimated_cost_per_1k = 0.003

        with patch("src.ai.llm_router.LLMRouter.list_models", return_value=[mock_model]), \
             patch("src.ai.llm_router.LLMRouter.default_model", new_callable=lambda: property(lambda self: "claude-sonnet")):
            result = runner.invoke(app, ["models"])

        assert result.exit_code == 0
        assert "claude-sonnet" in result.output


class TestAnalyzeCommand:
    """analyze 명령어 테스트."""

    def test_analyze_file_not_found(self, tmp_path: Path) -> None:
        result = runner.invoke(app, ["analyze", str(tmp_path / "nope.hwp")])
        assert result.exit_code == 1
        assert "찾을 수 없습니다" in result.output

    def test_analyze_success(self, tmp_path: Path) -> None:
        dummy = tmp_path / "test.hwp"
        dummy.write_bytes(b"dummy")

        mock_ctrl = MagicMock()

        with (
            patch("src.hwp_engine.com_controller.HwpController") as MockCtrl,
            patch("src.hwp_engine.table_reader.TableReader") as MockReader,
            patch("src.hwp_engine.cell_classifier.CellClassifier") as MockClassifier,
            patch("src.hwp_engine.schema_generator.SchemaGenerator") as MockGenerator,
        ):
            MockCtrl.return_value.__enter__ = MagicMock(return_value=mock_ctrl)
            MockCtrl.return_value.__exit__ = MagicMock(return_value=False)

            MockReader.return_value.read_all_tables.return_value = []
            MockGenerator.return_value.generate.return_value = {
                "document_name": "test.hwp",
                "total_tables": 1,
                "total_cells_to_fill": 3,
                "tables": [
                    {
                        "table_idx": 0,
                        "rows": 2,
                        "cols": 2,
                        "cells": [
                            {"row": 0, "col": 0, "cell_type": "label", "text": "사업명", "needs_fill": False},
                            {"row": 0, "col": 1, "cell_type": "empty", "text": "", "needs_fill": True},
                        ],
                    }
                ],
            }

            result = runner.invoke(app, ["analyze", str(dummy)])

        assert result.exit_code == 0
        assert "표 개수" in result.output


class TestValidateCommand:
    """validate 명령어 테스트."""

    def test_validate_file_not_found(self, tmp_path: Path) -> None:
        result = runner.invoke(app, ["validate", str(tmp_path / "nope.hwp")])
        assert result.exit_code == 1

    def test_validate_pass(self, tmp_path: Path) -> None:
        dummy = tmp_path / "test.hwp"
        dummy.write_bytes(b"dummy")

        mock_report = MagicMock()
        mock_report.passed = True
        mock_report.total_checks = 4
        mock_report.passed_checks = 4
        mock_report.warnings = []
        mock_report.errors = []

        with (
            patch("src.hwp_engine.com_controller.HwpController") as MockCtrl,
            patch("src.validator.format_checker.FormatChecker") as MockChecker,
        ):
            MockCtrl.return_value.__enter__ = MagicMock(return_value=MagicMock())
            MockCtrl.return_value.__exit__ = MagicMock(return_value=False)
            MockChecker.return_value.check_document.return_value = mock_report

            result = runner.invoke(app, ["validate", str(dummy), "-p", "기본"])

        assert result.exit_code == 0
        assert "PASS" in result.output


class TestIngestCommand:
    """ingest 명령어 테스트."""

    def test_ingest_success(self, tmp_path: Path) -> None:
        txt = tmp_path / "doc.txt"
        txt.write_text("테스트 문서", encoding="utf-8")

        with patch("src.ai.rag_engine.RAGEngine.ingest_document", return_value=5):
            result = runner.invoke(app, ["ingest", str(txt), "-p", "TIPS"])

        assert result.exit_code == 0
        assert "5개 청크" in result.output

    def test_ingest_file_not_found(self, tmp_path: Path) -> None:
        result = runner.invoke(app, ["ingest", str(tmp_path / "nope.txt")])
        assert result.exit_code == 1


class TestListFieldsCommand:
    """list-fields 명령어 테스트."""

    def test_list_fields_empty(self, tmp_path: Path) -> None:
        dummy = tmp_path / "test.hwp"
        dummy.write_bytes(b"dummy")

        with (
            patch("src.hwp_engine.com_controller.HwpController") as MockCtrl,
            patch("src.hwp_engine.field_manager.FieldManager") as MockFM,
        ):
            MockCtrl.return_value.__enter__ = MagicMock(return_value=MagicMock())
            MockCtrl.return_value.__exit__ = MagicMock(return_value=False)
            MockFM.return_value.list_fields.return_value = []

            result = runner.invoke(app, ["list-fields", str(dummy)])

        assert result.exit_code == 0
        assert "없습니다" in result.output

    def test_list_fields_with_data(self, tmp_path: Path) -> None:
        dummy = tmp_path / "test.hwp"
        dummy.write_bytes(b"dummy")

        mock_field = MagicMock()
        mock_field.name = "사업명"
        mock_field.value = "AI 프로젝트"

        with (
            patch("src.hwp_engine.com_controller.HwpController") as MockCtrl,
            patch("src.hwp_engine.field_manager.FieldManager") as MockFM,
        ):
            MockCtrl.return_value.__enter__ = MagicMock(return_value=MagicMock())
            MockCtrl.return_value.__exit__ = MagicMock(return_value=False)
            MockFM.return_value.list_fields.return_value = [mock_field]

            result = runner.invoke(app, ["list-fields", str(dummy)])

        assert result.exit_code == 0
        assert "사업명" in result.output


class TestServeCommand:
    """serve 명령어 테스트."""

    def test_serve_invokes_uvicorn(self) -> None:
        with patch("uvicorn.run") as mock_run:
            result = runner.invoke(app, ["serve", "--no-open", "--port", "9999"])

        assert result.exit_code == 0
        mock_run.assert_called_once()
        args, kwargs = mock_run.call_args
        assert kwargs.get("port") == 9999 or args == ("src.server:app",)
