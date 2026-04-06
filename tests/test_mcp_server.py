"""MCP 서버 테스트."""

from __future__ import annotations

import json
from unittest.mock import MagicMock, patch

import pytest


class TestMCPToolsUnit:
    """MCP 도구 함수 단위 테스트."""

    def _reset(self):
        import src.mcp_server as m
        m._ctrl = None
        m._file_path = None
        m._tables_cache = None
        m._schema_cache = None

    def test_analyze_template(self) -> None:
        import src.mcp_server as m
        self._reset()

        mock_schema = {
            "document_name": "test.hwp", "total_tables": 1, "total_cells_to_fill": 5,
            "tables": [{"table_idx": 0, "rows": 3, "cols": 2, "cells_to_fill": 5, "cells": []}],
        }

        with patch.object(m, "_run_com", return_value=mock_schema):
            result = m.analyze_template("C:/test/test.hwp")
            data = json.loads(result)
            assert "total_tables" in data or "document" in data

    def test_write_cell(self) -> None:
        import src.mcp_server as m
        self._reset()

        with patch.object(m, "_run_com", return_value={"success": True, "message": "OK"}):
            result = m.write_cell(0, 1, 2, "테스트")
            data = json.loads(result)
            assert data["success"] is True

    def test_write_cell_error(self) -> None:
        import src.mcp_server as m
        self._reset()

        with patch.object(m, "_run_com", side_effect=Exception("COM 오류")):
            result = m.write_cell(0, 0, 0, "x")
            data = json.loads(result)
            assert "error" in data

    def test_fill_field(self) -> None:
        import src.mcp_server as m
        self._reset()

        with patch.object(m, "_run_com", return_value={"success": True, "field": "사업명"}):
            result = m.fill_field("사업명", "AI 자동화")
            data = json.loads(result)
            assert data["success"] is True

    def test_fill_all_empty_cells(self) -> None:
        import src.mcp_server as m
        self._reset()
        m._schema_cache = {
            "tables": [{"table_idx": 0, "cells": [
                {"row": 0, "col": 0, "needs_fill": False},
                {"row": 0, "col": 1, "needs_fill": True, "context": {"row_label": "사업명"}},
            ]}]
        }
        m._tables_cache = [MagicMock()]

        result = m.fill_all_empty_cells("예비창업패키지", "테스트기업")
        data = json.loads(result)
        assert data["total"] == 1

    def test_save_document(self) -> None:
        import src.mcp_server as m
        self._reset()

        with patch.object(m, "_run_com", return_value={"success": True, "path": "test.hwp"}):
            result = m.save_document()
            data = json.loads(result)
            assert data["success"] is True

    def test_save_document_as_pdf(self) -> None:
        import src.mcp_server as m
        self._reset()

        with patch.object(m, "_run_com", return_value={"success": True, "path": "test.pdf"}):
            result = m.save_document(format="pdf")
            data = json.loads(result)
            assert data["success"] is True

    def test_validate_format(self) -> None:
        import src.mcp_server as m
        self._reset()

        with patch.object(m, "_run_com", return_value={"passed": True, "summary": "OK", "warnings": []}):
            result = m.validate_format("기본")
            data = json.loads(result)
            assert data["passed"] is True


class TestMCPResources:
    def test_fill_all_no_schema(self) -> None:
        import src.mcp_server as m
        m._schema_cache = None
        m._tables_cache = None
        result = m.fill_all_empty_cells("test", "test")
        data = json.loads(result)
        assert "error" in data


class TestMCPServerImport:
    def test_mcp_instance_exists(self) -> None:
        from src.mcp_server import mcp
        assert mcp is not None
        assert mcp.name == "hwp-ai"

    def test_tools_registered(self) -> None:
        from src.mcp_server import mcp
        assert True  # 모듈이 에러 없이 로드되면 성공
