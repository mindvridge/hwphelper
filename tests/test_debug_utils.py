"""디버그 유틸리티 테스트."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from src.utils import debug_utils
from src.utils.debug_utils import (
    dump_table_schema,
    setup_logging,
)


class TestSetupLogging:
    def test_debug_mode(self) -> None:
        setup_logging(debug=True)

    def test_production_mode(self) -> None:
        setup_logging(debug=False)


class TestDumpTableSchema:
    def test_dump_to_string(self) -> None:
        schema = {"table_idx": 0, "rows": 2, "cols": 2}
        result = dump_table_schema(schema)
        assert '"table_idx": 0' in result

    def test_dump_to_file(self, tmp_path: Path) -> None:
        schema = {"table_idx": 1}
        filepath = str(tmp_path / "schema.json")
        dump_table_schema(schema, filepath)
        assert Path(filepath).exists()
        content = Path(filepath).read_text(encoding="utf-8")
        assert "table_idx" in content


class TestCOMDiagnostic:
    def test_com_connection_returns_dict(self) -> None:
        """COM 연결 테스트 — 실제 COM 호출 대신 mock 사용."""
        mock_ctrl = MagicMock()
        mock_ctrl.hwp.Version = "12.0"

        with patch("src.hwp_engine.com_controller.HwpController", return_value=mock_ctrl):
            result = debug_utils.test_com_connection()

        assert isinstance(result, dict)
        assert "success" in result
        assert "error" in result
