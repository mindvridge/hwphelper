"""도구 정의 테스트."""

from __future__ import annotations

from src.ai.tool_definitions import DOCUMENT_MODIFYING_TOOLS, HWP_TOOLS, get_tools_for_provider


class TestHWPTools:
    """HWP 도구 정의 검증."""

    def test_tools_not_empty(self) -> None:
        assert len(HWP_TOOLS) > 0

    def test_all_tools_have_name(self) -> None:
        for t in HWP_TOOLS:
            assert "name" in t
            assert "description" in t
            assert "parameters" in t

    def test_required_tools_exist(self) -> None:
        names = {t["name"] for t in HWP_TOOLS}
        expected = {
            "analyze_document", "read_table", "read_cell", "write_cell",
            "fill_field", "fill_all_empty_cells", "validate_format",
            "undo", "save_document", "get_document_info",
        }
        assert expected.issubset(names)

    def test_modifying_tools_subset(self) -> None:
        all_names = {t["name"] for t in HWP_TOOLS}
        assert DOCUMENT_MODIFYING_TOOLS.issubset(all_names)

    def test_write_cell_params(self) -> None:
        tool = next(t for t in HWP_TOOLS if t["name"] == "write_cell")
        props = tool["parameters"]["properties"]
        assert "table_idx" in props
        assert "row" in props
        assert "col" in props
        assert "text" in props
        assert set(tool["parameters"]["required"]) == {"table_idx", "row", "col", "text"}


class TestProviderConversion:
    """프로바이더별 도구 형식 변환 테스트."""

    def test_anthropic_format(self) -> None:
        tools = get_tools_for_provider("anthropic")
        for t in tools:
            assert "name" in t
            assert "input_schema" in t
            assert "type" not in t

    def test_openai_format(self) -> None:
        tools = get_tools_for_provider("openai")
        for t in tools:
            assert t["type"] == "function"
            assert "function" in t
            assert "name" in t["function"]
            assert "parameters" in t["function"]
