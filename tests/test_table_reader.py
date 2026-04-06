"""표 구조 데이터 클래스 테스트.

COM 연동 테스트는 @pytest.mark.skipif(not HAS_HWP) 로 보호.
여기서는 데이터 클래스 자체의 동작을 검증한다.
"""

from __future__ import annotations

import pytest

from src.hwp_engine.com_controller import HAS_HWP
from src.hwp_engine.table_reader import Cell, CellStyle, CellType, Table


class TestCellStyle:
    """CellStyle 데이터 클래스 테스트."""

    def test_defaults(self) -> None:
        style = CellStyle()
        assert style.font_name == ""
        assert style.font_size == 10.0
        assert style.bold is False
        assert style.alignment == "left"

    def test_custom_values(self) -> None:
        style = CellStyle(font_name="맑은 고딕", font_size=12.0, bold=True, alignment="center")
        assert style.font_name == "맑은 고딕"
        assert style.font_size == 12.0
        assert style.bold is True
        assert style.alignment == "center"


class TestCell:
    """Cell 데이터 클래스 테스트."""

    def test_defaults(self) -> None:
        cell = Cell(row=0, col=0)
        assert cell.text == ""
        assert cell.row_span == 1
        assert cell.col_span == 1
        assert cell.cell_type == CellType.UNKNOWN
        assert cell.style is None

    def test_with_span(self) -> None:
        cell = Cell(row=0, col=0, row_span=2, col_span=3, text="병합 셀")
        assert cell.row_span == 2
        assert cell.col_span == 3


class TestTable:
    """Table 데이터 클래스 테스트."""

    def test_get_cell(self, sample_table: Table) -> None:
        cell = sample_table.get_cell(0, 0)
        assert cell is not None
        assert cell.text == "사업명"

    def test_get_cell_not_found(self, sample_table: Table) -> None:
        cell = sample_table.get_cell(99, 99)
        assert cell is None

    def test_empty_cells(self, classified_table: Table) -> None:
        empties = classified_table.empty_cells()
        positions = {(c.row, c.col) for c in empties}
        assert (0, 1) in positions  # 빈 셀
        assert (1, 1) in positions  # 플레이스홀더

    def test_label_cells(self, classified_table: Table) -> None:
        labels = classified_table.label_cells()
        label_texts = {c.text for c in labels}
        assert "사업명" in label_texts
        assert "기관명" in label_texts

    def test_to_dict(self, sample_table: Table) -> None:
        d = sample_table.to_dict()
        assert d["table_idx"] == 0
        assert d["rows"] == 2
        assert d["cols"] == 3
        assert len(d["cells"]) == 6
        assert d["cells"][0]["text"] == "사업명"
        assert d["cells"][0]["style"]["bold"] is True


class TestTableReaderCOM:
    """COM 연동 TableReader 테스트 (한/글 필요)."""

    @pytest.mark.skipif(not HAS_HWP, reason="한/글 미설치")
    def test_get_table_count(self) -> None:
        # 실제 COM 테스트는 한/글이 설치된 환경에서만 실행
        from src.hwp_engine.com_controller import HwpController
        from src.hwp_engine.table_reader import TableReader

        with HwpController(visible=False) as ctrl:
            reader = TableReader(ctrl)
            count = reader.get_table_count()
            assert isinstance(count, int)
