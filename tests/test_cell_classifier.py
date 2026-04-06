"""셀 분류기 테스트."""

from __future__ import annotations

import pytest

from src.hwp_engine.cell_classifier import CellClassifier
from src.hwp_engine.table_reader import Cell, CellStyle, CellType, Table


@pytest.fixture
def classifier() -> CellClassifier:
    return CellClassifier()


@pytest.fixture
def empty_table() -> Table:
    return Table(table_idx=0, rows=1, cols=1, cells=[])


class TestCellClassify:
    """단일 셀 분류 테스트."""

    def test_empty_cell(self, classifier: CellClassifier, empty_table: Table) -> None:
        cell = Cell(row=1, col=1, text="")
        assert classifier.classify(cell, empty_table) == CellType.EMPTY

    def test_whitespace_only_is_empty(self, classifier: CellClassifier, empty_table: Table) -> None:
        cell = Cell(row=1, col=1, text="   \n  ")
        assert classifier.classify(cell, empty_table) == CellType.EMPTY

    def test_placeholder_keyword(self, classifier: CellClassifier, empty_table: Table) -> None:
        cell = Cell(row=1, col=1, text="(내용 입력)")
        assert classifier.classify(cell, empty_table) == CellType.PLACEHOLDER

    def test_placeholder_with_context(self, classifier: CellClassifier, empty_table: Table) -> None:
        cell = Cell(row=1, col=1, text="여기에 작성하세요")
        assert classifier.classify(cell, empty_table) == CellType.PLACEHOLDER

    def test_placeholder_symbols(self, classifier: CellClassifier, empty_table: Table) -> None:
        cell = Cell(row=1, col=1, text="○○○")
        assert classifier.classify(cell, empty_table) == CellType.PLACEHOLDER

    def test_label_keyword(self, classifier: CellClassifier, empty_table: Table) -> None:
        cell = Cell(row=1, col=1, text="사업명")
        assert classifier.classify(cell, empty_table) == CellType.LABEL

    def test_label_first_row(self, classifier: CellClassifier, empty_table: Table) -> None:
        cell = Cell(row=0, col=2, text="기타항목")
        assert classifier.classify(cell, empty_table) == CellType.LABEL

    def test_label_first_col_short(self, classifier: CellClassifier, empty_table: Table) -> None:
        cell = Cell(row=3, col=0, text="합계")
        assert classifier.classify(cell, empty_table) == CellType.LABEL

    def test_label_bold_short(self, classifier: CellClassifier, empty_table: Table) -> None:
        bold = CellStyle(bold=True)
        cell = Cell(row=2, col=2, text="소계", style=bold)
        assert classifier.classify(cell, empty_table) == CellType.LABEL

    def test_prefilled_long_text(self, classifier: CellClassifier, empty_table: Table) -> None:
        cell = Cell(row=2, col=2, text="2026년까지 월 매출 5000만원을 달성하고 글로벌 시장에 진출한다.", style=CellStyle(text_color="0x00000000"))
        assert classifier.classify(cell, empty_table) == CellType.PREFILLED

    def test_prefilled_non_keyword(self, classifier: CellClassifier, empty_table: Table) -> None:
        cell = Cell(row=2, col=2, text="인공지능 기술을 활용한 혁신적 서비스 개발")
        assert classifier.classify(cell, empty_table) == CellType.PREFILLED


class TestClassifyTable:
    """테이블 전체 분류 테스트."""

    def test_classify_sample_table(self, classifier: CellClassifier, sample_table: Table) -> None:
        result = classifier.classify_table(sample_table)

        # 사업명 → LABEL
        assert result.cells[0].cell_type == CellType.LABEL
        # 빈 셀 → EMPTY
        assert result.cells[1].cell_type == CellType.EMPTY
        # 비고 → LABEL
        assert result.cells[2].cell_type == CellType.LABEL
        # 기관명 → LABEL
        assert result.cells[3].cell_type == CellType.LABEL
        # (내용 입력) → PLACEHOLDER
        assert result.cells[4].cell_type == CellType.PLACEHOLDER
        # 연락처 → LABEL
        assert result.cells[5].cell_type == CellType.LABEL

    def test_classify_large_table(self, classifier: CellClassifier, large_table: Table) -> None:
        result = classifier.classify_table(large_table)

        # 헤더 행 전부 LABEL
        header_cells = [c for c in result.cells if c.row == 0]
        assert all(c.cell_type == CellType.LABEL for c in header_cells)

        # 빈 셀 확인
        empty_cells = result.empty_cells()
        assert len(empty_cells) >= 2  # (1,2), (2,2) 등

    def test_empty_cells_helper(self, classified_table: Table) -> None:
        empties = classified_table.empty_cells()
        # EMPTY + PLACEHOLDER 반환
        assert len(empties) == 2

    def test_label_cells_helper(self, classified_table: Table) -> None:
        labels = classified_table.label_cells()
        assert len(labels) == 4  # 사업명, 비고, 기관명, 연락처


class TestCustomKeywords:
    """커스텀 키워드 분류 테스트."""

    def test_custom_label_keywords(self) -> None:
        classifier = CellClassifier(label_keywords=["커스텀라벨"])
        table = Table(table_idx=0, rows=1, cols=1, cells=[])
        cell = Cell(row=2, col=2, text="커스텀라벨")
        assert classifier.classify(cell, table) == CellType.LABEL

    def test_custom_placeholder_keywords(self) -> None:
        classifier = CellClassifier(placeholder_keywords=["FILL_HERE"])
        table = Table(table_idx=0, rows=1, cols=1, cells=[])
        cell = Cell(row=2, col=2, text="FILL_HERE please")
        assert classifier.classify(cell, table) == CellType.PLACEHOLDER


class TestColoredTextDetection:
    """파란색/빨간색 예시글 감지 테스트."""

    def test_blue_text_is_placeholder(self) -> None:
        classifier = CellClassifier()
        table = Table(table_idx=0, rows=1, cols=1, cells=[])
        # 파란색 텍스트 (BGR: 0xFF0000 = 순수 파랑)
        blue_style = CellStyle(text_color="0x00FF0000")
        cell = Cell(row=1, col=1, text="본 과제의 핵심 기술은 AI 기반 자연어 처리입니다.", style=blue_style)
        assert classifier.classify(cell, table) == CellType.PLACEHOLDER

    def test_red_text_is_placeholder(self) -> None:
        classifier = CellClassifier()
        table = Table(table_idx=0, rows=1, cols=1, cells=[])
        # 빨간색 텍스트 (BGR: 0x0000FF = 빨강)
        red_style = CellStyle(text_color="0x000000FF")
        cell = Cell(row=1, col=1, text="여기에 기술 설명을 작성하세요", style=red_style)
        assert classifier.classify(cell, table) == CellType.PLACEHOLDER

    def test_black_text_is_not_placeholder(self) -> None:
        classifier = CellClassifier()
        table = Table(table_idx=0, rows=1, cols=1, cells=[])
        # 검정 텍스트
        black_style = CellStyle(text_color="0x00000000")
        cell = Cell(row=2, col=2, text="2026년 매출 5000만원을 달성하고 글로벌 시장에 진출한다.", style=black_style)
        assert classifier.classify(cell, table) == CellType.PREFILLED

    def test_dark_blue_is_placeholder(self) -> None:
        classifier = CellClassifier()
        table = Table(table_idx=0, rows=1, cols=1, cells=[])
        # 어두운 파란색 (한/글에서 자주 사용)
        style = CellStyle(text_color=str(0x800000))  # BGR: 어두운 파랑
        cell = Cell(row=1, col=1, text="예시 내용", style=style)
        assert classifier.classify(cell, table) == CellType.PLACEHOLDER

    def test_detect_colored_disabled(self) -> None:
        classifier = CellClassifier(detect_colored_text=False)
        table = Table(table_idx=0, rows=1, cols=1, cells=[])
        blue_style = CellStyle(text_color="0x00FF0000")
        cell = Cell(row=2, col=2, text="파란 텍스트지만 감지 비활성", style=blue_style)
        # 색상 감지 꺼짐 → PREFILLED
        assert classifier.classify(cell, table) == CellType.PREFILLED
