"""스키마 생성기 테스트."""

from __future__ import annotations

import pytest

from src.hwp_engine.cell_classifier import CellClassifier
from src.hwp_engine.schema_generator import SchemaGenerator
from src.hwp_engine.table_reader import Table


@pytest.fixture
def generator() -> SchemaGenerator:
    return SchemaGenerator()


@pytest.fixture
def classifier() -> CellClassifier:
    return CellClassifier()


class TestSchemaGenerator:
    """SchemaGenerator 테스트."""

    def test_generate_single_table(
        self, generator: SchemaGenerator, classified_table: Table
    ) -> None:
        schema = generator.generate([classified_table], "test.hwp")
        assert schema["document_name"] == "test.hwp"
        assert schema["total_tables"] == 1
        assert schema["total_cells_to_fill"] == 2  # EMPTY + PLACEHOLDER

    def test_generate_table_schema(
        self, generator: SchemaGenerator, classified_table: Table
    ) -> None:
        tbl_schema = generator.generate_table_schema(classified_table)
        assert tbl_schema["table_idx"] == 0
        assert tbl_schema["rows"] == 2
        assert tbl_schema["cols"] == 3
        assert tbl_schema["cells_to_fill"] == 2

    def test_needs_fill_flag(
        self, generator: SchemaGenerator, classified_table: Table
    ) -> None:
        tbl_schema = generator.generate_table_schema(classified_table)
        for cell in tbl_schema["cells"]:
            if cell["cell_type"] in ("empty", "placeholder"):
                assert cell["needs_fill"] is True
            else:
                assert cell["needs_fill"] is False

    def test_cell_context_for_fill_targets(
        self, generator: SchemaGenerator, classified_table: Table
    ) -> None:
        tbl_schema = generator.generate_table_schema(classified_table)
        fill_cells = [c for c in tbl_schema["cells"] if c["needs_fill"]]
        for cell in fill_cells:
            assert "context" in cell
            # (0,1)의 왼쪽에 "사업명" 라벨이 있으므로 row_label 존재
            if cell["row"] == 0 and cell["col"] == 1:
                assert "row_label" in cell["context"]

    def test_generate_multiple_tables(
        self,
        generator: SchemaGenerator,
        classifier: CellClassifier,
        sample_table: Table,
        large_table: Table,
    ) -> None:
        t1 = classifier.classify_table(sample_table)
        t2 = classifier.classify_table(large_table)
        schema = generator.generate([t1, t2], "multi.hwp")
        assert schema["total_tables"] == 2
        assert schema["total_cells_to_fill"] > 0

    def test_empty_tables(self, generator: SchemaGenerator) -> None:
        schema = generator.generate([], "empty.hwp")
        assert schema["total_tables"] == 0
        assert schema["total_cells_to_fill"] == 0
        assert schema["tables"] == []
