"""JSON 스키마 생성 — 표 구조를 LLM에게 전달할 구조화된 JSON으로 변환.

출력 형태::

    {
        "document_name": "2024_예비창업패키지_사업계획서.hwp",
        "total_tables": 5,
        "total_cells_to_fill": 23,
        "tables": [
            {
                "table_idx": 0,
                "rows": 4,
                "cols": 3,
                "cells": [
                    {"row": 0, "col": 0, "text": "사업명", "cell_type": "label", ...},
                    {"row": 0, "col": 1, "text": "",       "cell_type": "empty", ...},
                    ...
                ]
            },
            ...
        ]
    }
"""

from __future__ import annotations

from typing import Any

import structlog

from .table_reader import Cell, CellType, Table

logger = structlog.get_logger()


class SchemaGenerator:
    """표 구조를 LLM 프롬프트에 적합한 JSON 스키마로 변환한다."""

    def generate(
        self,
        tables: list[Table],
        document_name: str = "",
    ) -> dict[str, Any]:
        """문서 전체의 표 스키마를 생성한다.

        Parameters
        ----------
        tables : list[Table]
            분류 완료된 Table 객체 목록.
        document_name : str
            문서 파일명 (메타 정보용).

        Returns
        -------
        dict
            LLM에 전달할 구조화된 스키마.
        """
        total_fill = 0
        table_schemas: list[dict[str, Any]] = []

        for table in tables:
            tbl_schema = self._generate_table_schema(table)
            fill_count = tbl_schema["cells_to_fill"]
            total_fill += fill_count
            table_schemas.append(tbl_schema)

        schema = {
            "document_name": document_name,
            "total_tables": len(tables),
            "total_cells_to_fill": total_fill,
            "tables": table_schemas,
        }

        logger.info(
            "문서 스키마 생성 완료",
            document=document_name,
            tables=len(tables),
            fill_targets=total_fill,
        )
        return schema

    def generate_table_schema(self, table: Table) -> dict[str, Any]:
        """단일 표의 스키마를 생성한다 (외부 호출용)."""
        return self._generate_table_schema(table)

    # ------------------------------------------------------------------
    # 내부 유틸리티
    # ------------------------------------------------------------------

    def _generate_table_schema(self, table: Table) -> dict[str, Any]:
        """단일 표의 스키마를 생성한다."""
        cells_data: list[dict[str, Any]] = []
        fill_count = 0

        for cell in table.cells:
            needs_fill = cell.cell_type in (CellType.EMPTY, CellType.PLACEHOLDER)
            if needs_fill:
                fill_count += 1

            cell_data: dict[str, Any] = {
                "row": cell.row,
                "col": cell.col,
                "text": cell.text,
                "cell_type": cell.cell_type.value,
                "needs_fill": needs_fill,
                "row_span": cell.row_span,
                "col_span": cell.col_span,
            }

            # 채울 셀에는 주변 라벨 정보 추가
            if needs_fill:
                cell_data["context"] = self._build_cell_context(cell, table)

            cells_data.append(cell_data)

        return {
            "table_idx": table.table_idx,
            "rows": table.rows,
            "cols": table.cols,
            "cells_to_fill": fill_count,
            "cells": cells_data,
        }

    def _build_cell_context(self, target: Cell, table: Table) -> dict[str, str]:
        """대상 셀의 주변 라벨/값 정보를 구성한다.

        LLM이 셀의 맥락을 이해하도록 왼쪽 라벨, 위쪽 헤더 등을 제공한다.
        """
        context: dict[str, str] = {}

        # 같은 행의 왼쪽 라벨
        left_labels: list[str] = []
        for c in table.cells:
            if c.row == target.row and c.col < target.col and c.cell_type == CellType.LABEL:
                left_labels.append(c.text)
        if left_labels:
            context["row_label"] = " > ".join(left_labels)

        # 같은 열의 위쪽 헤더
        top_headers: list[str] = []
        for c in table.cells:
            if c.col == target.col and c.row < target.row and c.cell_type == CellType.LABEL:
                top_headers.append(c.text)
        if top_headers:
            context["col_header"] = " > ".join(top_headers)

        # 첫 행(헤더) 정보
        first_row_labels: list[str] = []
        for c in table.cells:
            if c.row == 0 and c.cell_type == CellType.LABEL:
                first_row_labels.append(f"[{c.col}]{c.text}")
        if first_row_labels:
            context["table_header"] = ", ".join(first_row_labels)

        # 같은 행의 기 작성 내용 (참고용)
        same_row_prefilled: list[str] = []
        for c in table.cells:
            if c.row == target.row and c.col != target.col and c.cell_type == CellType.PREFILLED:
                same_row_prefilled.append(f"[{c.col}]{c.text[:50]}")
        if same_row_prefilled:
            context["same_row_content"] = ", ".join(same_row_prefilled)

        return context
