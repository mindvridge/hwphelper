"""COM 추출 결과와 비전 인식 결과를 병합하여 정확도를 높인다.

병합 규칙:
1. 구조(행/열 수): 비전 우선 (COM은 병합 셀에서 자주 오류)
2. 셀 텍스트: 양쪽 일치 시 확정, 불일치 시 비전 우선
3. 병합 정보: 비전 우선
4. 셀 분류: 비전의 색상 정보 + COM의 텍스트 분석 결합
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import structlog

from ..hwp_engine.table_reader import Cell, CellStyle, CellType, Table
from .vision_reader import VisionCell, VisionPageResult, VisionTable

logger = structlog.get_logger()


@dataclass
class ReconcileStats:
    """병합 통계."""

    total_cells: int = 0
    agreed: int = 0          # 양쪽 일치
    vision_preferred: int = 0  # 비전 채택
    com_preferred: int = 0     # COM 채택
    vision_only: int = 0       # 비전에만 존재
    com_only: int = 0          # COM에만 존재


class VisionReconciler:
    """COM과 비전 결과를 병합한다."""

    def reconcile_table(
        self,
        com_table: Table,
        vision_table: VisionTable,
    ) -> tuple[Table, ReconcileStats]:
        """한 표의 COM/비전 결과를 병합한다."""
        stats = ReconcileStats()

        # 구조: 비전 우선
        rows = vision_table.rows or com_table.rows
        cols = vision_table.cols or com_table.cols

        # 셀 병합
        merged_cells: list[Cell] = []
        vision_map: dict[tuple[int, int], VisionCell] = {
            (vc.row, vc.col): vc for vc in vision_table.cells
        }
        com_map: dict[tuple[int, int], Cell] = {
            (c.row, c.col): c for c in com_table.cells
        }

        # 모든 좌표 합집합
        all_coords = set(vision_map.keys()) | set(com_map.keys())
        stats.total_cells = len(all_coords)

        for coord in sorted(all_coords):
            vc = vision_map.get(coord)
            cc = com_map.get(coord)

            if vc and cc:
                # 양쪽 모두 존재
                v_text = vc.text.strip()
                c_text = cc.text.strip()

                if v_text == c_text:
                    stats.agreed += 1
                    text = c_text
                elif not c_text and v_text:
                    # COM이 못 읽은 셀 → 비전 채택
                    stats.vision_preferred += 1
                    text = v_text
                elif c_text and not v_text:
                    # 비전이 못 읽은 셀 → COM 채택
                    stats.com_preferred += 1
                    text = c_text
                else:
                    # 양쪽 다르면 비전 우선 (COM이 더 자주 틀림)
                    stats.vision_preferred += 1
                    text = v_text

                # 병합 정보: 비전 우선
                row_span = vc.row_span if vc.row_span > 1 else cc.row_span
                col_span = vc.col_span if vc.col_span > 1 else cc.col_span

                # 셀 타입: 비전의 시각 정보 활용
                cell_type = self._classify_cell(vc, cc)

                merged_cells.append(Cell(
                    row=coord[0],
                    col=coord[1],
                    row_span=row_span,
                    col_span=col_span,
                    text=text,
                    cell_type=cell_type,
                    style=cc.style,
                ))

            elif vc and not cc:
                stats.vision_only += 1
                merged_cells.append(Cell(
                    row=coord[0],
                    col=coord[1],
                    row_span=vc.row_span,
                    col_span=vc.col_span,
                    text=vc.text,
                    cell_type=CellType.EMPTY if vc.is_empty else CellType.UNKNOWN,
                ))

            elif cc and not vc:
                stats.com_only += 1
                merged_cells.append(cc)

        result = Table(
            table_idx=com_table.table_idx,
            rows=rows,
            cols=cols,
            cells=merged_cells,
        )

        logger.info(
            "표 병합 완료",
            table_idx=com_table.table_idx,
            agreed=stats.agreed,
            vision_preferred=stats.vision_preferred,
            com_preferred=stats.com_preferred,
        )

        return result, stats

    def reconcile_all(
        self,
        com_tables: list[Table],
        vision_results: list[VisionPageResult],
    ) -> list[tuple[Table, ReconcileStats]]:
        """모든 표를 병합한다.

        비전 결과의 table_idx와 COM의 table_idx를 매칭.
        """
        # 비전 결과에서 모든 표 추출
        vision_tables: dict[int, VisionTable] = {}
        for vpr in vision_results:
            for vt in vpr.tables:
                vision_tables[vt.table_idx] = vt

        results: list[tuple[Table, ReconcileStats]] = []
        for ct in com_tables:
            vt = vision_tables.get(ct.table_idx)
            if vt:
                merged, stats = self.reconcile_table(ct, vt)
                results.append((merged, stats))
            else:
                # 비전에서 못 찾은 표 → COM 결과 그대로
                results.append((ct, ReconcileStats(
                    total_cells=len(ct.cells),
                    com_only=len(ct.cells),
                )))

        return results

    @staticmethod
    def _classify_cell(vc: VisionCell, cc: Cell) -> CellType:
        """비전+COM 정보로 셀 타입을 결정한다."""
        if vc.is_guide:
            return CellType.PLACEHOLDER
        if vc.is_label:
            return CellType.LABEL
        if vc.is_empty:
            return CellType.EMPTY

        # 비전이 분류 못한 경우 COM 분류 사용
        if cc.cell_type != CellType.UNKNOWN:
            return cc.cell_type

        # 텍스트 기반 추정
        text = (vc.text or cc.text).strip()
        if not text:
            return CellType.EMPTY
        if text.startswith("※"):
            return CellType.PLACEHOLDER
        if vc.color == "blue":
            return CellType.PLACEHOLDER

        return CellType.PREFILLED
