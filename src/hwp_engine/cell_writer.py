"""셀 텍스트 삽입 — pyhwpx 네이티브 API로 정확한 셀 이동 + 서식 보존 쓰기.

핵심:
1. get_into_nth_table(n)으로 표 진입
2. TableRightCell()로 정확한 셀 이동
3. 기존 텍스트 선택 → 새 텍스트로 교체
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

import structlog

if TYPE_CHECKING:
    from .com_controller import HwpController

logger = structlog.get_logger()


@dataclass
class CellFill:
    """셀 채우기 정보."""

    row: int
    col: int
    text: str
    preserve_style: bool = True


class CellWriter:
    """표 셀에 텍스트를 서식 유지하면서 삽입한다."""

    def __init__(self, hwp_ctrl: HwpController) -> None:
        self._ctrl = hwp_ctrl

    def write_cell(
        self,
        table_idx: int,
        row: int,
        col: int,
        text: str,
        preserve_style: bool = True,
    ) -> None:
        """지정된 셀에 텍스트를 삽입한다."""
        hwp = self._ctrl.hwp

        # 1. 표 진입 + 셀 이동
        self._navigate_to_cell(hwp, table_idx, row, col)

        # 2. 기존 스타일 캡처
        original_char: dict[str, Any] | None = None
        if preserve_style:
            try:
                original_char = self._ctrl.get_char_shape()
            except Exception:
                pass

        # 3. 셀 내 텍스트 선택 → 교체
        self._select_and_replace(hwp, text)

        # 4. 삽입한 텍스트 전체를 선택하여 원래 스타일 적용
        if original_char and preserve_style:
            try:
                hwp.HAction.Run("SelectAll")
                self._ctrl.set_char_shape(original_char)
                hwp.Cancel()
            except Exception:
                pass

        logger.info("셀 쓰기 완료", table=table_idx, row=row, col=col, length=len(text))

    def write_cells_batch(self, table_idx: int, fills: list[CellFill]) -> list[bool]:
        """여러 셀에 텍스트를 일괄 삽입한다."""
        results: list[bool] = []
        for fill in fills:
            try:
                self.write_cell(table_idx, fill.row, fill.col, fill.text, fill.preserve_style)
                results.append(True)
            except Exception:
                logger.warning("셀 쓰기 실패", table=table_idx, row=fill.row, col=fill.col)
                results.append(False)
        return results

    # ------------------------------------------------------------------
    # 셀 이동 (셀 주소 기반 — 병합 셀 안전)
    # ------------------------------------------------------------------

    def _navigate_to_cell(self, hwp: Any, table_idx: int, row: int, col: int) -> None:
        """표의 특정 셀로 이동한다.

        1차: 셀 주소(get_cell_addr) 매칭으로 탐색
        2차 (주소 불가 시): XML 셀 순서 기반 인덱스 이동
        """
        hwp.get_into_nth_table(table_idx)

        first_addr = hwp.get_cell_addr()

        if isinstance(first_addr, str):
            # 주소 기반 탐색
            target_addr = f"{self._col_to_letter(col)}{row + 1}"
            if first_addr == target_addr:
                return
            for _ in range(500):
                hwp.TableRightCell()
                if hwp.get_cell_addr() == target_addr:
                    return
            logger.warning("주소 기반 셀 탐색 실패, 인덱스 폴백", target=target_addr)
            hwp.get_into_nth_table(table_idx)

        # 인덱스 기반 폴백: XML 셀 순서대로 TableRightCell 이동
        # XML에서 셀은 행 순서대로 나열되므로, (row, col) 이전의 셀 개수를 세어 이동
        target_idx = self._calc_cell_index(hwp, table_idx, row, col)
        hwp.get_into_nth_table(table_idx)
        for _ in range(target_idx):
            hwp.TableRightCell()

    def _calc_cell_index(self, hwp: Any, table_idx: int, target_row: int, target_col: int) -> int:
        """XML 기반으로 (row, col) 셀이 TableRightCell 순서에서 몇 번째인지 계산한다."""
        import xml.etree.ElementTree as ET
        try:
            hwp.get_into_nth_table(table_idx, select_cell=True)
            hwp.TableCellBlockExtend()
            hwp.TableCellBlockExtendAbs()
            xml_data = hwp.GetTextFile("HWPML2X", "saveblock")
            hwp.Cancel()

            root = ET.fromstring(xml_data)
            ns = root.tag.split("}")[0] + "}" if root.tag.startswith("{") else ""
            table_el = root.find(f".//{ns}TABLE")
            if table_el is None:
                return 0

            idx = 0
            for row_el in table_el.findall(f"{ns}ROW"):
                for cell_el in row_el.findall(f"{ns}CELL"):
                    r = int(cell_el.attrib.get("RowAddr", "0"))
                    c = int(cell_el.attrib.get("ColAddr", "0"))
                    if r == target_row and c == target_col:
                        return idx
                    idx += 1
        except Exception as exc:
            logger.debug("XML 셀 인덱스 계산 실패", error=str(exc))
        return 0

    @staticmethod
    def _col_to_letter(col: int) -> str:
        """0-based 열 인덱스를 알파벳 문자로 변환. 0→A, 1→B, 25→Z, 26→AA"""
        result = ""
        n = col + 1  # 1-based
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            result = chr(65 + remainder) + result
        return result

    # ------------------------------------------------------------------
    # 텍스트 교체
    # ------------------------------------------------------------------

    def _select_and_replace(self, hwp: Any, text: str) -> None:
        """현재 셀의 텍스트를 선택하고 새 텍스트로 교체한다."""
        # 셀 내 전체 선택 → 삭제
        hwp.HAction.Run("SelectAll")
        hwp.HAction.Run("Delete")

        # 텍스트 삽입
        lines = text.split("\n")
        for i, line in enumerate(lines):
            if line:
                hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                hwp.HParameterSet.HInsertText.Text = line
                hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
            if i < len(lines) - 1:
                hwp.HAction.Run("BreakPara")
