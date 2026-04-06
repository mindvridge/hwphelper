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

        # 4. 스타일 복원
        if original_char and preserve_style:
            try:
                self._ctrl.set_char_shape(original_char)
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
    # 셀 이동 (pyhwpx 네이티브)
    # ------------------------------------------------------------------

    def _navigate_to_cell(self, hwp: Any, table_idx: int, row: int, col: int) -> None:
        """표의 특정 셀로 정확하게 이동한다."""
        # 표 진입 (A1 셀로 이동)
        hwp.get_into_nth_table(table_idx)

        # 표의 열 수 파악 (get_cell_addr로)
        # A1에서 (row, col)까지 이동: row * cols + col 번 TableRightCell
        # 열 수를 모르면 한 행씩 이동
        target_moves = 0

        # 먼저 열 수를 파악: 오른쪽으로 이동하다가 행이 바뀌는 시점
        cols = self._detect_cols(hwp)

        target_moves = row * cols + col

        # A1에서 다시 시작
        hwp.get_into_nth_table(table_idx)

        for _ in range(target_moves):
            hwp.TableRightCell()

    def _detect_cols(self, hwp: Any) -> int:
        """현재 표의 열 수를 파악한다."""
        try:
            # get_cell_addr()가 "A1", "B1", "C1" 등을 반환
            # 오른쪽으로 이동하면서 행 번호가 바뀌면 열 수 확인
            start_addr = hwp.get_cell_addr()
            start_row = self._addr_to_row(start_addr)
            cols = 1
            for i in range(50):  # 최대 50열
                hwp.TableRightCell()
                addr = hwp.get_cell_addr()
                curr_row = self._addr_to_row(addr)
                if curr_row != start_row:
                    # 행이 바뀌었으면 이전까지가 열 수
                    return cols
                cols += 1
            return cols
        except Exception:
            return 1

    @staticmethod
    def _addr_to_row(addr: str) -> int:
        """셀 주소에서 행 번호를 추출한다. "A1" → 1, "B3" → 3"""
        try:
            digits = "".join(c for c in addr if c.isdigit())
            return int(digits) if digits else 0
        except Exception:
            return 0

    # ------------------------------------------------------------------
    # 텍스트 교체
    # ------------------------------------------------------------------

    def _select_and_replace(self, hwp: Any, text: str) -> None:
        """현재 셀의 텍스트를 선택하고 새 텍스트로 교체한다."""
        # 셀 내 전체 선택
        try:
            hwp.HAction.Run("MoveColBegin")
            hwp.HAction.Run("MoveSelColEnd")
        except Exception:
            try:
                hwp.HAction.Run("MoveLineBegin")
                hwp.HAction.Run("MoveSelLineEnd")
            except Exception:
                pass

        # 텍스트 삽입 (선택 영역 덮어쓰기)
        lines = text.split("\n")
        for i, line in enumerate(lines):
            if line:
                hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                hwp.HParameterSet.HInsertText.Text = line
                hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
            if i < len(lines) - 1:
                hwp.HAction.Run("BreakPara")
