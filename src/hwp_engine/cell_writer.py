"""셀 텍스트 삽입 — 서식(CharShape/ParaShape)을 유지하면서 셀에 텍스트를 삽입.

핵심 전략:
1. 대상 셀로 커서 이동 (pyhwpx ShapeObjTableSelCell)
2. 기존 CharShape + ParaShape 캡처
3. 셀 내 텍스트만 선택 → 새 텍스트로 교체 (InsertText가 선택 영역을 덮어씀)
4. 표 구조는 건드리지 않음
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
        """지정된 셀에 텍스트를 삽입한다 (서식 보존, 표 구조 유지)."""
        hwp = self._ctrl.hwp

        # 1. 셀로 이동
        self._move_to_cell(table_idx, row, col)

        # 2. 기존 스타일 캡처
        original_char: dict[str, Any] | None = None
        if preserve_style:
            try:
                original_char = self._ctrl.get_char_shape()
            except Exception:
                pass

        # 3. 셀 내 텍스트만 선택 (표 구조는 건드리지 않음)
        self._select_cell_text_only(hwp)

        # 4. 스타일 복원 후 텍스트 삽입 (선택 영역을 덮어씀)
        if original_char and preserve_style:
            try:
                self._ctrl.set_char_shape(original_char)
            except Exception:
                pass

        # 5. 텍스트 삽입
        self._insert_text(hwp, text)

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
    # 셀 이동
    # ------------------------------------------------------------------

    def _move_to_cell(self, table_idx: int, row: int, col: int) -> None:
        """표 내 특정 셀로 커서를 이동한다."""
        hwp = self._ctrl.hwp
        try:
            if hasattr(hwp, "ShapeObjTableSelCell"):
                hwp.ShapeObjTableSelCell(table_idx, row, col)
                return
        except Exception:
            pass

        # 폴백: 탭으로 이동
        hwp.MovePos(2)
        ctrl = hwp.HeadCtrl
        idx = 0
        while ctrl:
            if ctrl.CtrlID == "tbl":
                if idx == table_idx:
                    try:
                        hwp.SetPosBySet(ctrl.GetAnchorPos(0))
                    except Exception:
                        pass
                    break
                idx += 1
            ctrl = ctrl.Next

        tbl = ctrl
        cols = 1
        if tbl:
            try:
                cols = tbl.ColCount
            except AttributeError:
                pass
        target = row * cols + col
        for _ in range(target):
            hwp.HAction.Run("TableRightCell")

    # ------------------------------------------------------------------
    # 셀 텍스트만 선택 (표 삭제 방지)
    # ------------------------------------------------------------------

    def _select_cell_text_only(self, hwp: Any) -> None:
        """현재 셀의 텍스트만 선택한다. 표 구조는 건드리지 않는다.

        방법: 셀 시작 → Shift+Ctrl+End (셀 내 텍스트만 선택)
        한/글에서 Ctrl+A는 표 전체를 선택할 수 있으므로 사용하지 않는다.
        """
        try:
            # 셀 처음으로 이동
            hwp.HAction.Run("MoveColBegin")
            # Shift+End로 셀 끝까지 선택
            hwp.HAction.Run("MoveSelColEnd")
        except Exception:
            try:
                # 폴백: Home → Shift+End
                hwp.HAction.Run("MoveLineBegin")
                hwp.HAction.Run("MoveSelLineEnd")
            except Exception:
                logger.debug("셀 텍스트 선택 실패")

    # ------------------------------------------------------------------
    # 텍스트 삽입
    # ------------------------------------------------------------------

    def _insert_text(self, hwp: Any, text: str) -> None:
        """현재 위치에 텍스트를 삽입한다 (선택 영역이 있으면 교체)."""
        lines = text.split("\n")
        for i, line in enumerate(lines):
            if line:
                hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                hwp.HParameterSet.HInsertText.Text = line
                hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
            if i < len(lines) - 1:
                hwp.HAction.Run("BreakPara")
