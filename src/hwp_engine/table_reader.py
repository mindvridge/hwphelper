"""표 구조 읽기 — HWP 문서의 표를 파싱하여 구조화된 데이터로 변환.

한/글 COM API의 ``GetText()`` 와 ``Action`` 을 사용하여
표의 셀 텍스트, 병합 정보, 글자/문단 스타일을 추출한다.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import TYPE_CHECKING, Any

import structlog

if TYPE_CHECKING:
    from .com_controller import HwpController

logger = structlog.get_logger()


# ------------------------------------------------------------------
# 데이터 클래스
# ------------------------------------------------------------------

class CellType(str, Enum):
    """셀 타입 (분류 전 기본값은 UNKNOWN)."""

    LABEL = "label"
    EMPTY = "empty"
    PREFILLED = "prefilled"
    PLACEHOLDER = "placeholder"
    UNKNOWN = "unknown"


@dataclass
class CellStyle:
    """셀의 글자·문단 스타일 정보."""

    font_name: str = ""
    font_size: float = 10.0       # pt
    bold: bool = False
    italic: bool = False
    char_spacing: float = 0.0     # %
    line_spacing: float = 160.0   # %
    alignment: str = "left"
    text_color: str = "0x00000000"


@dataclass
class Cell:
    """표 셀 하나."""

    row: int
    col: int
    row_span: int = 1
    col_span: int = 1
    text: str = ""
    cell_type: CellType = CellType.UNKNOWN
    style: CellStyle | None = None


@dataclass
class Table:
    """표 전체."""

    table_idx: int
    rows: int = 0
    cols: int = 0
    cells: list[Cell] = field(default_factory=list)

    # ----- 편의 메서드 -----

    def get_cell(self, row: int, col: int) -> Cell | None:
        """(row, col)에 해당하는 셀을 반환한다."""
        for c in self.cells:
            if c.row == row and c.col == col:
                return c
        return None

    def empty_cells(self) -> list[Cell]:
        """채워야 할 빈 셀 목록."""
        return [c for c in self.cells if c.cell_type in (CellType.EMPTY, CellType.PLACEHOLDER)]

    def label_cells(self) -> list[Cell]:
        """라벨(항목명) 셀 목록."""
        return [c for c in self.cells if c.cell_type == CellType.LABEL]

    def to_dict(self) -> dict[str, Any]:
        """직렬화용 딕셔너리 변환."""
        return {
            "table_idx": self.table_idx,
            "rows": self.rows,
            "cols": self.cols,
            "cells": [
                {
                    "row": c.row,
                    "col": c.col,
                    "row_span": c.row_span,
                    "col_span": c.col_span,
                    "text": c.text,
                    "cell_type": c.cell_type.value,
                    "style": {
                        "font_name": c.style.font_name,
                        "font_size": c.style.font_size,
                        "bold": c.style.bold,
                        "italic": c.style.italic,
                        "char_spacing": c.style.char_spacing,
                        "line_spacing": c.style.line_spacing,
                        "alignment": c.style.alignment,
                        "text_color": c.style.text_color,
                    }
                    if c.style
                    else None,
                }
                for c in self.cells
            ],
        }


# ------------------------------------------------------------------
# TableReader
# ------------------------------------------------------------------

# 한/글 Alignment 상수 → 문자열
_ALIGNMENT_MAP: dict[int, str] = {
    0: "justify",
    1: "left",
    2: "right",
    3: "center",
    4: "distribute",
    5: "distribute_space",
}


class TableReader:
    """HWP 문서의 표를 읽어서 구조화된 ``Table`` 객체로 변환한다."""

    def __init__(self, hwp_ctrl: HwpController) -> None:
        self._ctrl = hwp_ctrl

    # ------------------------------------------------------------------
    # public API
    # ------------------------------------------------------------------

    def get_table_count(self) -> int:
        """문서 내 표 개수를 반환한다."""
        hwp = self._ctrl.hwp

        # 문서 처음으로 이동
        hwp.MovePos(2)  # moveBOF

        count = 0
        # 컨트롤 순회하면서 표(tbl) 개수를 센다
        ctrl = hwp.HeadCtrl
        while ctrl:
            if ctrl.CtrlID == "tbl":
                count += 1
            ctrl = ctrl.Next

        logger.info("표 개수 조회", count=count)
        return count

    def read_table(self, table_idx: int) -> Table:
        """지정된 인덱스의 표를 읽어 Table 객체로 반환한다.

        pyhwpx의 get_into_nth_table + TableRightCell로 순회.
        """
        hwp = self._ctrl.hwp

        # 표 진입 (A1)
        hwp.get_into_nth_table(table_idx)

        # 열 수 파악: A1에서 오른쪽으로 이동하며 행 번호 변화 감지
        cols = 1
        start_row = self._addr_row(hwp.get_cell_addr())
        for _ in range(100):
            hwp.TableRightCell()
            addr = hwp.get_cell_addr()
            if self._addr_row(addr) != start_row:
                break
            cols += 1

        # 다시 처음으로
        hwp.get_into_nth_table(table_idx)

        # 전체 셀 순회
        cells: list[Cell] = []
        max_row = 0
        prev_addr = ""
        for _ in range(5000):  # 안전 상한
            addr = hwp.get_cell_addr()
            if addr == prev_addr and len(cells) > 0:
                break  # 더 이상 이동 안 됨
            prev_addr = addr

            r = self._addr_row(addr) - 1  # 1-based → 0-based
            c = self._addr_col(addr)
            max_row = max(max_row, r)

            # 텍스트 읽기
            text = self._read_cell_text_native(hwp)

            # 스타일 읽기
            style = self._read_current_cell_style()

            cells.append(cell)

            # 다음 셀로 이동
            hwp.TableRightCell()

        rows = max_row + 1
        table = Table(table_idx=table_idx, rows=rows, cols=cols, cells=cells)

        logger.info("표 읽기 완료", table_idx=table_idx, rows=rows, cols=cols, cells=len(cells))
        return table

    def read_all_tables(self) -> list[Table]:
        """문서 내 모든 표를 읽는다."""
        count = self.get_table_count()
        tables: list[Table] = []
        for i in range(count):
            try:
                tables.append(self.read_table(i))
            except Exception:
                logger.warning("표 읽기 실패, 건너뜀", table_idx=i)
        logger.info("전체 표 읽기 완료", total=count, success=len(tables))
        return tables

    def read_cell_style(self, table_idx: int, row: int, col: int) -> CellStyle:
        """특정 셀의 스타일만 읽는다."""
        if not self._move_to_cell(table_idx, row, col):
            raise ValueError(f"셀 ({row}, {col})로 이동할 수 없습니다.")
        return self._read_current_cell_style()

    # ------------------------------------------------------------------
    # 주소 파싱 유틸리티
    # ------------------------------------------------------------------

    @staticmethod
    def _addr_row(addr: str) -> int:
        """셀 주소에서 행 번호 추출. 'A1' → 1, 'C3' → 3"""
        digits = "".join(c for c in addr if c.isdigit())
        return int(digits) if digits else 0

    @staticmethod
    def _addr_col(addr: str) -> int:
        """셀 주소에서 열 인덱스 추출. 'A1' → 0, 'B1' → 1, 'AA1' → 26"""
        letters = "".join(c for c in addr if c.isalpha())
        col = 0
        for ch in letters.upper():
            col = col * 26 + (ord(ch) - ord("A") + 1)
        return col - 1  # 0-based

    def _read_cell_text_native(self, hwp: Any) -> str:
        """현재 셀의 텍스트를 pyhwpx로 읽는다."""
        try:
            # 셀 전체 선택 후 텍스트 가져오기
            hwp.HAction.Run("MoveColBegin")
            hwp.HAction.Run("MoveSelColEnd")
            text = hwp.GetTextFile("TEXT", "")
            # 선택 해제
            hwp.HAction.Run("MoveColBegin")
            return text.strip() if text else ""
        except Exception:
            return ""

    # ------------------------------------------------------------------
    # 내부 유틸리티 (레거시 — 폴백용)
    # ------------------------------------------------------------------

    def _get_table_ctrl(self, table_idx: int) -> Any:
        """table_idx번째 표 컨트롤을 반환한다."""
        hwp = self._ctrl.hwp
        ctrl = hwp.HeadCtrl
        idx = 0
        while ctrl:
            if ctrl.CtrlID == "tbl":
                if idx == table_idx:
                    return ctrl
                idx += 1
            ctrl = ctrl.Next
        return None

    def _get_table_size(self, tbl_ctrl: Any) -> tuple[int, int]:
        """표 컨트롤에서 (행, 열) 크기를 추출한다."""
        try:
            # pyhwpx 스타일
            rows = tbl_ctrl.RowCount
            cols = tbl_ctrl.ColCount
            return rows, cols
        except AttributeError:
            pass

        try:
            tset = tbl_ctrl.TableSet
            return tset.Rows, tset.Cols
        except AttributeError:
            pass

        # 폴백: 셀을 순회하며 추정
        logger.warning("표 크기를 직접 추출할 수 없어 순회 추정합니다.")
        return self._estimate_table_size(tbl_ctrl)

    def _estimate_table_size(self, tbl_ctrl: Any) -> tuple[int, int]:
        """셀 순회로 표 크기를 추정한다."""
        max_row, max_col = 0, 0
        cell = tbl_ctrl.FirstCell if hasattr(tbl_ctrl, "FirstCell") else None
        while cell:
            r = getattr(cell, "Row", 0)
            c = getattr(cell, "Col", 0)
            max_row = max(max_row, r)
            max_col = max(max_col, c)
            cell = cell.Next if hasattr(cell, "Next") else None
        return max_row + 1, max_col + 1

    def _move_to_table(self, table_idx: int) -> bool:
        """table_idx번째 표로 커서를 이동한다."""
        hwp = self._ctrl.hwp
        hwp.MovePos(2)  # 문서 처음
        for _ in range(table_idx + 1):
            # 다음 표 찾기
            hwp.HAction.Run("MoveNextParaBegin")

        # 표 안으로 진입
        ctrl = self._get_table_ctrl(table_idx)
        if ctrl:
            try:
                hwp.SetPosBySet(ctrl.GetAnchorPos(0))
                return True
            except Exception:
                pass
        return False

    def _move_to_cell(self, table_idx: int, row: int, col: int) -> bool:
        """표 내 특정 셀로 커서를 이동한다."""
        hwp = self._ctrl.hwp
        try:
            # pyhwpx의 표 셀 이동 API 시도
            if hasattr(hwp, "ShapeObjTableSelCell"):
                hwp.ShapeObjTableSelCell(table_idx, row, col)
                return True
        except Exception:
            pass

        try:
            # 대안: 표 컨트롤에서 셀 리스트 접근
            tbl = self._get_table_ctrl(table_idx)
            if tbl:
                cell_list = tbl.CellList
                # row * cols + col 로 셀 인덱스 계산
                _, cols = self._get_table_size(tbl)
                cell_idx = row * cols + col
                if cell_idx < cell_list.Count:
                    cell = cell_list.Item(cell_idx)
                    hwp.SetPosBySet(cell.GetAnchorPos(0))
                    return True
        except Exception:
            pass

        # 최종 폴백: 탭 키로 이동
        return self._move_to_cell_by_tab(table_idx, row, col)

    def _move_to_cell_by_tab(self, table_idx: int, table_row: int, table_col: int) -> bool:
        """탭 키를 사용하여 셀 이동 (폴백)."""
        hwp = self._ctrl.hwp
        try:
            self._move_to_table(table_idx)
            _, cols = 0, 1
            tbl = self._get_table_ctrl(table_idx)
            if tbl:
                _, cols = self._get_table_size(tbl)

            target = table_row * cols + table_col
            for _ in range(target):
                hwp.HAction.Run("TableRightCell")
            return True
        except Exception:
            logger.debug("탭 이동 실패", row=table_row, col=table_col)
            return False

    def _read_cell_text(self) -> str:
        """현재 셀의 전체 텍스트를 읽는다."""
        hwp = self._ctrl.hwp
        parts: list[str] = []

        # 셀 전체 선택
        try:
            hwp.HAction.Run("SelectAll")
            text = hwp.GetTextFile("TEXT", "")
            if text:
                return text.strip()
        except Exception:
            pass

        # 폴백: GetText 반복
        try:
            while True:
                code, text = hwp.GetText()
                if code in (0, 1):  # 0=끝, 1=다른 컨트롤
                    break
                if text:
                    parts.append(text)
        except Exception:
            pass

        return "".join(parts).strip()

    def _read_current_cell_style(self) -> CellStyle:
        """현재 커서 위치의 셀 스타일을 읽는다."""
        try:
            char = self._ctrl.get_char_shape()
            para = self._ctrl.get_para_shape()

            alignment_raw = para.get("alignment", 1)
            alignment = _ALIGNMENT_MAP.get(alignment_raw, "left") if isinstance(alignment_raw, int) else str(alignment_raw)

            return CellStyle(
                font_name=str(char.get("font_name", "")),
                font_size=float(char.get("font_size", 10.0)),
                bold=bool(char.get("bold", False)),
                italic=bool(char.get("italic", False)),
                char_spacing=float(char.get("char_spacing", 0)),
                line_spacing=float(para.get("line_spacing", 160)),
                alignment=alignment,
                text_color=str(char.get("text_color", "0x00000000")),
            )
        except Exception:
            logger.debug("셀 스타일 읽기 실패, 기본값 반환")
            return CellStyle()

    def _get_cell_span(self, row: int, col: int, tbl_ctrl: Any) -> tuple[int, int]:
        """셀의 병합 범위(row_span, col_span)를 가져온다."""
        try:
            _, cols = self._get_table_size(tbl_ctrl)
            cell_list = tbl_ctrl.CellList
            cell_idx = row * cols + col
            if cell_idx < cell_list.Count:
                cell = cell_list.Item(cell_idx)
                rs = getattr(cell, "RowSpan", 1) or 1
                cs = getattr(cell, "ColSpan", 1) or 1
                return rs, cs
        except Exception:
            pass
        return 1, 1
