"""셀 쓰기 테스트.

COM 연동 테스트는 @pytest.mark.skipif(not HAS_HWP) 로 보호.
CellFill 데이터 클래스 테스트는 단위 테스트.
"""

from __future__ import annotations

import pytest

from src.hwp_engine.cell_writer import CellFill
from src.hwp_engine.com_controller import HAS_HWP


class TestCellFill:
    """CellFill 데이터 클래스 테스트."""

    def test_defaults(self) -> None:
        fill = CellFill(row=0, col=1, text="테스트")
        assert fill.row == 0
        assert fill.col == 1
        assert fill.text == "테스트"
        assert fill.preserve_style is True

    def test_no_preserve(self) -> None:
        fill = CellFill(row=0, col=1, text="테스트", preserve_style=False)
        assert fill.preserve_style is False

    def test_multiline(self) -> None:
        fill = CellFill(row=0, col=0, text="첫째 줄\n둘째 줄\n셋째 줄")
        assert "\n" in fill.text
        assert fill.text.count("\n") == 2


@pytest.mark.skipif(not HAS_HWP, reason="한/글 미설치")
class TestCellWriterCOM:
    """실제 COM 연동 셀 쓰기 테스트."""

    def test_write_cell_placeholder(self) -> None:
        # 실제 한/글 환경에서 실행
        pass
