"""필드 매니저 테스트."""

from __future__ import annotations

import pytest

from src.hwp_engine.com_controller import HAS_HWP
from src.hwp_engine.field_manager import FieldInfo


class TestFieldInfo:
    """FieldInfo 데이터 클래스 테스트."""

    def test_defaults(self) -> None:
        info = FieldInfo(name="사업명")
        assert info.name == "사업명"
        assert info.value == ""
        assert info.direction == ""

    def test_with_value(self) -> None:
        info = FieldInfo(name="사업명", value="AI 자동화", direction="사업 이름을 입력하세요")
        assert info.value == "AI 자동화"
        assert info.direction == "사업 이름을 입력하세요"


@pytest.mark.skipif(not HAS_HWP, reason="한/글 미설치")
class TestFieldManagerCOM:
    """실제 COM 연동 필드 매니저 테스트."""

    def test_list_fields(self) -> None:
        pass
