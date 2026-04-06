"""COM 컨트롤러 테스트.

실제 COM 테스트는 한/글이 설치된 환경에서만 실행.
기본 속성·로직은 단위 테스트.
"""

from __future__ import annotations

import pytest

from src.hwp_engine.com_controller import HAS_HWP, HwpController


class TestHwpControllerUnit:
    """COM 연결 없이 검증 가능한 단위 테스트."""

    def test_init_defaults(self) -> None:
        ctrl = HwpController()
        assert ctrl._visible is False
        assert ctrl._security_module is True
        assert ctrl._hwp is None
        assert ctrl.file_path is None

    def test_init_custom(self) -> None:
        ctrl = HwpController(visible=True, security_module=False)
        assert ctrl._visible is True
        assert ctrl._security_module is False

    def test_hwp_property_raises_without_com(self) -> None:
        """COM 없는 환경에서 hwp 접근 시 에러."""
        ctrl = HwpController()
        if not HAS_HWP:
            with pytest.raises(Exception):
                _ = ctrl.hwp

    def test_open_nonexistent_file(self) -> None:
        """존재하지 않는 파일 열기 시 FileNotFoundError."""
        ctrl = HwpController()
        ctrl._hwp = object()  # 더미 객체로 None 체크 우회
        with pytest.raises(FileNotFoundError):
            ctrl.open("/nonexistent/path/file.hwp")

    def test_context_manager_quit_called(self) -> None:
        """컨텍스트 매니저 종료 시 quit 호출 확인."""
        quit_called = False
        ctrl = HwpController()

        original_quit = ctrl.quit

        def mock_quit() -> None:
            nonlocal quit_called
            quit_called = True
            ctrl._hwp = None

        ctrl.quit = mock_quit  # type: ignore[assignment]
        ctrl._hwp = object()  # connect() 스킵

        with ctrl:
            pass

        assert quit_called


@pytest.mark.skipif(not HAS_HWP, reason="한/글 미설치")
class TestHwpControllerCOM:
    """실제 COM 연동 테스트.

    NOTE: 한/글 COM은 단일 프로세스에서 연속 생성/종료 시 불안정하므로
    하나의 테스트에서 connect → 검증 → quit을 모두 수행한다.
    """

    def test_connect_and_quit(self) -> None:
        ctrl = HwpController(visible=False)
        ctrl.connect()
        assert ctrl._hwp is not None
        ctrl.quit()
        assert ctrl._hwp is None
