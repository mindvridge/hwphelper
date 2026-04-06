"""문서 매니저 테스트.

COM 연동 없이 세션 관리 로직을 검증한다.
"""

from __future__ import annotations

import tempfile
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from src.hwp_engine.document_manager import DocumentManager, DocumentSession, SnapshotInfo


@pytest.fixture
def tmp_dirs(tmp_path: Path) -> tuple[Path, Path]:
    """임시 upload/output 디렉토리."""
    upload = tmp_path / "uploads"
    output = tmp_path / "outputs"
    upload.mkdir()
    output.mkdir()
    return upload, output


@pytest.fixture
def manager(tmp_dirs: tuple[Path, Path]) -> DocumentManager:
    """DocumentManager 인스턴스."""
    upload, output = tmp_dirs
    return DocumentManager(upload_dir=str(upload), output_dir=str(output))


@pytest.fixture
def dummy_hwp(tmp_path: Path) -> Path:
    """더미 HWP 파일."""
    f = tmp_path / "test.hwp"
    f.write_bytes(b"dummy hwp content")
    return f


class TestDocumentManagerUnit:
    """COM 없이 검증 가능한 단위 테스트."""

    def test_init_creates_dirs(self, tmp_path: Path) -> None:
        upload = tmp_path / "new_uploads"
        output = tmp_path / "new_outputs"
        dm = DocumentManager(upload_dir=str(upload), output_dir=str(output))
        assert upload.exists()
        assert output.exists()

    def test_active_sessions_empty(self, manager: DocumentManager) -> None:
        assert manager.active_sessions == []

    def test_get_session_not_found(self, manager: DocumentManager) -> None:
        with pytest.raises(KeyError):
            manager.get_session("nonexistent")

    @patch("src.hwp_engine.document_manager.HwpController")
    def test_create_session(
        self,
        mock_ctrl_cls: MagicMock,
        manager: DocumentManager,
        dummy_hwp: Path,
    ) -> None:
        """세션 생성 시 원본 복사, 작업 사본, 스냅샷 생성을 확인."""
        mock_ctrl = MagicMock()
        mock_ctrl_cls.return_value = mock_ctrl

        session_id = manager.create_session(str(dummy_hwp))

        assert session_id in manager.active_sessions
        session = manager.get_session(session_id)
        assert Path(session.original_path).exists()
        assert Path(session.working_path).exists()
        assert len(session.snapshots) == 1  # 초기 스냅샷
        assert session.snapshot_descriptions[0] == "초기 상태"
        mock_ctrl.connect.assert_called_once()
        mock_ctrl.open.assert_called_once()

    @patch("src.hwp_engine.document_manager.HwpController")
    def test_save_snapshot(
        self,
        mock_ctrl_cls: MagicMock,
        manager: DocumentManager,
        dummy_hwp: Path,
    ) -> None:
        mock_ctrl_cls.return_value = MagicMock()
        session_id = manager.create_session(str(dummy_hwp))

        idx = manager.save_snapshot(session_id, "표 1 채우기")
        assert idx == 1
        session = manager.get_session(session_id)
        assert len(session.snapshots) == 2
        assert session.snapshot_descriptions[1] == "표 1 채우기"

    @patch("src.hwp_engine.document_manager.HwpController")
    def test_undo_redo(
        self,
        mock_ctrl_cls: MagicMock,
        manager: DocumentManager,
        dummy_hwp: Path,
    ) -> None:
        mock_ctrl = MagicMock()
        mock_ctrl_cls.return_value = mock_ctrl

        session_id = manager.create_session(str(dummy_hwp))
        manager.save_snapshot(session_id, "작업 1")
        manager.save_snapshot(session_id, "작업 2")

        session = manager.get_session(session_id)
        assert session.current_snapshot_idx == 2

        # undo
        assert manager.undo(session_id) is True
        assert session.current_snapshot_idx == 1

        # undo 한 번 더
        assert manager.undo(session_id) is True
        assert session.current_snapshot_idx == 0

        # 더 이상 undo 불가
        assert manager.undo(session_id) is False

        # redo
        assert manager.redo(session_id) is True
        assert session.current_snapshot_idx == 1

        # redo 한 번 더
        assert manager.redo(session_id) is True
        assert session.current_snapshot_idx == 2

        # 더 이상 redo 불가
        assert manager.redo(session_id) is False

    @patch("src.hwp_engine.document_manager.HwpController")
    def test_undo_then_new_snapshot_clears_redo(
        self,
        mock_ctrl_cls: MagicMock,
        manager: DocumentManager,
        dummy_hwp: Path,
    ) -> None:
        """undo 후 새 작업 시 redo 브랜치가 제거되는지 확인."""
        mock_ctrl_cls.return_value = MagicMock()
        session_id = manager.create_session(str(dummy_hwp))

        manager.save_snapshot(session_id, "작업 1")
        manager.save_snapshot(session_id, "작업 2")
        manager.undo(session_id)  # 작업 1로

        # 새 작업
        manager.save_snapshot(session_id, "작업 3 (새 분기)")
        session = manager.get_session(session_id)

        # 작업 2는 제거되고, [초기, 작업1, 작업3]
        assert len(session.snapshots) == 3
        assert session.snapshot_descriptions[-1] == "작업 3 (새 분기)"
        assert manager.redo(session_id) is False  # redo 불가

    @patch("src.hwp_engine.document_manager.HwpController")
    def test_get_history(
        self,
        mock_ctrl_cls: MagicMock,
        manager: DocumentManager,
        dummy_hwp: Path,
    ) -> None:
        mock_ctrl_cls.return_value = MagicMock()
        session_id = manager.create_session(str(dummy_hwp))
        manager.save_snapshot(session_id, "작업 A")

        history = manager.get_history(session_id)
        assert len(history) == 2
        assert isinstance(history[0], SnapshotInfo)
        assert history[0].description == "초기 상태"
        assert history[1].description == "작업 A"

    @patch("src.hwp_engine.document_manager.HwpController")
    def test_close_session(
        self,
        mock_ctrl_cls: MagicMock,
        manager: DocumentManager,
        dummy_hwp: Path,
    ) -> None:
        mock_ctrl = MagicMock()
        mock_ctrl_cls.return_value = mock_ctrl

        session_id = manager.create_session(str(dummy_hwp))
        final = manager.close_session(session_id)

        assert session_id not in manager.active_sessions
        assert final is not None
        mock_ctrl.save.assert_called()
        mock_ctrl.quit.assert_called()

    def test_create_session_file_not_found(self, manager: DocumentManager) -> None:
        with pytest.raises(FileNotFoundError):
            manager.create_session("/nonexistent/file.hwp")
