"""문서 세션 관리 — 열기/닫기/스냅샷/되돌리기/다시실행.

대화형 편집의 핵심 모듈.
채팅에서 사용자가 "아까 수정한 거 되돌려줘"라고 하면 undo를 실행한다.

동작 원리:
- 문서를 열 때 원본을 uploads/에 보관하고 작업용 사본을 outputs/{session_id}/에 생성
- 매 편집 단위(표 채우기 등)마다 스냅샷(파일 복사)을 저장
- undo/redo는 스냅샷 파일을 다시 열어 적용
"""

from __future__ import annotations

import shutil
import uuid
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

import structlog

from .com_controller import HwpController

logger = structlog.get_logger()


@dataclass
class SnapshotInfo:
    """스냅샷 메타 정보."""

    index: int
    path: str
    description: str
    created_at: datetime


@dataclass
class DocumentSession:
    """문서 편집 세션."""

    session_id: str
    original_path: str          # 업로드된 원본
    working_path: str           # 현재 작업 중인 파일
    hwp_ctrl: HwpController
    schema: dict[str, Any] | None = None

    # 스냅샷 관리
    snapshots: list[str] = field(default_factory=list)             # 스냅샷 파일 경로
    snapshot_descriptions: list[str] = field(default_factory=list)  # 스냅샷 설명
    current_snapshot_idx: int = -1                                   # 현재 위치

    created_at: datetime = field(default_factory=datetime.now)


class DocumentManager:
    """문서 세션을 관리한다.

    세션 라이프사이클::

        session_id = manager.create_session("template.hwp")
        session = manager.get_session(session_id)
        # ... 편집 작업 ...
        manager.save_snapshot(session_id, "사업 개요 표 채우기")
        # ... 추가 편집 ...
        manager.undo(session_id)   # 마지막 스냅샷으로 되돌리기
        final_path = manager.close_session(session_id)
    """

    def __init__(
        self,
        upload_dir: str = "./uploads",
        output_dir: str = "./outputs",
    ) -> None:
        self._upload_dir = Path(upload_dir)
        self._output_dir = Path(output_dir)
        self._upload_dir.mkdir(parents=True, exist_ok=True)
        self._output_dir.mkdir(parents=True, exist_ok=True)
        self._sessions: dict[str, DocumentSession] = {}

    # ------------------------------------------------------------------
    # 세션 라이프사이클
    # ------------------------------------------------------------------

    def create_session(
        self,
        file_path: str,
        visible: bool = False,
    ) -> str:
        """문서 세션을 생성한다.

        1. 원본 파일을 uploads/에 복사
        2. 작업용 사본 생성 (outputs/{session_id}/)
        3. HwpController로 열기
        4. 초기 스냅샷 저장
        5. session_id 반환

        Parameters
        ----------
        file_path : str
            업로드된 HWP/HWPX 파일 경로.
        visible : bool
            한/글 창 표시 여부.

        Returns
        -------
        str
            세션 ID.
        """
        session_id = uuid.uuid4().hex[:12]
        src = Path(file_path).resolve()

        if not src.exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {src}")

        # 원본 보관
        upload_path = self._upload_dir / f"{session_id}_original{src.suffix}"
        shutil.copy2(src, upload_path)

        # 작업 디렉토리 생성
        work_dir = self._output_dir / session_id
        work_dir.mkdir(parents=True, exist_ok=True)
        working_path = work_dir / f"working{src.suffix}"
        shutil.copy2(src, working_path)

        # COM 컨트롤러 생성 및 문서 열기
        hwp_ctrl = HwpController(visible=visible)
        hwp_ctrl.connect()
        hwp_ctrl.open(str(working_path))

        session = DocumentSession(
            session_id=session_id,
            original_path=str(upload_path),
            working_path=str(working_path),
            hwp_ctrl=hwp_ctrl,
        )
        self._sessions[session_id] = session

        # 초기 스냅샷
        self._save_snapshot_file(session, "초기 상태")

        logger.info(
            "세션 생성 완료",
            session_id=session_id,
            original=str(src),
            working=str(working_path),
        )
        return session_id

    def get_session(self, session_id: str) -> DocumentSession:
        """세션을 조회한다."""
        session = self._sessions.get(session_id)
        if session is None:
            raise KeyError(f"세션을 찾을 수 없습니다: {session_id}")
        return session

    def close_session(self, session_id: str) -> str:
        """세션을 종료하고 최종 파일 경로를 반환한다."""
        session = self.get_session(session_id)

        # 최종 저장
        final_path = session.working_path
        try:
            session.hwp_ctrl.save()
        except Exception:
            logger.warning("세션 종료 시 저장 실패", session_id=session_id)

        # COM 종료
        try:
            session.hwp_ctrl.quit()
        except Exception:
            logger.debug("COM 종료 오류 (무시)")

        del self._sessions[session_id]
        logger.info("세션 종료", session_id=session_id, final_path=final_path)
        return final_path

    @property
    def active_sessions(self) -> list[str]:
        """활성 세션 ID 목록."""
        return list(self._sessions.keys())

    # ------------------------------------------------------------------
    # 스냅샷 / 되돌리기
    # ------------------------------------------------------------------

    def save_snapshot(self, session_id: str, description: str = "") -> int:
        """현재 상태를 스냅샷으로 저장한다.

        Parameters
        ----------
        description : str
            작업 설명 (예: "사업 개요 표 채우기").

        Returns
        -------
        int
            스냅샷 인덱스.
        """
        session = self.get_session(session_id)

        # 현재 상태 저장
        try:
            session.hwp_ctrl.save()
        except Exception:
            pass

        # redo 브랜치 제거 (중간 상태에서 새 작업 시)
        if session.current_snapshot_idx < len(session.snapshots) - 1:
            # 현재 위치 이후의 스냅샷 삭제
            for old_path in session.snapshots[session.current_snapshot_idx + 1 :]:
                try:
                    Path(old_path).unlink(missing_ok=True)
                except Exception:
                    pass
            session.snapshots = session.snapshots[: session.current_snapshot_idx + 1]
            session.snapshot_descriptions = session.snapshot_descriptions[
                : session.current_snapshot_idx + 1
            ]

        idx = self._save_snapshot_file(session, description)
        logger.info("스냅샷 저장", session_id=session_id, index=idx, description=description)
        return idx

    def undo(self, session_id: str) -> bool:
        """마지막 스냅샷으로 되돌린다.

        Returns
        -------
        bool
            되돌리기 성공 여부.
        """
        session = self.get_session(session_id)

        if session.current_snapshot_idx <= 0:
            logger.warning("더 이상 되돌릴 수 없습니다.", session_id=session_id)
            return False

        # 이전 스냅샷으로 이동
        session.current_snapshot_idx -= 1
        snapshot_path = session.snapshots[session.current_snapshot_idx]

        # 스냅샷 파일을 작업 파일에 복사하고 다시 열기
        self._restore_snapshot(session, snapshot_path)
        logger.info(
            "되돌리기 완료",
            session_id=session_id,
            index=session.current_snapshot_idx,
            description=session.snapshot_descriptions[session.current_snapshot_idx],
        )
        return True

    def redo(self, session_id: str) -> bool:
        """되돌리기를 취소한다 (다시 실행).

        Returns
        -------
        bool
            다시 실행 성공 여부.
        """
        session = self.get_session(session_id)

        if session.current_snapshot_idx >= len(session.snapshots) - 1:
            logger.warning("더 이상 다시 실행할 수 없습니다.", session_id=session_id)
            return False

        session.current_snapshot_idx += 1
        snapshot_path = session.snapshots[session.current_snapshot_idx]

        self._restore_snapshot(session, snapshot_path)
        logger.info(
            "다시 실행 완료",
            session_id=session_id,
            index=session.current_snapshot_idx,
            description=session.snapshot_descriptions[session.current_snapshot_idx],
        )
        return True

    def get_history(self, session_id: str) -> list[SnapshotInfo]:
        """작업 히스토리(스냅샷 목록)를 반환한다."""
        session = self.get_session(session_id)
        history: list[SnapshotInfo] = []

        for i, (path, desc) in enumerate(
            zip(session.snapshots, session.snapshot_descriptions)
        ):
            stat = Path(path).stat() if Path(path).exists() else None
            created = datetime.fromtimestamp(stat.st_mtime) if stat else session.created_at
            history.append(
                SnapshotInfo(index=i, path=path, description=desc, created_at=created)
            )

        return history

    # ------------------------------------------------------------------
    # 내부 유틸리티
    # ------------------------------------------------------------------

    def _save_snapshot_file(self, session: DocumentSession, description: str) -> int:
        """스냅샷 파일을 생성한다."""
        idx = len(session.snapshots)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        suffix = Path(session.working_path).suffix
        snapshot_dir = Path(session.working_path).parent / "snapshots"
        snapshot_dir.mkdir(exist_ok=True)

        snapshot_path = snapshot_dir / f"snap_{idx:03d}_{timestamp}{suffix}"
        shutil.copy2(session.working_path, snapshot_path)

        session.snapshots.append(str(snapshot_path))
        session.snapshot_descriptions.append(description or f"스냅샷 #{idx}")
        session.current_snapshot_idx = idx

        return idx

    def _restore_snapshot(self, session: DocumentSession, snapshot_path: str) -> None:
        """스냅샷 파일을 현재 작업 파일로 복원하고 다시 연다."""
        # 현재 문서 닫기
        try:
            session.hwp_ctrl.close()
        except Exception:
            pass

        # 스냅샷 → 작업 파일 복사
        shutil.copy2(snapshot_path, session.working_path)

        # 다시 열기
        session.hwp_ctrl.open(session.working_path)
