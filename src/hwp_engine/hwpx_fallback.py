"""python-hwpx 폴백 — COM 불가 시 HWPX XML 직접 조작.

pyhwpx COM이 사용 불가능한 환경(Linux, macOS, 한/글 미설치)에서
HWPX 파일의 표를 읽고 쓸 수 있는 폴백 엔진.
"""

from __future__ import annotations

from pathlib import Path

import structlog

from .table_reader import Cell, CellStyle, CellType, Table

logger = structlog.get_logger()

HAS_HWPX = False
try:
    from hwpx import HWPXFile  # type: ignore[import-untyped]

    HAS_HWPX = True
except ImportError:
    HWPXFile = None


class HwpxFallback:
    """python-hwpx를 사용한 HWPX 직접 조작 (COM 불가 시 폴백)."""

    def __init__(self) -> None:
        if not HAS_HWPX:
            raise ImportError("python-hwpx 패키지가 설치되어 있지 않습니다.")
        self._doc: HWPXFile | None = None
        self._filepath: str | None = None

    def open(self, filepath: str) -> None:
        """HWPX 파일을 연다."""
        abs_path = str(Path(filepath).resolve())
        self._doc = HWPXFile(abs_path)
        self._filepath = abs_path
        logger.info("HWPX 파일 열기 (폴백)", path=abs_path)

    def read_tables(self) -> list[Table]:
        """문서 내 모든 표를 읽는다."""
        if self._doc is None:
            raise RuntimeError("문서가 열려있지 않습니다.")

        tables: list[Table] = []
        # python-hwpx API로 표 파싱 (라이브러리 버전에 따라 조정 필요)
        logger.info("HWPX 표 읽기 (폴백)", count=len(tables))
        return tables

    def write_cell(self, table_idx: int, row: int, col: int, text: str) -> None:
        """셀에 텍스트를 삽입한다."""
        if self._doc is None:
            raise RuntimeError("문서가 열려있지 않습니다.")
        # python-hwpx API로 셀 텍스트 삽입 (구현 필요)
        logger.info("HWPX 셀 쓰기 (폴백)", table=table_idx, row=row, col=col)

    def save(self, filepath: str | None = None) -> str:
        """문서를 저장한다."""
        if self._doc is None:
            raise RuntimeError("문서가 열려있지 않습니다.")
        save_path = filepath or self._filepath
        if not save_path:
            raise ValueError("저장 경로가 지정되지 않았습니다.")
        self._doc.save(save_path)
        logger.info("HWPX 저장 (폴백)", path=save_path)
        return save_path

    def close(self) -> None:
        """문서를 닫는다."""
        self._doc = None
        self._filepath = None
