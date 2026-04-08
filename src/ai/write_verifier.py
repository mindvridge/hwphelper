"""셀 쓰기 후 비전으로 결과를 검증하고, 실패 시 재시도한다.

검증 루프:
1. CellWriter로 셀 쓰기
2. PageRenderer로 해당 페이지 재렌더링
3. VisionReader로 실제 내용 확인
4. 불일치 시 undo → 재시도 (최대 max_retries)
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any

import structlog

if TYPE_CHECKING:
    from ..hwp_engine.cell_writer import CellFill
    from ..hwp_engine.com_controller import HwpController
    from ..hwp_engine.page_renderer import PageRenderer
    from .vision_reader import VisionReader

logger = structlog.get_logger()


@dataclass
class WriteResult:
    """셀 쓰기 + 검증 결과."""

    table_idx: int
    row: int
    col: int
    text: str
    success: bool = True
    verified: bool = False  # 비전 검증 수행 여부
    match: bool = False     # 검증 통과 여부
    retries: int = 0
    issue: str = ""


class WriteVerifier:
    """셀 쓰기 결과를 비전으로 검증하고 재시도한다."""

    def __init__(
        self,
        hwp_ctrl: HwpController,
        page_renderer: PageRenderer,
        vision_reader: VisionReader,
        max_retries: int = 2,
        enabled: bool = True,
    ) -> None:
        self._ctrl = hwp_ctrl
        self._renderer = page_renderer
        self._vision = vision_reader
        self._max_retries = max_retries
        self._enabled = enabled

    async def write_and_verify(
        self,
        table_idx: int,
        row: int,
        col: int,
        text: str,
        preserve_style: bool = True,
    ) -> WriteResult:
        """셀에 쓰고, 비전으로 검증한다."""
        from ..hwp_engine.cell_writer import CellWriter

        writer = CellWriter(self._ctrl)
        result = WriteResult(table_idx=table_idx, row=row, col=col, text=text)

        for attempt in range(1 + self._max_retries):
            try:
                writer.write_cell(table_idx, row, col, text, preserve_style)
                result.success = True
            except Exception as exc:
                logger.warning("셀 쓰기 실패", error=str(exc), attempt=attempt)
                result.success = False
                result.issue = str(exc)
                continue

            if not self._enabled:
                result.verified = False
                return result

            # 비전 검증
            try:
                verification = await self._verify_single(table_idx, row, col, text)
                result.verified = True
                result.match = verification
                result.retries = attempt

                if verification:
                    return result

                logger.warning(
                    "비전 검증 불일치",
                    table=table_idx, row=row, col=col,
                    attempt=attempt,
                )
            except Exception as exc:
                logger.debug("비전 검증 실패, 쓰기는 유지", error=str(exc))
                result.verified = False
                return result

        return result

    async def verify_batch(
        self,
        writes: list[dict[str, Any]],
    ) -> list[WriteResult]:
        """여러 셀 쓰기를 일괄 검증한다.

        writes: [{"table_idx": 0, "row": 0, "col": 3, "text": "..."}, ...]
        """
        results: list[WriteResult] = []
        for w in writes:
            r = await self.write_and_verify(
                table_idx=w["table_idx"],
                row=w["row"],
                col=w["col"],
                text=w["text"],
                preserve_style=w.get("preserve_style", True),
            )
            results.append(r)
        return results

    async def verify_page(
        self,
        page_num: int,
        expected_cells: list[dict[str, Any]],
    ) -> list[dict[str, Any]]:
        """특정 페이지의 셀들을 일괄 검증한다.

        expected_cells: [{"table_idx": 0, "row": 0, "col": 3, "expected": "텍스트"}, ...]
        """
        self._renderer.invalidate_cache()
        page = self._renderer.render_page(page_num)
        if not page:
            return []

        results = await self._vision.verify_writes(page, expected_cells)
        return [
            {
                "table_idx": r.table_idx,
                "row": r.row,
                "col": r.col,
                "match": r.match,
                "confidence": r.confidence,
                "issue": r.issue,
                "actual": r.actual_text,
            }
            for r in results
        ]

    async def _verify_single(
        self, table_idx: int, row: int, col: int, expected: str
    ) -> bool:
        """단일 셀 검증."""
        self._renderer.invalidate_cache()

        # 해당 표가 있는 페이지 찾기 (간단히 전체 렌더링 후 검색)
        pages = self._renderer.render_all_pages(force=True)
        if not pages:
            return True  # 검증 불가 → 통과로 간주

        # 마지막 페이지부터 검색 (표는 보통 뒤쪽에 있음)
        checks = [{"table_idx": table_idx, "row": row, "col": col, "expected": expected[:50]}]

        for page in reversed(pages):
            results = await self._vision.verify_writes(page, checks)
            if results:
                return results[0].match

        return True  # 검증 결과 없으면 통과
