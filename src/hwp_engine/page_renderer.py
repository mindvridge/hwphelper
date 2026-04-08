"""HWP 문서를 페이지별 이미지(PNG)로 변환한다.

HWP → PDF 내보내기 → PyMuPDF로 페이지별 PNG 렌더링.

사용법::

    renderer = PageRenderer(ctrl)
    pages = renderer.render_all_pages()
    page0_b64 = pages[0].base64_data  # LLM API 전달용
"""

from __future__ import annotations

import base64
import os
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING

import structlog

if TYPE_CHECKING:
    from .com_controller import HwpController

logger = structlog.get_logger()


@dataclass
class PageImage:
    """렌더링된 한 페이지 이미지."""

    page_num: int
    image_bytes: bytes
    width: int
    height: int
    mime_type: str = "image/png"

    @property
    def base64_data(self) -> str:
        return base64.b64encode(self.image_bytes).decode("ascii")

    @property
    def data_url(self) -> str:
        return f"data:{self.mime_type};base64,{self.base64_data}"


class PageRenderer:
    """HWP 문서를 페이지별 PNG 이미지로 렌더링한다."""

    def __init__(self, hwp_ctrl: HwpController, dpi: int = 150) -> None:
        self._ctrl = hwp_ctrl
        self._dpi = dpi
        self._cache: dict[str, list[PageImage]] = {}

    def render_all_pages(self, force: bool = False) -> list[PageImage]:
        """모든 페이지를 PNG로 렌더링한다.

        결과는 캐시되며, force=True로 강제 갱신 가능.
        """
        cache_key = self._ctrl.file_path or "unknown"
        if not force and cache_key in self._cache:
            return self._cache[cache_key]

        pages = self._render_via_pdf()
        self._cache[cache_key] = pages
        logger.info("페이지 렌더링 완료", pages=len(pages), dpi=self._dpi)
        return pages

    def render_page(self, page_num: int) -> PageImage | None:
        """특정 페이지만 렌더링한다 (0-based)."""
        pages = self.render_all_pages()
        if 0 <= page_num < len(pages):
            return pages[page_num]
        return None

    def invalidate_cache(self) -> None:
        """캐시를 무효화한다. 문서 수정 후 호출."""
        self._cache.clear()

    def get_page_count(self) -> int:
        """문서 페이지 수를 반환한다."""
        try:
            return self._ctrl.hwp.PageCount
        except Exception:
            return 0

    def _render_via_pdf(self) -> list[PageImage]:
        """PDF 내보내기 → PyMuPDF로 변환."""
        import fitz  # PyMuPDF

        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = os.path.join(tmpdir, "doc.pdf")

            # HWP → PDF 내보내기
            try:
                self._ctrl.hwp.SaveAs(os.path.abspath(pdf_path), "PDF")
            except Exception as e:
                logger.warning("PDF 내보내기 실패", error=str(e))
                return []

            if not Path(pdf_path).exists():
                logger.warning("PDF 파일 생성 안 됨")
                return []

            # PDF → PNG
            pages: list[PageImage] = []
            doc = fitz.open(pdf_path)
            try:
                for i in range(len(doc)):
                    page = doc[i]
                    mat = fitz.Matrix(self._dpi / 72, self._dpi / 72)
                    pix = page.get_pixmap(matrix=mat)
                    png_bytes = pix.tobytes("png")

                    pages.append(PageImage(
                        page_num=i,
                        image_bytes=png_bytes,
                        width=pix.width,
                        height=pix.height,
                    ))
            finally:
                doc.close()

        return pages
