"""셀 콘텐츠 생성기 — LLM을 호출하여 빈 셀의 내용을 생성.

표 스키마와 사업 정보를 기반으로 각 셀에 적합한 콘텐츠를 생성한다.
셀 단위 독립 호출 또는 배치 호출을 지원한다.
"""

from __future__ import annotations

import asyncio
import json
from typing import Any, Callable

import structlog

from src.hwp_engine.cell_writer import CellFill

from .llm_router import LLMResponse, LLMRouter
from .prompt_builder import PromptBuilder
from .rag_engine import RAGEngine

logger = structlog.get_logger()


class CellGenerator:
    """표 셀 단위로 LLM을 호출하여 콘텐츠를 생성한다."""

    def __init__(
        self,
        llm_router: LLMRouter,
        rag_engine: RAGEngine | None = None,
    ) -> None:
        self._router = llm_router
        self._rag = rag_engine
        self._prompt_builder = PromptBuilder()

    async def generate_single_cell(
        self,
        cell_schema: dict[str, Any],
        program_name: str = "",
        company_info: str = "",
        model_id: str | None = None,
    ) -> str:
        """단일 셀의 내용을 생성한다."""
        # RAG 컨텍스트 검색
        rag_context = ""
        if self._rag:
            label = cell_schema.get("context", {}).get("row_label", "")
            if label:
                rag_context = self._rag.get_context(label, program_name)

        prompt = self._prompt_builder.build_cell_prompt(
            cell_schema=cell_schema,
            program_name=program_name,
            company_info=company_info,
            rag_context=rag_context,
        )

        messages = [
            {"role": "system", "content": self._prompt_builder.system_prompt},
            {"role": "user", "content": prompt},
        ]

        response = await self._router.chat(messages, model_id=model_id)
        if not isinstance(response, LLMResponse):
            raise TypeError(f"Expected LLMResponse, got {type(response)}")

        content = response.content.strip()
        # 따옴표 제거 (LLM이 따옴표로 감싸는 경우)
        if content.startswith('"') and content.endswith('"'):
            content = content[1:-1]

        logger.info(
            "셀 콘텐츠 생성",
            row=cell_schema.get("row"),
            col=cell_schema.get("col"),
            length=len(content),
        )
        return content

    async def generate_batch(
        self,
        cells: list[dict[str, Any]],
        program_name: str = "",
        company_info: str = "",
        model_id: str | None = None,
    ) -> list[CellFill]:
        """여러 셀의 내용을 한 번의 LLM 호출로 생성한다."""
        prompt = self._prompt_builder.build_batch_prompt(
            cells=cells,
            program_name=program_name,
            company_info=company_info,
        )

        messages = [
            {"role": "system", "content": self._prompt_builder.system_prompt},
            {"role": "user", "content": prompt},
        ]

        response = await self._router.chat(messages, model_id=model_id, max_tokens=8192)
        if not isinstance(response, LLMResponse):
            raise TypeError(f"Expected LLMResponse, got {type(response)}")

        fills = self._parse_batch_response(response.content)
        logger.info("배치 생성 완료", requested=len(cells), generated=len(fills))
        return fills

    async def generate_all(
        self,
        schema: dict[str, Any],
        program_name: str = "",
        company_info: str = "",
        model_id: str | None = None,
        concurrency: int = 3,
        on_progress: Callable[[int, int, dict], Any] | None = None,
    ) -> list[CellFill]:
        """문서 전체의 빈 셀 콘텐츠를 생성한다.

        Parameters
        ----------
        schema : dict
            SchemaGenerator가 생성한 문서 스키마.
        concurrency : int
            동시 LLM 호출 수.
        on_progress : Callable
            진행 콜백 (current, total, cell_info).

        Returns
        -------
        list[CellFill]
        """
        # 모든 빈 셀 수집
        all_cells: list[tuple[int, dict]] = []
        for table in schema.get("tables", []):
            table_idx = table["table_idx"]
            for cell in table["cells"]:
                if cell.get("needs_fill"):
                    all_cells.append((table_idx, cell))

        total = len(all_cells)
        if total == 0:
            logger.info("채울 셀이 없습니다.")
            return []

        results: list[CellFill] = []
        semaphore = asyncio.Semaphore(concurrency)
        completed = 0

        async def _generate_one(table_idx: int, cell: dict) -> CellFill | None:
            nonlocal completed
            async with semaphore:
                try:
                    content = await self.generate_single_cell(
                        cell_schema=cell,
                        program_name=program_name,
                        company_info=company_info,
                        model_id=model_id,
                    )
                    completed += 1
                    if on_progress:
                        await on_progress(completed, total, cell)
                    return CellFill(
                        row=cell["row"],
                        col=cell["col"],
                        text=content,
                    )
                except Exception as exc:
                    logger.warning("셀 생성 실패", row=cell["row"], col=cell["col"], error=str(exc))
                    completed += 1
                    return None

        tasks = [_generate_one(ti, c) for ti, c in all_cells]
        task_results = await asyncio.gather(*tasks)
        results = [r for r in task_results if r is not None]

        logger.info("전체 셀 생성 완료", total=total, success=len(results))
        return results

    # ------------------------------------------------------------------
    # 내부 유틸리티
    # ------------------------------------------------------------------

    @staticmethod
    def _parse_batch_response(content: str) -> list[CellFill]:
        """배치 응답(JSON)을 CellFill 목록으로 파싱한다."""
        fills: list[CellFill] = []

        # JSON 추출 (코드 블록 등 제거)
        text = content.strip()
        if "```" in text:
            parts = text.split("```")
            for part in parts:
                part = part.strip()
                if part.startswith("json"):
                    part = part[4:].strip()
                if part.startswith("{"):
                    text = part
                    break

        try:
            data = json.loads(text)
            cells = data.get("cells", [])
            for c in cells:
                fills.append(CellFill(
                    row=c["row"],
                    col=c["col"],
                    text=c.get("content", ""),
                ))
        except (json.JSONDecodeError, KeyError):
            logger.warning("배치 응답 JSON 파싱 실패", content=content[:200])

        return fills
