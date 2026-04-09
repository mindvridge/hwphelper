"""대화형 에이전트 — 채팅 인터페이스와 HWP 편집을 연결하는 핵심 모듈.

사용자의 자연어 명령을 이해하고, 적절한 도구를 호출하여 문서를 편집한다.
WebSocket을 통해 ChatEvent를 스트리밍한다.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from typing import Any, AsyncIterator

import structlog

from src.ai.image_generator import ImageGenerator
from src.hwp_engine.cell_classifier import CellClassifier
from src.hwp_engine.cell_writer import CellFill, CellWriter
from src.hwp_engine.document_manager import DocumentManager
from src.hwp_engine.field_manager import FieldManager
from src.hwp_engine.page_renderer import PageRenderer
from src.hwp_engine.schema_generator import SchemaGenerator
from src.hwp_engine.table_reader import TableReader
from src.hwp_engine.template_filler import TemplateFiller

from .llm_router import LLMResponse, LLMRouter, ToolCall
from .tool_definitions import DOCUMENT_MODIFYING_TOOLS, HWP_TOOLS
from .vision_reader import VisionReader
from .vision_reconciler import VisionReconciler

logger = structlog.get_logger()


# ------------------------------------------------------------------
# ChatEvent
# ------------------------------------------------------------------

@dataclass
class ChatEvent:
    """채팅 이벤트 — WebSocket으로 클라이언트에 전달."""

    type: str  # text_delta, tool_start, tool_result, document_updated, error, done
    data: Any = None

    def to_dict(self) -> dict[str, Any]:
        return {"type": self.type, "data": self.data}


# ------------------------------------------------------------------
# ChatAgent
# ------------------------------------------------------------------

SYSTEM_PROMPT = """당신은 한국 정부지원사업 계획서 작성 전문 컨설턴트입니다.
15년 이상의 정부과제 수주 경험을 보유하고 있으며, 예비창업패키지, TIPS, 창업성장기술개발,
데이터바우처 등 주요 정부지원사업의 계획서 작성과 평가 기준을 깊이 이해하고 있습니다.

사용자가 업로드한 HWP 문서의 표를 분석하고, 대화를 통해 각 셀에 전문적인 내용을 작성합니다.

## 작업 흐름 (반드시 이 순서를 따르세요)
1. analyze_document를 1번만 호출하여 표 구조를 파악합니다
2. 즉시 write_cell을 호출하여 빈 셀(EMPTY/PLACEHOLDER)에 내용을 작성합니다
3. read_table은 필요할 때만 호출합니다. 분석 반복은 금지입니다
4. 사용자 피드백을 반영하여 수정합니다
5. 완료 후 save_document로 저장합니다

## 중요: 분석 후 반드시 write_cell을 호출하세요
- analyze_document 결과에서 needs_fill=true인 셀을 찾으세요
- 해당 셀의 table_idx, row, col을 확인하고 write_cell로 내용을 작성하세요
- 한 번에 하나의 write_cell을 호출하세요
- 모든 빈 셀을 채울 때까지 반복하세요

## 작성 원칙
- **구체성**: 추상적 표현 대신 수치, 기간, 기술명을 명시합니다
  예: "매출 증가 예상" → "2027년까지 월 매출 5,000만원 달성 (전년 대비 150% 성장)"
- **논리 구조**: 문제 → 해결방안 → 기대효과 순서로 서술합니다
- **평가위원 관점**: 차별성, 실현가능성, 시장성, 성장성을 강조합니다
- **정량적 목표**: KPI, 매출 목표, 고용 계획 등 구체적 숫자를 포함합니다
- **간결한 문체**: 한 문장은 50자 이내, 불필요한 수식어를 제거합니다
- **셀 크기 적합**: 셀의 행/열 크기에 맞는 분량으로 작성합니다
  (제목 셀: 1~2줄, 내용 셀: 3~10줄, 개요 셀: 5~15줄)

## 항목별 작성 가이드
- **사업명**: [핵심기술] 기반 [목표시장] [솔루션유형] 개발 (20자 내외)
- **사업 개요**: 문제-솔루션-차별점-시장규모-목표를 5줄 이내로
- **기술 설명**: 핵심기술 → 구현방식 → 기존 대비 우위점 → TRL 단계
- **시장 분석**: TAM → SAM → SOM 순, 출처 명시, 성장률 포함
- **사업화 전략**: 타겟 고객 → 채널 → 가격 → BM → 매출 로드맵
- **개발 일정**: 단계별 마일스톤, 월 단위, 담당자 명시
- **예산**: 인건비/재료비/외주비 등 세목별 산출근거 포함
- **대표자 이력**: 학력-경력-성과를 역순으로, 핵심 실적 강조

## 금지사항
- 마크다운 문법(**, ##, ```, - 등)을 셀 내용에 사용하지 않습니다
- HTML 태그를 사용하지 않습니다
- "~할 것입니다", "~하겠습니다" 같은 의지형 종결보다 "~한다", "~이다" 체로 작성합니다
- 모호한 표현("다양한", "혁신적인", "효율적인")을 단독으로 쓰지 않습니다
- 셀 내용만 반환하고, 따옴표나 부연 설명을 추가하지 않습니다

## 도구 사용법
- analyze_document: 문서를 열어 표 구조 파악
- read_table: 특정 표의 전체 셀 내용 읽기
- write_cell: 셀에 텍스트 삽입 (서식 자동 보존)
- fill_field: 누름틀 필드에 텍스트 삽입
- validate_format: 서식 규정 검증
- undo: 마지막 작업 되돌리기
- save_document: 문서 저장
- generate_image: AI 이미지 생성

## 대화 스타일
- 한국어로 전문적이면서 친근하게 대화합니다
- 작업 전: "표 1의 '사업 개요' 셀을 작성하겠습니다"처럼 무엇을 할지 먼저 알립니다
- 작업 후: 작성한 내용을 요약하고, 수정 제안이 있으면 함께 안내합니다
- 사용자가 기업 정보를 제공하면 해당 정보를 모든 셀 작성에 일관되게 반영합니다"""

# 자동채우기 파이프라인 전용 프롬프트 — 인사말/설명 없이 값만 반환
AUTOFILL_PROMPT = """당신은 정부사업 계획서 셀 값 생성기입니다.

## 절대 규칙
- 요청된 셀 값만 출력하세요
- 인사말, 자기소개, "안녕하세요", "컨설턴트입니다" 등 절대 금지
- "작성해 드리겠습니다", "분석하겠습니다" 등 메타 설명 절대 금지
- 마크다운(**, ##, ```) 절대 금지
- 따옴표로 감싸지 마세요
- "~한다", "~이다" 체로 작성
- 구체적 수치, 기간, 기술명을 포함

## 출력 형식
셀에 들어갈 텍스트만 출력합니다. 다른 어떤 것도 출력하지 않습니다."""


class ChatAgent:
    """채팅 인터페이스와 HWP 편집을 연결하는 에이전트."""

    def __init__(
        self,
        llm_router: LLMRouter,
        doc_manager: DocumentManager,
        com_executor: Any = None,
    ) -> None:
        self.llm = llm_router
        self.doc_mgr = doc_manager
        self.image_gen = ImageGenerator()
        self.com_executor = com_executor
        self.conversation_history: dict[str, list[dict[str, Any]]] = {}
        self._filling_in_progress: set[str] = set()  # 중복 실행 방지

    # ------------------------------------------------------------------
    # 대화 관리
    # ------------------------------------------------------------------

    def get_history(self, session_id: str) -> list[dict[str, Any]]:
        """세션의 대화 히스토리를 반환한다."""
        return self.conversation_history.setdefault(session_id, [])

    def clear_history(self, session_id: str) -> None:
        """대화 히스토리를 초기화한다."""
        self.conversation_history.pop(session_id, None)

    @staticmethod
    def _parse_image_message(raw: str) -> tuple[str, str]:
        """[IMAGE:data:...]\\n텍스트 형식을 파싱한다."""
        if raw.startswith("[IMAGE:") and "]\n" in raw:
            end = raw.index("]\n")
            image_url = raw[7:end]
            text = raw[end + 2:].strip()
            return image_url, text
        elif raw.startswith("[IMAGE:") and raw.endswith("]"):
            return raw[7:-1], ""
        return "", raw

    # ------------------------------------------------------------------
    # 메시지 처리
    # ------------------------------------------------------------------

    async def process_message(
        self,
        session_id: str,
        user_message: str,
        model_id: str | None = None,
    ) -> AsyncIterator[ChatEvent]:
        """사용자 메시지를 처리하고 ChatEvent를 스트리밍한다.

        Yields
        ------
        ChatEvent
            - text_delta: 텍스트 토큰
            - tool_start: 도구 호출 시작
            - tool_result: 도구 실행 결과
            - document_updated: 문서 수정됨
            - error: 에러
            - done: 응답 완료
        """
        history = self.get_history(session_id)

        # 이미지 첨부 감지 → 비전 메시지 형식으로 변환
        if user_message.startswith("[IMAGE:"):
            image_part, text = self._parse_image_message(user_message)
            history.append({
                "role": "user",
                "content": [
                    {"type": "image_url", "image_url": {"url": image_part}},
                    {"type": "text", "text": text or "이 이미지를 분석해주세요."},
                ],
            })
        else:
            history.append({"role": "user", "content": user_message})

        try:
            # "채워줘", "작성해줘" 등 자동 채우기 요청 감지
            fill_keywords = ["채워", "작성해", "빈 셀", "빈셀", "내용을 넣", "자동으로"]
            is_fill_request = any(kw in user_message for kw in fill_keywords)

            if is_fill_request and session_id:
                # 직접 파이프라인: 분석 → LLM 생성 → COM 쓰기
                async for event in self._auto_fill_pipeline(session_id, user_message, model_id):
                    yield event
            else:
                # 일반 대화 (도구 없이)
                messages = [{"role": "system", "content": SYSTEM_PROMPT}] + history
                response = await self.llm.chat(messages=messages, model_id=model_id)
                if not isinstance(response, LLMResponse):
                    raise TypeError(f"Expected LLMResponse, got {type(response)}")
                if response.content:
                    yield ChatEvent(type="text_delta", data=response.content)
                    history.append({"role": "assistant", "content": response.content})

            yield ChatEvent(type="done", data={"model": model_id or self.llm.default_model})

        except Exception as e:
            logger.exception("메시지 처리 오류", session_id=session_id)
            yield ChatEvent(type="error", data={"message": str(e)})

    # ------------------------------------------------------------------
    # 자동 채우기 파이프라인
    # ------------------------------------------------------------------

    async def _auto_fill_pipeline(
        self, session_id: str, user_message: str, model_id: str | None,
    ) -> AsyncIterator[ChatEvent]:
        """양식 구조를 인식하여 적절한 위치에 내용을 삽입한다."""
        if session_id in self._filling_in_progress:
            yield ChatEvent(type="text_delta", data="이미 자동 채우기가 진행 중입니다.")
            return
        self._filling_in_progress.add(session_id)

        try:
            session = self.doc_mgr.get_session(session_id)
        except KeyError:
            self._filling_in_progress.discard(session_id)
            yield ChatEvent(type="text_delta", data="HWP 파일을 먼저 업로드해주세요.")
            return

        working_path = session.working_path

        # 업로드 시 생성된 COM 세션을 완전히 닫아야 새 COM에서 파일을 열 수 있음
        try:
            await self._run_com(lambda: session.hwp_ctrl.quit())
            logger.info("업로드 COM 세션 종료", session_id=session_id)
        except Exception:
            pass

        # 1단계: 양식 구조 분석 (COM + 비전)
        yield ChatEvent(type="tool_start", data={"tool": "analyze_template", "args": {}})

        def _analyze():
            # COM 스레드에서 새로 문서를 열어서 분석+쓰기
            from src.hwp_engine.com_controller import HwpController
            ctrl = HwpController(visible=True)
            ctrl.connect()
            ctrl.open(working_path)
            hwp_local = ctrl.hwp

            filler = TemplateFiller(hwp_local)
            structure = filler.analyze_template()

            # 페이지 이미지 렌더링 (비전용)
            try:
                renderer = PageRenderer(ctrl)
                page_images = renderer.render_all_pages()
            except Exception:
                page_images = []

            return structure, filler, page_images, ctrl, hwp_local

        structure, filler, page_images, com_ctrl, hwp = await self._run_com(lambda: _analyze())

        # 비전 보강: 별도 스레드에서 동기 실행 (async generator 내 await 교착 방지)
        vision_context = ""
        if page_images:
            yield ChatEvent(type="text_delta", data="📷 문서 이미지 분석 중...\n")
            try:
                import time as _time
                vision_reader = VisionReader(self.llm)
                logger.info("비전 분석 시작", model=vision_reader._model_id)
                t_start = _time.time()
                # 동기 호출을 스레드에서 실행 (async generator 교착 방지)
                vision_result = await self._run_com(
                    lambda: vision_reader._sync_read_page(page_images[0])
                )
                elapsed = _time.time() - t_start
                table_count = len(vision_result.tables)
                logger.info("비전 분석 완료", tables=table_count, elapsed=f"{elapsed:.1f}s")
                if table_count > 0:
                    vision_context = f" (비전 인식: {table_count}개 표, {elapsed:.1f}초)"
                    yield ChatEvent(type="text_delta", data=f"📷 문서 이미지 분석 완료{vision_context}\n")
            except Exception as exc:
                logger.warning("비전 분석 실패", error=str(exc), type=type(exc).__name__)

        items = filler.get_fillable_summary(structure)

        info_count = len([i for i in items if i["type"] == "info"])
        body_count = len([i for i in items if i["type"] == "body_section"])
        narrative_count = len([i for i in items if i["type"] == "narrative"])
        data_table_count = len([i for i in items if i["type"] == "data_table"])

        yield ChatEvent(type="tool_result", data={
            "tool": "analyze_template",
            "result": {
                "info_fields": info_count,
                "body_sections": body_count,
                "narrative_sections": narrative_count,
                "data_tables": data_table_count,
                "total": len(items),
            },
        })

        if not items:
            self._filling_in_progress.discard(session_id)
            await self._run_com(lambda: com_ctrl.quit())
            yield ChatEvent(type="text_delta", data="채울 항목을 찾지 못했습니다.")
            return

        total = len(items)
        parts = []
        if info_count:
            parts.append(f"기업정보 {info_count}개")
        if body_count:
            parts.append(f"본문섹션 {body_count}개")
        if narrative_count:
            parts.append(f"서술항목 {narrative_count}개")
        if data_table_count:
            parts.append(f"데이터표 {data_table_count}개")
        yield ChatEvent(type="text_delta", data=f"{', '.join(parts)} 발견. 작성을 시작합니다...\n\n")

        # 2단계: 항목별 LLM 생성 → 적절한 위치에 쓰기
        wrote = 0
        for idx, item in enumerate(items):
            yield ChatEvent(type="progress", data={
                "current": idx + 1, "total": total,
                "description": f"{item.get('label', '') or item.get('title', '')} 작성 중...",
            })

            try:
                if item["type"] == "info":
                    # 기업정보: 짧은 값 생성
                    prompt = (
                        f"정부사업 신청서의 '{item['label']}' 항목에 들어갈 값을 작성하세요.\n"
                        f"현재 예시값: {item['example']}\n"
                        f"사용자 정보: {user_message}\n"
                        f"\n값만 작성하세요. 따옴표, 설명 금지. 예시와 같은 형식으로."
                    )
                    use_vision_info = bool(page_images) and self._model_supports_vision(model_id)
                    if use_vision_info:
                        user_content_info = self._build_vision_prompt(prompt, page_images, max_pages=1)
                    else:
                        user_content_info = prompt
                    resp = await self.llm.chat(
                        messages=[{"role": "system", "content": AUTOFILL_PROMPT}, {"role": "user", "content": user_content_info}],
                        model_id=model_id,
                    )
                    if not isinstance(resp, LLMResponse):
                        continue
                    value = resp.content.strip().strip('"').strip("'")
                    if value and len(value) > 1:
                        await self._run_com(lambda t=item["table_idx"], i=item["index"], v=value: filler.fill_info_field(t, i, v))
                        wrote += 1
                        yield ChatEvent(type="tool_result", data={"tool": "fill_info", "result": {"label": item["label"], "value": value}})

                elif item["type"] == "body_section":
                    # 본문 섹션: ◦/- 마커 유무에 따라 다른 전략
                    # 비전 모델이면 페이지 이미지 첨부하여 정확도 향상
                    use_vision = bool(page_images) and self._model_supports_vision(model_id)
                    markers = item.get("markers", [])
                    marker_count = len(markers)

                    from src.hwp_engine.template_filler import BodySection
                    bs = BodySection(
                        section_num=item["section_num"],
                        title=item["title"],
                        guide_text=item.get("guide", ""),
                        markers=markers,
                        table_idx=item["table_idx"],
                    )

                    if marker_count > 0:
                        # 마커 있음: ◦/- 구조에 맞춰 내용 생성
                        marker_desc = []
                        for m in markers:
                            marker_desc.append(f"  {m['type']}")
                        marker_structure = "\n".join(marker_desc)

                        prompt = (
                            f"정부사업 계획서의 '{item['title']}' 섹션을 작성하세요.\n"
                        )
                        if item.get("guide"):
                            prompt += f"평가기준: {item['guide']}\n"
                        prompt += f"\n사용자 정보: {user_message}\n"
                        prompt += (
                            f"\n원본 양식에 {marker_count}개의 빈 항목이 있습니다.\n"
                            f"구조:\n{marker_structure}\n"
                            f"\n정확히 {marker_count}개의 항목 내용을 작성하세요.\n"
                            f"각 항목은 줄바꿈(\\n)으로 구분하세요.\n"
                            f"◦ 항목은 소제목(핵심 키워드 1~2줄), - 항목은 세부 설명(2~4줄)입니다.\n"
                            f"◦, - 기호는 포함하지 마세요. 내용만 작성하세요.\n"
                            f"※ 안내문, 마크다운, 따옴표 금지."
                        )
                        # 비전 모델이면 페이지 이미지 첨부
                        if use_vision:
                            user_content = self._build_vision_prompt(prompt, page_images, max_pages=1)
                        else:
                            user_content = prompt
                        resp = await self.llm.chat(
                            messages=[{"role": "system", "content": AUTOFILL_PROMPT}, {"role": "user", "content": user_content}],
                            model_id=model_id,
                        )
                        if not isinstance(resp, LLMResponse):
                            continue
                        raw = resp.content.strip().strip('"')
                        lines = [l.strip() for l in raw.split("\n") if l.strip()]
                        contents_list = []
                        for l in lines:
                            cleaned = re.sub(r"^[◦○\-]\s*", "", l).strip()
                            if cleaned:
                                contents_list.append(cleaned)

                        if contents_list:
                            await self._run_com(
                                lambda t=item["table_idx"], s=bs, c=contents_list: filler.fill_body_section(t, s, c)
                            )
                            wrote += 1
                            yield ChatEvent(type="tool_result", data={
                                "tool": "fill_body_section",
                                "result": {"title": item["title"], "filled": len(contents_list)},
                            })
                    else:
                        # 마커 없음: ※ 안내문 삭제 후 서술형으로 내용 작성
                        prompt = (
                            f"정부사업 계획서의 '{item['title']}' 섹션을 작성하세요.\n"
                        )
                        if item.get("guide"):
                            prompt += f"평가기준: {item['guide']}\n"
                        prompt += f"\n사용자 정보: {user_message}\n"
                        prompt += (
                            f"\n다음 형식으로 작성하세요:\n"
                            f"◦ [소제목1]\n"
                            f"  - [세부 설명 2~4줄]\n"
                            f"◦ [소제목2]\n"
                            f"  - [세부 설명 2~4줄]\n"
                            f"\n3~5개의 소제목(◦)과 각각의 세부 설명(-)을 작성하세요.\n"
                            f"※ 안내문은 포함하지 마세요. 마크다운 금지."
                        )
                        if use_vision:
                            user_content = self._build_vision_prompt(prompt, page_images, max_pages=1)
                        else:
                            user_content = prompt
                        resp = await self.llm.chat(
                            messages=[{"role": "system", "content": AUTOFILL_PROMPT}, {"role": "user", "content": user_content}],
                            model_id=model_id,
                        )
                        if not isinstance(resp, LLMResponse):
                            continue
                        content = resp.content.strip().strip('"')
                        if content and len(content) > 10:
                            await self._run_com(
                                lambda t=item["table_idx"], s=bs, c=content: filler.fill_body_narrative(t, s, c)
                            )
                            wrote += 1
                            yield ChatEvent(type="tool_result", data={
                                "tool": "fill_body_narrative",
                                "result": {"title": item["title"]},
                            })
                        yield ChatEvent(type="document_updated", data={"tool": "fill_body_section"})

                elif item["type"] == "narrative":
                    # 서술형 (하위 호환): 제목 유지 + 내용 생성
                    sub_labels = ", ".join(item.get("sub_items", []))
                    prompt = (
                        f"정부사업 계획서의 '{item['title']}' 섹션을 작성하세요.\n"
                    )
                    if sub_labels:
                        prompt += f"하위 항목: {sub_labels}\n"
                    if item.get("guide"):
                        prompt += f"안내문: {item['guide']}\n"
                    prompt += f"\n사용자 정보: {user_message}\n"
                    prompt += (
                        f"\n다음 형식으로 작성하세요:\n"
                        f"{item['title']}\n"
                        f"  1) [구체적인 내용 3~5줄]\n"
                        f"  2) [구체적인 내용 3~5줄]\n"
                        f"\n※ 안내문은 포함하지 마세요. 마크다운 금지."
                    )
                    use_vision_n = bool(page_images) and self._model_supports_vision(model_id)
                    if use_vision_n:
                        user_content_n = self._build_vision_prompt(prompt, page_images, max_pages=1)
                    else:
                        user_content_n = prompt
                    resp = await self.llm.chat(
                        messages=[{"role": "system", "content": AUTOFILL_PROMPT}, {"role": "user", "content": user_content_n}],
                        model_id=model_id,
                    )
                    if not isinstance(resp, LLMResponse):
                        continue
                    content = resp.content.strip().strip('"')
                    if content and len(content) > 10:
                        await self._run_com(lambda t=item["table_idx"], c=content: filler.fill_narrative(t, c))
                        wrote += 1
                        yield ChatEvent(type="tool_result", data={"tool": "fill_narrative", "result": {"title": item["title"]}})
                        yield ChatEvent(type="document_updated", data={"tool": "fill_narrative"})

                elif item["type"] == "data_table":
                    # 데이터 표 (예산, 일정 등): 예시 데이터를 사용자 정보로 교체
                    empty_cells = item.get("empty_cells", [])
                    headers = item.get("headers", [])
                    if not empty_cells:
                        continue

                    # 표 구조와 예시 데이터 설명
                    header_desc = " | ".join(headers)
                    prompt = (
                        f"정부사업 계획서의 '{item['title']}' 표를 작성하세요.\n"
                        f"표 열: {header_desc}\n"
                        f"\n현재 예시 데이터를 사용자 정보에 맞게 교체하세요.\n"
                        f"\n사용자 정보: {user_message}\n"
                        f"\n교체할 셀 목록 (행, 열이름, 현재 예시값):\n"
                    )
                    for ec in empty_cells:
                        example = ec.get("example", "")
                        prompt += f"  - 행{ec['row']}, '{ec['header']}': {example}\n"
                    prompt += (
                        f"\n정확히 {len(empty_cells)}개의 값을 줄바꿈으로 구분하여 작성하세요.\n"
                        f"순서대로 각 셀에 들어갈 값만 작성하세요.\n"
                        f"예산이면 산출근거(•항목명(수량×단가))와 금액을,\n"
                        f"일정이면 구체적 내용과 기간을 작성하세요.\n"
                        f"예시와 동일한 형식으로 작성하세요.\n"
                        f"따옴표, 마크다운, 설명 금지. 값만 작성."
                    )
                    use_vision_dt = bool(page_images) and self._model_supports_vision(model_id)
                    if use_vision_dt:
                        user_content_dt = self._build_vision_prompt(prompt, page_images, max_pages=1)
                    else:
                        user_content_dt = prompt
                    resp = await self.llm.chat(
                        messages=[{"role": "system", "content": AUTOFILL_PROMPT}, {"role": "user", "content": user_content_dt}],
                        model_id=model_id,
                    )
                    if not isinstance(resp, LLMResponse):
                        continue
                    raw = resp.content.strip().strip('"')
                    values = [l.strip() for l in raw.split("\n") if l.strip()]

                    filled_count = 0
                    for vi, ec in enumerate(empty_cells):
                        if vi >= len(values):
                            break
                        val = values[vi]
                        if val:
                            await self._run_com(
                                lambda t=item["table_idx"], r=ec["row"], c=ec["col"], v=val: filler.fill_data_cell(t, r, c, v)
                            )
                            filled_count += 1

                    if filled_count:
                        wrote += 1
                        yield ChatEvent(type="tool_result", data={
                            "tool": "fill_data_table",
                            "result": {"title": item["title"], "filled": filled_count, "total": len(empty_cells)},
                        })
                        yield ChatEvent(type="document_updated", data={"tool": "fill_data_table"})

            except Exception as e:
                logger.warning("항목 작성 실패", item=item.get("title", item.get("label", "")), error=str(e))

            # API rate limit 방지: 항목 간 2초 딜레이
            import asyncio as _aio_delay
            await _aio_delay.sleep(2.0)

        # 3단계: 완료 — COM 저장 및 정리, 세션 COM 재연결
        try:
            await self._run_com(lambda: com_ctrl.save())
            logger.info("자동채우기 결과 저장 완료")
        except Exception as exc:
            logger.warning("저장 실패", error=str(exc))
        try:
            await self._run_com(lambda: com_ctrl.quit())
        except Exception:
            pass

        # 세션 COM을 다시 열어서 이후 작업(다운로드 등) 가능하게
        try:
            def _reopen():
                from src.hwp_engine.com_controller import HwpController
                new_ctrl = HwpController(visible=True)
                new_ctrl.connect()
                new_ctrl.open(working_path)
                session.hwp_ctrl = new_ctrl
            await self._run_com(_reopen)
            logger.info("세션 COM 재연결 완료")
        except Exception as exc:
            logger.warning("세션 COM 재연결 실패", error=str(exc))

        self._filling_in_progress.discard(session_id)
        summary = f"\n\n{wrote}/{total}개 항목 작성 완료. 한/글에서 결과를 확인하세요.\n수정이 필요하면 말씀해주세요."
        yield ChatEvent(type="text_delta", data=summary)
        self.get_history(session_id).append({"role": "assistant", "content": summary})

    # ------------------------------------------------------------------
    # 도구 실행
    # ------------------------------------------------------------------

    def _model_supports_vision(self, model_id: str | None = None) -> bool:
        """모델이 비전(이미지 인식)을 지원하는지 확인한다."""
        mid = model_id or self.llm._default_model
        cfg = self.llm._models.get(mid, {})
        return bool(cfg.get("supports_vision", False))

    def _build_vision_prompt(
        self,
        text_prompt: str,
        page_images: list[Any],
        max_pages: int = 1,
    ) -> list[dict[str, Any]]:
        """텍스트 프롬프트에 페이지 이미지를 결합한 멀티모달 메시지를 만든다.

        비전 모델이 문서 이미지를 보면서 더 정확한 콘텐츠를 생성할 수 있다.
        """
        content: list[dict[str, Any]] = []

        # 페이지 이미지 추가 (최대 max_pages장)
        for page in page_images[:max_pages]:
            content.append({
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": page.mime_type,
                    "data": page.base64_data,
                },
            })

        # 텍스트 프롬프트
        content.append({"type": "text", "text": text_prompt})

        return content

    # 전용 COM 스레드 — 항상 같은 스레드에서 실행되어야 COM 객체 접근 가능
    _com_thread_executor: Any = None

    def _get_com_executor(self) -> Any:
        """파이프라인 전용 COM 스레드풀을 반환한다 (항상 같은 스레드 보장)."""
        import concurrent.futures
        if self._com_thread_executor is None:
            self._com_thread_executor = concurrent.futures.ThreadPoolExecutor(
                max_workers=1, thread_name_prefix="com-pipeline"
            )
        return self._com_thread_executor

    async def _run_com(self, func: Any, *args: Any) -> Any:
        """COM 작업을 전용 단일 스레드에서 실행한다."""
        import asyncio
        loop = asyncio.get_running_loop()

        def _wrapped():
            import pythoncom
            pythoncom.CoInitialize()
            try:
                return func(*args)
            finally:
                pass

        executor = self._get_com_executor()
        return await loop.run_in_executor(executor, _wrapped)

    async def execute_tool(self, session_id: str, tool_call: ToolCall) -> dict[str, Any]:
        """도구를 실행한다."""
        args = tool_call.arguments
        name = tool_call.name

        # 이미지 생성은 세션 없이도 가능
        if name == "generate_image":
            return await self._tool_generate_image(args)

        try:
            session = self.doc_mgr.get_session(session_id)
        except KeyError:
            return {"error": f"세션을 찾을 수 없습니다: {session_id}"}

        ctrl = session.hwp_ctrl

        # COM 도구는 전용 스레드에서 실행
        try:
            if name == "analyze_document":
                return await self._run_com(self._tool_analyze_sync, ctrl, session)

            elif name == "read_table":
                return await self._run_com(self._tool_read_table_sync, ctrl, args)

            elif name == "read_cell":
                return await self._run_com(self._tool_read_cell_sync, ctrl, args)

            elif name == "write_cell":
                return await self._run_com(self._tool_write_cell_sync, ctrl, args)

            elif name == "fill_field":
                return await self._run_com(self._tool_fill_field_sync, ctrl, args)

            elif name == "fill_all_empty_cells":
                return await self._run_com(self._tool_fill_all_sync, ctrl, session, args)

            elif name == "validate_format":
                return await self._run_com(self._tool_validate_sync, ctrl, args)

            elif name == "undo":
                ok = self.doc_mgr.undo(session_id)
                return {"success": ok, "message": "되돌리기 완료" if ok else "되돌릴 수 없습니다"}

            elif name == "save_document":
                fmt = args.get("format", "hwp")
                ctrl.save()
                return {"success": True, "format": fmt, "path": ctrl.file_path}

            elif name == "get_document_info":
                return await self._tool_document_info(session)

            else:
                return {"error": f"알 수 없는 도구: {name}"}

        except Exception as e:
            logger.warning("도구 실행 오류", tool=name, error=str(e))
            return {"error": str(e)}

    # ------------------------------------------------------------------
    # 도구 구현
    # ------------------------------------------------------------------

    # sync 버전 (COM 스레드에서 실행)
    def _tool_analyze_sync(self, ctrl: Any, session: Any) -> dict:
        reader = TableReader(ctrl)
        classifier = CellClassifier()
        generator = SchemaGenerator()

        tables = reader.read_all_tables()
        for t in tables:
            classifier.classify_table(t)

        # 비전 페이지 이미지 캐시 (비동기 파이프라인에서 사용)
        try:
            renderer = PageRenderer(ctrl)
            pages = renderer.render_all_pages()
            session.page_images = pages
            logger.info("페이지 이미지 렌더링 완료", pages=len(pages))
        except Exception as exc:
            session.page_images = []
            logger.debug("페이지 렌더링 스킵", error=str(exc))

        schema = generator.generate(tables, ctrl.file_path or "")
        session.schema = schema

        return {
            "total_tables": schema["total_tables"],
            "total_cells_to_fill": schema["total_cells_to_fill"],
            "page_images": len(getattr(session, "page_images", [])),
            "tables_summary": [
                {
                    "table_idx": t["table_idx"],
                    "rows": t["rows"],
                    "cols": t["cols"],
                    "cells_to_fill": t["cells_to_fill"],
                }
                for t in schema["tables"]
            ],
        }

    def _tool_read_table_sync(self, ctrl: Any, args: dict) -> dict:
        reader = TableReader(ctrl)
        classifier = CellClassifier()
        table = reader.read_table(args["table_idx"])
        classifier.classify_table(table)
        return table.to_dict()

    def _tool_read_cell_sync(self, ctrl: Any, args: dict) -> dict:
        reader = TableReader(ctrl)
        style = reader.read_cell_style(args["table_idx"], args["row"], args["col"])
        table = reader.read_table(args["table_idx"])
        cell = table.get_cell(args["row"], args["col"])
        if cell:
            return {
                "text": cell.text,
                "style": {
                    "font_name": style.font_name,
                    "font_size": style.font_size,
                    "bold": style.bold,
                },
            }
        return {"error": "셀을 찾을 수 없습니다"}

    def _tool_write_cell_sync(self, ctrl: Any, args: dict) -> dict:
        writer = CellWriter(ctrl)
        writer.write_cell(args["table_idx"], args["row"], args["col"], args["text"])
        return {"success": True, "row": args["row"], "col": args["col"]}

    def _tool_fill_field_sync(self, ctrl: Any, args: dict) -> dict:
        fm = FieldManager(ctrl)
        ok = fm.fill_field(args["field_name"], args["text"])
        return {"success": ok, "field": args["field_name"]}

    def _tool_fill_all_sync(self, ctrl: Any, session: Any, args: dict) -> dict:
        # 간단한 구현 — 실제로는 CellGenerator를 사용해야 함
        reader = TableReader(ctrl)
        classifier = CellClassifier()
        tables = reader.read_all_tables()
        filled = 0
        for t in tables:
            classifier.classify_table(t)
            for cell in t.empty_cells():
                filled += 1
        return {
            "message": f"{filled}개 빈 셀 감지됨. CellGenerator로 내용 생성 필요.",
            "cells_to_fill": filled,
            "program_name": args.get("program_name", ""),
        }

    def _tool_validate_sync(self, ctrl: Any, args: dict) -> dict:
        # 기본 구현
        return {
            "program_name": args.get("program_name", "기본"),
            "issues": [],
            "message": "서식 검증 기능은 향후 구현 예정입니다.",
        }

    async def _tool_document_info(self, session: Any) -> dict:
        ctrl = session.hwp_ctrl
        reader = TableReader(ctrl)
        table_count = reader.get_table_count()
        history = self.doc_mgr.get_history(session.session_id)
        return {
            "file_path": ctrl.file_path,
            "table_count": table_count,
            "snapshot_count": len(session.snapshots),
            "history": [{"index": h.index, "description": h.description} for h in history],
        }

    async def _tool_generate_image(self, args: dict) -> dict:
        prompt = args.get("prompt", "")
        size = args.get("size", "1024x1024")
        result = await self.image_gen.generate(prompt=prompt, size=size)
        if result.error:
            return {"error": result.error}
        return {
            "success": True,
            "url": result.url,
            "base64": result.base64_data[:50] + "..." if result.base64_data else "",
            "prompt": prompt,
            "image_url": result.url,  # 프론트엔드에서 표시용
        }

    # ------------------------------------------------------------------
    # 메시지 빌더 (프로바이더별 형식)
    # ------------------------------------------------------------------

    def _build_assistant_message(self, response: LLMResponse, model_id: str | None) -> dict:
        """어시스턴트 응답 메시지를 히스토리 형식으로 구성."""
        mid = model_id or self.llm.default_model
        cfg = self.llm._models.get(mid, {})
        provider = cfg.get("provider", "openai")

        if provider == "anthropic":
            content: list[dict] = []
            if response.content:
                content.append({"type": "text", "text": response.content})
            for tc in response.tool_calls:
                content.append({"type": "tool_use", "id": tc.id, "name": tc.name, "input": tc.arguments})
            return {"role": "assistant", "content": content}
        else:
            msg: dict[str, Any] = {"role": "assistant", "content": response.content or None}
            if response.tool_calls:
                msg["tool_calls"] = [
                    {
                        "id": tc.id,
                        "type": "function",
                        "function": {"name": tc.name, "arguments": json.dumps(tc.arguments, ensure_ascii=False)},
                    }
                    for tc in response.tool_calls
                ]
            return msg

    def _build_tool_result_message(self, tc: ToolCall, result: dict, model_id: str | None) -> dict:
        """도구 실행 결과 메시지를 히스토리 형식으로 구성."""
        mid = model_id or self.llm.default_model
        cfg = self.llm._models.get(mid, {})
        provider = cfg.get("provider", "openai")
        result_str = json.dumps(result, ensure_ascii=False)

        if provider == "anthropic":
            return {
                "role": "user",
                "content": [{"type": "tool_result", "tool_use_id": tc.id, "content": result_str}],
            }
        else:
            return {
                "role": "tool",
                "tool_call_id": tc.id,
                "content": result_str,
            }
