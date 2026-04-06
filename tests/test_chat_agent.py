"""ChatAgent 테스트."""

from __future__ import annotations

from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from src.ai.chat_agent import ChatAgent, ChatEvent
from src.ai.llm_router import LLMResponse, LLMRouter, ToolCall
from src.hwp_engine.document_manager import DocumentManager


@pytest.fixture
def mock_router() -> LLMRouter:
    router = MagicMock(spec=LLMRouter)
    router.default_model = "test-model"
    router._models = {"test-model": {"provider": "openai", "model": "test"}}
    return router


@pytest.fixture
def mock_doc_mgr(tmp_path) -> DocumentManager:
    mgr = MagicMock(spec=DocumentManager)
    return mgr


@pytest.fixture
def agent(mock_router, mock_doc_mgr) -> ChatAgent:
    return ChatAgent(llm_router=mock_router, doc_manager=mock_doc_mgr)


class TestChatAgentHistory:
    """대화 히스토리 관리 테스트."""

    def test_get_history_empty(self, agent: ChatAgent) -> None:
        history = agent.get_history("session1")
        assert history == []

    def test_get_history_creates_session(self, agent: ChatAgent) -> None:
        h = agent.get_history("new_session")
        assert isinstance(h, list)

    def test_clear_history(self, agent: ChatAgent) -> None:
        agent.get_history("s1").append({"role": "user", "content": "hi"})
        agent.clear_history("s1")
        assert agent.get_history("s1") == []


class TestChatEventSerialization:
    """ChatEvent 직렬화 테스트."""

    def test_text_delta(self) -> None:
        ev = ChatEvent(type="text_delta", data="안녕하세요")
        d = ev.to_dict()
        assert d["type"] == "text_delta"
        assert d["data"] == "안녕하세요"

    def test_tool_start(self) -> None:
        ev = ChatEvent(type="tool_start", data={"tool": "read_table", "args": {"table_idx": 0}})
        d = ev.to_dict()
        assert d["data"]["tool"] == "read_table"

    def test_done(self) -> None:
        ev = ChatEvent(type="done", data={"model": "claude-sonnet"})
        assert ev.type == "done"


class TestProcessMessage:
    """메시지 처리 테스트."""

    @pytest.mark.asyncio
    async def test_simple_text_response(self, agent: ChatAgent, mock_router) -> None:
        """도구 호출 없는 단순 텍스트 응답."""
        mock_router.chat = AsyncMock(return_value=LLMResponse(
            content="안녕하세요! 무엇을 도와드릴까요?",
            tool_calls=[],
            model="test",
        ))

        events = []
        async for ev in agent.process_message("s1", "안녕"):
            events.append(ev)

        types = [e.type for e in events]
        assert "text_delta" in types
        assert "done" in types
        text_ev = next(e for e in events if e.type == "text_delta")
        assert "안녕하세요" in text_ev.data

    @pytest.mark.asyncio
    async def test_tool_call_response(self, agent: ChatAgent, mock_router, mock_doc_mgr) -> None:
        """도구 호출 포함 응답."""
        # 첫 호출: 도구 호출
        tool_response = LLMResponse(
            content="",
            tool_calls=[ToolCall(id="tc1", name="get_document_info", arguments={})],
            model="test",
        )
        # 두 번째 호출: 텍스트 응답
        text_response = LLMResponse(
            content="문서에 표가 3개 있습니다.",
            tool_calls=[],
            model="test",
        )
        mock_router.chat = AsyncMock(side_effect=[tool_response, text_response])

        # 세션 mock
        mock_session = MagicMock()
        mock_session.session_id = "s1"
        mock_session.hwp_ctrl = MagicMock()
        mock_session.hwp_ctrl.file_path = "test.hwp"
        mock_session.snapshots = []
        mock_doc_mgr.get_session.return_value = mock_session

        # TableReader mock
        with patch("src.ai.chat_agent.TableReader") as mock_reader_cls:
            mock_reader = MagicMock()
            mock_reader.get_table_count.return_value = 3
            mock_reader_cls.return_value = mock_reader

            events = []
            async for ev in agent.process_message("s1", "문서 정보 알려줘"):
                events.append(ev)

        types = [e.type for e in events]
        assert "tool_start" in types
        assert "tool_result" in types
        assert "done" in types

    @pytest.mark.asyncio
    async def test_error_handling(self, agent: ChatAgent, mock_router) -> None:
        """LLM 호출 오류 처리."""
        mock_router.chat = AsyncMock(side_effect=Exception("API 오류"))

        events = []
        async for ev in agent.process_message("s1", "테스트"):
            events.append(ev)

        types = [e.type for e in events]
        assert "error" in types

    @pytest.mark.asyncio
    async def test_history_persistence(self, agent: ChatAgent, mock_router) -> None:
        """대화 히스토리가 올바르게 유지되는지 확인."""
        mock_router.chat = AsyncMock(return_value=LLMResponse(content="응답1", tool_calls=[], model="test"))

        async for _ in agent.process_message("s1", "메시지1"):
            pass

        history = agent.get_history("s1")
        assert len(history) == 2  # user + assistant
        assert history[0]["role"] == "user"
        assert history[0]["content"] == "메시지1"
        assert history[1]["role"] == "assistant"
