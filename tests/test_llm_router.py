"""LLM 라우터 테스트.

실제 LLM 호출은 mock 처리.
"""

from __future__ import annotations

import json
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from src.ai.llm_router import LLMResponse, LLMRouter, ModelInfo, TokenUsage, ToolCall


@pytest.fixture
def config_file(tmp_path):
    """테스트용 LLM 설정 파일."""
    config = {
        "models": {
            "test-claude": {
                "provider": "anthropic",
                "model": "claude-sonnet-4-20250514",
                "api_key_env": "TEST_ANTHROPIC_KEY",
                "max_tokens": 4096,
                "description": "테스트 Claude",
            },
            "test-gpt": {
                "provider": "openai",
                "model": "gpt-4o",
                "api_key_env": "TEST_OPENAI_KEY",
                "max_tokens": 4096,
                "description": "테스트 GPT",
            },
            "test-local": {
                "provider": "openai_compatible",
                "model": "default",
                "base_url_env": "TEST_LOCAL_URL",
                "api_key_env": "TEST_LOCAL_KEY",
                "default_base_url": "http://localhost:8000/v1",
                "default_api_key": "not-needed",
                "max_tokens": 2048,
                "description": "테스트 로컬",
            },
        },
        "default_model": "test-claude",
    }
    path = tmp_path / "llm_config.yaml"
    import yaml
    path.write_text(yaml.dump(config, allow_unicode=True), encoding="utf-8")
    return str(path)


@pytest.fixture
def router(config_file: str) -> LLMRouter:
    return LLMRouter(config_path=config_file)


class TestLLMRouterConfig:
    """설정 로드 테스트."""

    def test_load_models(self, router: LLMRouter) -> None:
        models = router.list_models()
        assert len(models) == 3
        ids = {m.id for m in models}
        assert "test-claude" in ids
        assert "test-gpt" in ids
        assert "test-local" in ids

    def test_default_model(self, router: LLMRouter) -> None:
        assert router.default_model == "test-claude"

    def test_switch_default(self, router: LLMRouter) -> None:
        router.default_model = "test-gpt"
        assert router.default_model == "test-gpt"

    def test_switch_invalid_model(self, router: LLMRouter) -> None:
        with pytest.raises(ValueError, match="등록되지 않은"):
            router.default_model = "nonexistent"

    def test_model_availability(self, router: LLMRouter) -> None:
        models = router.list_models()
        local = next(m for m in models if m.id == "test-local")
        # default_api_key가 있으므로 available
        assert local.available is True

    def test_missing_config(self, tmp_path) -> None:
        router = LLMRouter(config_path=str(tmp_path / "nonexistent.yaml"))
        assert router.default_model == "claude-sonnet"

    def test_model_info_fields(self, router: LLMRouter) -> None:
        models = router.list_models()
        claude = next(m for m in models if m.id == "test-claude")
        assert claude.provider == "anthropic"
        assert claude.model == "claude-sonnet-4-20250514"
        assert claude.description == "테스트 Claude"
        assert claude.max_tokens == 4096


class TestLLMRouterChat:
    """LLM 호출 테스트 (mock)."""

    @pytest.mark.asyncio
    async def test_chat_anthropic(self, router: LLMRouter) -> None:
        """Anthropic 호출 mock 테스트."""
        mock_response = MagicMock()
        mock_response.content = [MagicMock(type="text", text="안녕하세요!")]
        mock_response.usage = MagicMock(input_tokens=10, output_tokens=5)

        mock_client = AsyncMock()
        mock_client.messages.create = AsyncMock(return_value=mock_response)

        with patch.object(router, "_get_client", return_value=(mock_client, {"provider": "anthropic", "model": "test"})):
            resp = await router.chat(
                messages=[{"role": "user", "content": "안녕"}],
                model_id="test-claude",
            )

        assert isinstance(resp, LLMResponse)
        assert resp.content == "안녕하세요!"
        assert resp.usage.input_tokens == 10

    @pytest.mark.asyncio
    async def test_chat_openai(self, router: LLMRouter) -> None:
        """OpenAI 호출 mock 테스트."""
        mock_msg = MagicMock()
        mock_msg.content = "Hello!"
        mock_msg.tool_calls = None

        mock_choice = MagicMock()
        mock_choice.message = mock_msg

        mock_response = MagicMock()
        mock_response.choices = [mock_choice]
        mock_response.usage = MagicMock(prompt_tokens=8, completion_tokens=3)

        mock_client = AsyncMock()
        mock_client.chat.completions.create = AsyncMock(return_value=mock_response)

        with patch.object(router, "_get_client", return_value=(mock_client, {"provider": "openai", "model": "gpt-4o"})):
            resp = await router.chat(
                messages=[{"role": "user", "content": "Hi"}],
                model_id="test-gpt",
            )

        assert isinstance(resp, LLMResponse)
        assert resp.content == "Hello!"

    @pytest.mark.asyncio
    async def test_chat_with_tool_calls(self, router: LLMRouter) -> None:
        """도구 호출 응답 파싱 테스트 (OpenAI)."""
        mock_tc = MagicMock()
        mock_tc.id = "call_123"
        mock_tc.function.name = "read_table"
        mock_tc.function.arguments = '{"table_idx": 0}'

        mock_msg = MagicMock()
        mock_msg.content = ""
        mock_msg.tool_calls = [mock_tc]

        mock_choice = MagicMock()
        mock_choice.message = mock_msg

        mock_response = MagicMock()
        mock_response.choices = [mock_choice]
        mock_response.usage = MagicMock(prompt_tokens=20, completion_tokens=10)

        mock_client = AsyncMock()
        mock_client.chat.completions.create = AsyncMock(return_value=mock_response)

        with patch.object(router, "_get_client", return_value=(mock_client, {"provider": "openai", "model": "gpt-4o"})):
            resp = await router.chat(
                messages=[{"role": "user", "content": "표를 분석해줘"}],
                model_id="test-gpt",
                tools=[{"name": "read_table", "description": "표 읽기", "parameters": {}}],
            )

        assert len(resp.tool_calls) == 1
        assert resp.tool_calls[0].name == "read_table"
        assert resp.tool_calls[0].arguments == {"table_idx": 0}


class TestDataClasses:
    """데이터 클래스 테스트."""

    def test_tool_call(self) -> None:
        tc = ToolCall(id="abc", name="write_cell", arguments={"row": 0, "col": 1, "text": "hi"})
        assert tc.name == "write_cell"
        assert tc.arguments["text"] == "hi"

    def test_token_usage(self) -> None:
        usage = TokenUsage(input_tokens=100, output_tokens=50, estimated_cost_usd=0.001)
        assert usage.input_tokens == 100
        assert usage.estimated_cost_usd == 0.001

    def test_llm_response(self) -> None:
        resp = LLMResponse(content="Hello", model="test")
        assert resp.content == "Hello"
        assert resp.tool_calls == []

    def test_model_info(self) -> None:
        info = ModelInfo(id="test", provider="anthropic", model="claude", available=True)
        assert info.available is True


class TestToolConversion:
    """도구 형식 변환 테스트."""

    def test_anthropic_format(self) -> None:
        tools = [{"name": "test", "description": "desc", "parameters": {"type": "object"}}]
        result = LLMRouter._convert_tools_to_anthropic(tools)
        assert result[0]["name"] == "test"
        assert "input_schema" in result[0]
        assert "parameters" not in result[0]

    def test_openai_format(self) -> None:
        tools = [{"name": "test", "description": "desc", "parameters": {"type": "object"}}]
        result = LLMRouter._convert_tools_to_openai(tools)
        assert result[0]["type"] == "function"
        assert result[0]["function"]["name"] == "test"

    def test_split_system_message(self) -> None:
        messages = [
            {"role": "system", "content": "You are helpful"},
            {"role": "user", "content": "Hi"},
        ]
        system, chat = LLMRouter._split_system_message(messages)
        assert system == "You are helpful"
        assert len(chat) == 1
        assert chat[0]["role"] == "user"
