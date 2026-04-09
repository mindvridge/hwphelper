"""멀티 LLM 라우터 — Anthropic, OpenAI, OpenAI 호환 API를 통합.

사용법::

    router = LLMRouter("config/llm_config.yaml")
    response = await router.chat(
        messages=[{"role": "user", "content": "안녕하세요"}],
        model_id="claude-sonnet",
    )
    print(response.content)
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, AsyncIterator, Callable

from dotenv import load_dotenv

load_dotenv()

import structlog
import yaml

logger = structlog.get_logger()


# ------------------------------------------------------------------
# 데이터 클래스
# ------------------------------------------------------------------


@dataclass
class ToolCall:
    """도구 호출 정보."""

    id: str
    name: str
    arguments: dict[str, Any]


@dataclass
class TokenUsage:
    """토큰 사용량."""

    input_tokens: int = 0
    output_tokens: int = 0
    estimated_cost_usd: float = 0.0


@dataclass
class LLMResponse:
    """LLM 응답."""

    content: str = ""
    tool_calls: list[ToolCall] = field(default_factory=list)
    model: str = ""
    usage: TokenUsage = field(default_factory=TokenUsage)
    raw_response: Any = None


@dataclass
class ModelInfo:
    """모델 메타 정보."""

    id: str
    provider: str
    model: str
    description: str = ""
    available: bool = False
    max_tokens: int = 4096
    estimated_cost_per_1k: float = 0.0


# ------------------------------------------------------------------
# 비용 추정 (1K input token 기준, USD)
# ------------------------------------------------------------------

_COST_MAP: dict[str, tuple[float, float]] = {
    # (input_per_1k, output_per_1k)
    "claude-sonnet-4-20250514": (0.003, 0.015),
    "claude-opus-4-20250514": (0.015, 0.075),
    "gpt-4o": (0.0025, 0.01),
    "gpt-4o-mini": (0.00015, 0.0006),
    "deepseek-chat": (0.00014, 0.00028),
}


def _estimate_cost(model: str, usage: TokenUsage) -> float:
    rates = _COST_MAP.get(model, (0.001, 0.002))
    return (usage.input_tokens / 1000 * rates[0]) + (usage.output_tokens / 1000 * rates[1])


# ------------------------------------------------------------------
# LLMRouter
# ------------------------------------------------------------------


class LLMRouter:
    """여러 LLM 프로바이더를 통합하는 라우터."""

    def __init__(self, config_path: str = "config/llm_config.yaml") -> None:
        self._config_path = config_path
        self._models: dict[str, dict[str, Any]] = {}
        self._default_model: str = ""
        self._clients: dict[str, Any] = {}  # provider → client

        self._load_config()

    # ------------------------------------------------------------------
    # 설정 로드
    # ------------------------------------------------------------------

    def _load_config(self) -> None:
        """YAML 설정 파일에서 모델 목록을 로드한다."""
        path = Path(self._config_path)
        if not path.exists():
            logger.warning("LLM 설정 파일 없음, 기본 설정 사용", path=self._config_path)
            self._default_model = "claude-sonnet"
            return

        with open(path, encoding="utf-8") as f:
            config = yaml.safe_load(f)

        self._models = config.get("models", {})
        self._default_model = config.get("default_model", "claude-sonnet")
        logger.info("LLM 설정 로드", models=list(self._models.keys()), default=self._default_model)

    def _get_client(self, model_id: str) -> tuple[Any, dict[str, Any]]:
        """model_id에 대한 클라이언트와 모델 설정을 반환한다."""
        model_cfg = self._models.get(model_id)
        if not model_cfg:
            raise ValueError(f"등록되지 않은 모델: {model_id}. 사용 가능: {list(self._models.keys())}")

        provider = model_cfg["provider"]
        api_key = os.environ.get(model_cfg.get("api_key_env", ""), "") or model_cfg.get("default_api_key", "")
        base_url = os.environ.get(model_cfg.get("base_url_env", ""), "") or model_cfg.get("default_base_url", "")

        cache_key = f"{provider}:{api_key[:8]}:{base_url}"

        if cache_key not in self._clients:
            if provider == "anthropic":
                from anthropic import AsyncAnthropic

                self._clients[cache_key] = AsyncAnthropic(api_key=api_key or None)
                logger.info("Anthropic 클라이언트 초기화", model_id=model_id)

            elif provider == "openai":
                from openai import AsyncOpenAI

                kwargs: dict[str, Any] = {}
                if api_key:
                    kwargs["api_key"] = api_key
                self._clients[cache_key] = AsyncOpenAI(**kwargs)
                logger.info("OpenAI 클라이언트 초기화", model_id=model_id)

            elif provider == "openai_compatible":
                # OpenAI 호환 API는 SDK가 아닌 httpx를 직접 사용
                # (SDK의 인증 헤더가 호환 서버와 충돌하는 문제 방지)
                import httpx

                self._clients[cache_key] = {
                    "_httpx": httpx.AsyncClient(timeout=120),
                    "_base_url": base_url.rstrip("/"),
                    "_api_key": api_key,
                }
                logger.info("OpenAI 호환 클라이언트 초기화 (httpx)", model_id=model_id, base_url=base_url)

            else:
                raise ValueError(f"지원하지 않는 프로바이더: {provider}")

        return self._clients[cache_key], model_cfg

    # ------------------------------------------------------------------
    # public API
    # ------------------------------------------------------------------

    @property
    def default_model(self) -> str:
        return self._default_model

    @default_model.setter
    def default_model(self, model_id: str) -> None:
        if model_id not in self._models:
            raise ValueError(f"등록되지 않은 모델: {model_id}")
        self._default_model = model_id

    def list_models(self) -> list[ModelInfo]:
        """사용 가능한 모델 목록을 반환한다."""
        result: list[ModelInfo] = []
        for mid, cfg in self._models.items():
            api_key_env = cfg.get("api_key_env", "")
            has_key = bool(os.environ.get(api_key_env, "")) or bool(cfg.get("default_api_key", ""))

            model_name = cfg.get("model", "")
            cost_rates = _COST_MAP.get(model_name, (0.001, 0.002))

            result.append(ModelInfo(
                id=mid,
                provider=cfg.get("provider", ""),
                model=model_name,
                description=cfg.get("description", ""),
                available=has_key,
                max_tokens=cfg.get("max_tokens", 4096),
                estimated_cost_per_1k=cost_rates[0],
            ))
        return result

    async def chat(
        self,
        messages: list[dict[str, Any]],
        model_id: str | None = None,
        tools: list[dict[str, Any]] | None = None,
        temperature: float = 0.3,
        max_tokens: int = 4096,
        stream: bool = False,
    ) -> LLMResponse | AsyncIterator[str]:
        """통합 채팅 API.

        Parameters
        ----------
        messages : list[dict]
            대화 메시지 목록.
        model_id : str | None
            사용할 모델 ID. None이면 default_model.
        tools : list[dict] | None
            도구 정의 (function calling).
        temperature : float
        max_tokens : int
        stream : bool
            True이면 AsyncIterator[str]로 토큰 단위 스트리밍.

        Returns
        -------
        LLMResponse | AsyncIterator[str]
        """
        mid = model_id or self._default_model
        client, cfg = self._get_client(mid)
        provider = cfg["provider"]
        model_name = cfg["model"]

        # supports_tools가 false이면 도구를 전달하지 않음
        effective_tools = tools if cfg.get("supports_tools", True) else None

        # 429 Rate Limit 재시도 (최대 3회, 지수 백오프)
        import asyncio as _aio
        max_retries = 3
        for attempt in range(max_retries + 1):
            try:
                if provider == "anthropic":
                    if stream:
                        return self._stream_anthropic(client, model_name, messages, effective_tools, temperature, max_tokens)
                    return await self._call_anthropic(client, model_name, messages, effective_tools, temperature, max_tokens)
                elif provider == "openai_compatible":
                    tc_override = cfg.get("tool_choice", "auto") if effective_tools else None
                    if stream:
                        return self._stream_httpx(client, model_name, messages, effective_tools, temperature, max_tokens)
                    return await self._call_httpx(client, model_name, messages, effective_tools, temperature, max_tokens, tc_override)
                else:
                    if stream:
                        return self._stream_openai(client, model_name, messages, effective_tools, temperature, max_tokens)
                    return await self._call_openai(client, model_name, messages, effective_tools, temperature, max_tokens)
            except Exception as exc:
                # httpx.HTTPStatusError → exc.response.status_code
                status = getattr(exc, "status_code", 0) or getattr(exc, "status", 0)
                if hasattr(exc, "response"):
                    status = getattr(exc.response, "status_code", status)
                if status == 429 and attempt < max_retries:
                    wait = 5 * (attempt + 1)  # 5, 10, 15초
                    logger.warning("Rate limit (429), 재시도", attempt=attempt + 1, wait=wait)
                    await _aio.sleep(wait)
                    continue
                raise

    async def chat_with_tools(
        self,
        messages: list[dict[str, Any]],
        tools: list[dict[str, Any]],
        model_id: str | None = None,
        max_tool_rounds: int = 5,
        tool_executor: Callable | None = None,
    ) -> LLMResponse:
        """도구 호출을 자동으로 실행하는 루프.

        1. LLM 호출
        2. 도구 호출 응답이면 → tool_executor()로 실행
        3. 실행 결과를 messages에 추가
        4. 다시 LLM 호출 (최대 max_tool_rounds회)
        5. 최종 텍스트 응답 반환
        """
        mid = model_id or self._default_model
        _, cfg = self._get_client(mid)
        provider = cfg["provider"]
        working_messages = list(messages)

        for round_num in range(max_tool_rounds):
            response = await self.chat(working_messages, model_id=mid, tools=tools)
            assert isinstance(response, LLMResponse)

            if not response.tool_calls:
                return response

            logger.info("도구 호출 감지", round=round_num + 1, tools=[tc.name for tc in response.tool_calls])

            # 어시스턴트 메시지 추가
            if provider == "anthropic":
                working_messages.append({
                    "role": "assistant",
                    "content": self._build_anthropic_assistant_content(response),
                })
            else:
                working_messages.append({
                    "role": "assistant",
                    "content": response.content or None,
                    "tool_calls": [
                        {
                            "id": tc.id,
                            "type": "function",
                            "function": {"name": tc.name, "arguments": json.dumps(tc.arguments, ensure_ascii=False)},
                        }
                        for tc in response.tool_calls
                    ],
                })

            # 도구 실행
            for tc in response.tool_calls:
                if tool_executor:
                    result = await tool_executor(tc)
                else:
                    result = {"error": "도구 실행기가 설정되지 않았습니다."}

                result_str = json.dumps(result, ensure_ascii=False) if isinstance(result, dict) else str(result)

                if provider == "anthropic":
                    working_messages.append({
                        "role": "user",
                        "content": [{"type": "tool_result", "tool_use_id": tc.id, "content": result_str}],
                    })
                else:
                    working_messages.append({
                        "role": "tool",
                        "tool_call_id": tc.id,
                        "content": result_str,
                    })

        # max_tool_rounds 초과
        logger.warning("도구 호출 라운드 초과", max_rounds=max_tool_rounds)
        return await self.chat(working_messages, model_id=mid)

    # ------------------------------------------------------------------
    # Anthropic 호출
    # ------------------------------------------------------------------

    async def _call_anthropic(
        self, client: Any, model: str, messages: list[dict], tools: list[dict] | None,
        temperature: float, max_tokens: int,
    ) -> LLMResponse:
        """Anthropic Messages API 호출."""
        # system 메시지 분리
        system_msg, chat_msgs = self._split_system_message(messages)

        # 멀티모달 콘텐츠 변환 (image_url → image 블록)
        chat_msgs = [
            {**m, "content": self._convert_content_for_anthropic(m["content"])}
            for m in chat_msgs
        ]

        kwargs: dict[str, Any] = {
            "model": model,
            "messages": chat_msgs,
            "max_tokens": max_tokens,
            "temperature": temperature,
        }
        if system_msg:
            kwargs["system"] = system_msg
        if tools:
            kwargs["tools"] = self._convert_tools_to_anthropic(tools)

        resp = await client.messages.create(**kwargs)

        # 응답 파싱
        content_text = ""
        tool_calls: list[ToolCall] = []
        for block in resp.content:
            if block.type == "text":
                content_text += block.text
            elif block.type == "tool_use":
                tool_calls.append(ToolCall(
                    id=block.id,
                    name=block.name,
                    arguments=block.input if isinstance(block.input, dict) else {},
                ))

        usage = TokenUsage(
            input_tokens=resp.usage.input_tokens,
            output_tokens=resp.usage.output_tokens,
        )
        usage.estimated_cost_usd = _estimate_cost(model, usage)

        return LLMResponse(
            content=content_text,
            tool_calls=tool_calls,
            model=model,
            usage=usage,
            raw_response=resp,
        )

    async def _stream_anthropic(
        self, client: Any, model: str, messages: list[dict], tools: list[dict] | None,
        temperature: float, max_tokens: int,
    ) -> AsyncIterator[str]:
        """Anthropic 스트리밍."""
        system_msg, chat_msgs = self._split_system_message(messages)

        kwargs: dict[str, Any] = {
            "model": model,
            "messages": chat_msgs,
            "max_tokens": max_tokens,
            "temperature": temperature,
        }
        if system_msg:
            kwargs["system"] = system_msg
        if tools:
            kwargs["tools"] = self._convert_tools_to_anthropic(tools)

        async with client.messages.stream(**kwargs) as stream:
            async for text in stream.text_stream:
                yield text

    # ------------------------------------------------------------------
    # OpenAI / OpenAI-compatible 호출
    # ------------------------------------------------------------------

    async def _call_openai(
        self, client: Any, model: str, messages: list[dict], tools: list[dict] | None,
        temperature: float, max_tokens: int,
    ) -> LLMResponse:
        """OpenAI Chat Completions API 호출."""
        # 멀티모달 콘텐츠 변환 (image → image_url 블록)
        messages = [
            {**m, "content": self._convert_content_for_openai(m["content"])}
            for m in messages
        ]
        kwargs: dict[str, Any] = {
            "model": model,
            "messages": messages,
            "max_tokens": max_tokens,
            "temperature": temperature,
        }
        if tools:
            kwargs["tools"] = self._convert_tools_to_openai(tools)
            kwargs["tool_choice"] = "auto"

        # openai_compatible 서버는 SDK 인증과 충돌할 수 있으므로 extra_headers 사용
        if hasattr(client, "_default_headers") and "Authorization" in (client._default_headers or {}):
            kwargs["extra_headers"] = dict(client._default_headers)

        resp = await client.chat.completions.create(**kwargs)
        choice = resp.choices[0]
        msg = choice.message

        # 도구 호출 파싱
        tool_calls: list[ToolCall] = []
        if msg.tool_calls:
            for tc in msg.tool_calls:
                try:
                    args = json.loads(tc.function.arguments)
                except (json.JSONDecodeError, TypeError):
                    args = {}
                tool_calls.append(ToolCall(
                    id=tc.id,
                    name=tc.function.name,
                    arguments=args,
                ))

        usage_data = resp.usage
        usage = TokenUsage(
            input_tokens=getattr(usage_data, "prompt_tokens", 0) if usage_data else 0,
            output_tokens=getattr(usage_data, "completion_tokens", 0) if usage_data else 0,
        )
        usage.estimated_cost_usd = _estimate_cost(model, usage)

        return LLMResponse(
            content=msg.content or "",
            tool_calls=tool_calls,
            model=model,
            usage=usage,
            raw_response=resp,
        )

    async def _stream_openai(
        self, client: Any, model: str, messages: list[dict], tools: list[dict] | None,
        temperature: float, max_tokens: int,
    ) -> AsyncIterator[str]:
        """OpenAI 스트리밍."""
        kwargs: dict[str, Any] = {
            "model": model,
            "messages": messages,
            "max_tokens": max_tokens,
            "temperature": temperature,
            "stream": True,
        }
        if tools:
            kwargs["tools"] = self._convert_tools_to_openai(tools)

        stream = await client.chat.completions.create(**kwargs)
        async for chunk in stream:
            if chunk.choices and chunk.choices[0].delta.content:
                yield chunk.choices[0].delta.content

    # ------------------------------------------------------------------
    # OpenAI 호환 (httpx 직접 호출)
    # ------------------------------------------------------------------

    async def _call_httpx(
        self, client_info: dict, model: str, messages: list[dict], tools: list[dict] | None,
        temperature: float, max_tokens: int, tool_choice_override: str | None = None,
    ) -> LLMResponse:
        """httpx로 OpenAI 호환 API를 직접 호출한다."""
        http: Any = client_info["_httpx"]
        base_url: str = client_info["_base_url"]
        api_key: str = client_info["_api_key"]

        # 멀티모달 콘텐츠 변환 (image → image_url 블록)
        messages = [
            {**m, "content": self._convert_content_for_openai(m["content"])}
            for m in messages
        ]
        body: dict[str, Any] = {
            "model": model,
            "messages": messages,
            "max_tokens": max_tokens,
            "temperature": temperature,
        }
        if tools:
            body["tools"] = self._convert_tools_to_openai(tools)
            body["tool_choice"] = tool_choice_override or "auto"

        headers = {"Content-Type": "application/json"}
        if api_key:
            headers["Authorization"] = f"Bearer {api_key}"

        import asyncio as _aio_retry
        for _attempt in range(4):
            resp = await http.post(f"{base_url}/chat/completions", json=body, headers=headers)
            if resp.status_code == 429 and _attempt < 3:
                wait = 5 * (_attempt + 1)  # 5, 10, 15초
                logger.warning("httpx 429 rate limit, 재시도", attempt=_attempt + 1, wait=wait)
                await _aio_retry.sleep(wait)
                continue
            resp.raise_for_status()
            break
        data = resp.json()

        choice = data["choices"][0]
        msg = choice["message"]

        tool_calls: list[ToolCall] = []
        for tc in msg.get("tool_calls") or []:
            try:
                args = json.loads(tc["function"]["arguments"]) if isinstance(tc["function"]["arguments"], str) else tc["function"]["arguments"]
            except (json.JSONDecodeError, KeyError, TypeError):
                args = {}
            tool_calls.append(ToolCall(id=tc.get("id", ""), name=tc["function"]["name"], arguments=args))

        usage_data = data.get("usage", {})
        usage = TokenUsage(
            input_tokens=usage_data.get("prompt_tokens", 0),
            output_tokens=usage_data.get("completion_tokens", 0),
        )
        usage.estimated_cost_usd = _estimate_cost(model, usage)

        return LLMResponse(content=msg.get("content", ""), tool_calls=tool_calls, model=model, usage=usage, raw_response=data)

    async def _stream_httpx(
        self, client_info: dict, model: str, messages: list[dict], tools: list[dict] | None,
        temperature: float, max_tokens: int,
    ) -> AsyncIterator[str]:
        """httpx로 OpenAI 호환 API를 스트리밍한다."""
        http: Any = client_info["_httpx"]
        base_url: str = client_info["_base_url"]
        api_key: str = client_info["_api_key"]

        body: dict[str, Any] = {
            "model": model,
            "messages": messages,
            "max_tokens": max_tokens,
            "temperature": temperature,
            "stream": True,
        }

        headers = {"Content-Type": "application/json"}
        if api_key:
            headers["Authorization"] = f"Bearer {api_key}"

        async with http.stream("POST", f"{base_url}/chat/completions", json=body, headers=headers) as resp:
            async for line in resp.aiter_lines():
                if not line.startswith("data:"):
                    continue
                data_str = line[5:].strip()
                if data_str == "[DONE]":
                    break
                try:
                    chunk = json.loads(data_str)
                    content = chunk.get("choices", [{}])[0].get("delta", {}).get("content")
                    if content:
                        yield content
                except json.JSONDecodeError:
                    continue

    # ------------------------------------------------------------------
    # 유틸리티
    # ------------------------------------------------------------------

    @staticmethod
    def _split_system_message(messages: list[dict]) -> tuple[str, list[dict]]:
        """system 역할 메시지를 분리한다 (Anthropic용)."""
        system_parts: list[str] = []
        chat_msgs: list[dict] = []
        for msg in messages:
            if msg.get("role") == "system":
                content = msg["content"]
                if isinstance(content, str):
                    system_parts.append(content)
                # 멀티모달 system은 무시 (텍스트만 추출)
                elif isinstance(content, list):
                    for block in content:
                        if isinstance(block, dict) and block.get("type") == "text":
                            system_parts.append(block["text"])
            else:
                chat_msgs.append(msg)
        return "\n\n".join(system_parts), chat_msgs

    @staticmethod
    def _convert_content_for_anthropic(content: Any) -> Any:
        """멀티모달 콘텐츠를 Anthropic 형식으로 변환한다.

        OpenAI 형식 image_url → Anthropic 형식 image 블록 변환.
        """
        if isinstance(content, str):
            return content
        if not isinstance(content, list):
            return content

        converted: list[dict] = []
        for block in content:
            if not isinstance(block, dict):
                converted.append(block)
                continue

            block_type = block.get("type", "")

            if block_type == "text":
                converted.append({"type": "text", "text": block["text"]})

            elif block_type == "image":
                # 이미 Anthropic 형식
                converted.append(block)

            elif block_type == "image_url":
                # OpenAI 형식 → Anthropic 형식
                url = block.get("image_url", {}).get("url", "")
                if url.startswith("data:"):
                    # data:image/png;base64,... 파싱
                    header, b64_data = url.split(",", 1)
                    media_type = header.split(":")[1].split(";")[0]
                    converted.append({
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": media_type,
                            "data": b64_data,
                        },
                    })
                else:
                    converted.append({
                        "type": "image",
                        "source": {"type": "url", "url": url},
                    })
            else:
                converted.append(block)

        return converted

    @staticmethod
    def _convert_content_for_openai(content: Any) -> Any:
        """멀티모달 콘텐츠를 OpenAI 형식으로 변환한다.

        Anthropic 형식 image 블록 → OpenAI 형식 image_url 변환.
        """
        if isinstance(content, str):
            return content
        if not isinstance(content, list):
            return content

        converted: list[dict] = []
        for block in content:
            if not isinstance(block, dict):
                converted.append(block)
                continue

            block_type = block.get("type", "")

            if block_type == "text":
                converted.append({"type": "text", "text": block["text"]})

            elif block_type == "image_url":
                converted.append(block)

            elif block_type == "image":
                source = block.get("source", {})
                if source.get("type") == "base64":
                    media_type = source.get("media_type", "image/png")
                    data = source.get("data", "")
                    converted.append({
                        "type": "image_url",
                        "image_url": {"url": f"data:{media_type};base64,{data}"},
                    })
                elif source.get("type") == "url":
                    converted.append({
                        "type": "image_url",
                        "image_url": {"url": source["url"]},
                    })
            else:
                converted.append(block)

        return converted

    @staticmethod
    def _convert_tools_to_anthropic(tools: list[dict]) -> list[dict]:
        """통합 도구 정의를 Anthropic 형식으로 변환."""
        result = []
        for t in tools:
            result.append({
                "name": t["name"],
                "description": t.get("description", ""),
                "input_schema": t.get("parameters", {"type": "object", "properties": {}}),
            })
        return result

    @staticmethod
    def _convert_tools_to_openai(tools: list[dict]) -> list[dict]:
        """통합 도구 정의를 OpenAI 형식으로 변환."""
        result = []
        for t in tools:
            result.append({
                "type": "function",
                "function": {
                    "name": t["name"],
                    "description": t.get("description", ""),
                    "parameters": t.get("parameters", {"type": "object", "properties": {}}),
                },
            })
        return result

    @staticmethod
    def _build_anthropic_assistant_content(response: LLMResponse) -> list[dict]:
        """Anthropic 멀티턴용 어시스턴트 content 블록을 구성한다."""
        blocks: list[dict] = []
        if response.content:
            blocks.append({"type": "text", "text": response.content})
        for tc in response.tool_calls:
            blocks.append({"type": "tool_use", "id": tc.id, "name": tc.name, "input": tc.arguments})
        return blocks
