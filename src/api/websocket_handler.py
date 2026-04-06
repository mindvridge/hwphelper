"""WebSocket 채팅 핸들러 — 실시간 양방향 통신으로 ChatAgent 이벤트를 스트리밍.

프로토콜 (JSON 메시지):

클라이언트 → 서버::

    { "type": "message", "content": "사업 개요를 작성해줘", "model_id": "claude-sonnet" }
    { "type": "ping" }

서버 → 클라이언트::

    { "type": "text_delta", "content": "네, " }
    { "type": "tool_start", "tool": "analyze_document", "description": "문서 분석 중..." }
    { "type": "tool_result", "tool": "write_cell", "success": true, "description": "..." }
    { "type": "document_updated", "changes": [...] }
    { "type": "progress", "current": 5, "total": 23, "description": "셀 채우기..." }
    { "type": "done", "usage": { "input_tokens": 1234, "output_tokens": 567, "cost": 0.003 } }
    { "type": "error", "message": "..." }
    { "type": "pong" }
"""

from __future__ import annotations

import json
from typing import Any

import structlog
from fastapi import WebSocket, WebSocketDisconnect

from src.ai.chat_agent import ChatAgent, ChatEvent

logger = structlog.get_logger()


# ------------------------------------------------------------------
# 도구 설명 매핑 (UI 표시용)
# ------------------------------------------------------------------

_TOOL_DESCRIPTIONS: dict[str, str] = {
    "analyze_document": "문서 구조를 분석하고 있습니다...",
    "read_table": "표를 읽고 있습니다...",
    "read_cell": "셀 내용을 확인하고 있습니다...",
    "write_cell": "셀에 내용을 작성하고 있습니다...",
    "fill_field": "필드를 채우고 있습니다...",
    "fill_all_empty_cells": "빈 셀을 자동으로 채우고 있습니다...",
    "validate_format": "서식을 검증하고 있습니다...",
    "undo": "되돌리기를 실행합니다...",
    "save_document": "문서를 저장하고 있습니다...",
    "get_document_info": "문서 정보를 조회하고 있습니다...",
}


# ------------------------------------------------------------------
# ChatWebSocketHandler
# ------------------------------------------------------------------


class ChatWebSocketHandler:
    """WebSocket으로 실시간 채팅을 처리한다."""

    def __init__(self, chat_agent: ChatAgent) -> None:
        self.agent = chat_agent

    async def handle(self, websocket: WebSocket, session_id: str) -> None:
        """WebSocket 연결을 처리한다."""
        await websocket.accept()
        logger.info("WebSocket 연결", session_id=session_id)

        try:
            while True:
                raw = await websocket.receive_text()
                try:
                    msg = json.loads(raw)
                except json.JSONDecodeError:
                    await self._send(websocket, {"type": "error", "message": "잘못된 JSON 형식"})
                    continue

                msg_type = msg.get("type", "")

                if msg_type == "ping":
                    await self._send(websocket, {"type": "pong"})

                elif msg_type == "message":
                    content = msg.get("content", "").strip()
                    model_id = msg.get("model_id")
                    image_data = msg.get("image")  # data:URL 형식

                    if not content and not image_data:
                        await self._send(websocket, {"type": "error", "message": "빈 메시지"})
                        continue

                    # 이미지가 있으면 비전 형식으로 변환
                    if image_data:
                        content = self._build_vision_content(content, image_data)

                    await self._process_chat(websocket, session_id, content, model_id)

                else:
                    await self._send(websocket, {"type": "error", "message": f"알 수 없는 타입: {msg_type}"})

        except WebSocketDisconnect:
            logger.info("WebSocket 연결 해제", session_id=session_id)
        except Exception:
            logger.exception("WebSocket 오류", session_id=session_id)
            try:
                await self._send(websocket, {"type": "error", "message": "서버 내부 오류"})
            except Exception:
                pass

    async def _process_chat(
        self,
        websocket: WebSocket,
        session_id: str,
        content: str,
        model_id: str | None,
    ) -> None:
        """ChatAgent를 호출하고 이벤트를 WebSocket으로 전달한다."""
        try:
            async for event in self.agent.process_message(session_id, content, model_id):
                ws_msg = self._event_to_ws_message(event)
                await self._send(websocket, ws_msg)
        except Exception as e:
            logger.exception("채팅 처리 오류", session_id=session_id)
            await self._send(websocket, {"type": "error", "message": str(e)})

    @staticmethod
    def _build_vision_content(text: str, image_data_url: str) -> str:
        """이미지 data:URL을 포함한 비전 메시지 텍스트를 구성한다.

        ChatAgent.process_message는 문자열만 받으므로,
        특수 마커를 삽입하여 LLM 호출 시 비전 형식으로 변환한다.
        """
        return f"[IMAGE:{image_data_url}]\n{text}"

    @staticmethod
    def _event_to_ws_message(event: ChatEvent) -> dict[str, Any]:
        """ChatEvent를 WebSocket 메시지 형식으로 변환한다."""
        if event.type == "text_delta":
            return {"type": "text_delta", "content": event.data}

        elif event.type == "tool_start":
            tool_name = event.data.get("tool", "") if isinstance(event.data, dict) else ""
            return {
                "type": "tool_start",
                "tool": tool_name,
                "description": _TOOL_DESCRIPTIONS.get(tool_name, f"{tool_name} 실행 중..."),
                "args": event.data.get("args", {}) if isinstance(event.data, dict) else {},
            }

        elif event.type == "tool_result":
            data = event.data if isinstance(event.data, dict) else {}
            tool_name = data.get("tool", "")
            result = data.get("result", {})
            success = not bool(result.get("error")) if isinstance(result, dict) else True
            return {
                "type": "tool_result",
                "tool": tool_name,
                "success": success,
                "result": result,
                "description": _build_result_description(tool_name, result),
            }

        elif event.type == "document_updated":
            return {
                "type": "document_updated",
                "changes": event.data if isinstance(event.data, dict) else {},
            }

        elif event.type == "done":
            return {
                "type": "done",
                "usage": event.data if isinstance(event.data, dict) else {},
            }

        elif event.type == "error":
            msg = event.data.get("message", str(event.data)) if isinstance(event.data, dict) else str(event.data)
            return {"type": "error", "message": msg}

        else:
            return {"type": event.type, "data": event.data}

    @staticmethod
    async def _send(websocket: WebSocket, data: dict[str, Any]) -> None:
        """WebSocket으로 JSON 메시지를 전송한다."""
        await websocket.send_json(data)


# ------------------------------------------------------------------
# 결과 설명 생성
# ------------------------------------------------------------------


def _build_result_description(tool_name: str, result: Any) -> str:
    """도구 실행 결과를 사람이 읽을 수 있는 설명으로 변환한다."""
    if not isinstance(result, dict):
        return f"{tool_name} 완료"

    if result.get("error"):
        return f"오류: {result['error']}"

    if tool_name == "analyze_document":
        tables = result.get("total_tables", 0)
        cells = result.get("total_cells_to_fill", 0)
        return f"표 {tables}개 발견, 채울 셀 {cells}개"

    if tool_name == "read_table":
        rows = result.get("rows", 0)
        cols = result.get("cols", 0)
        return f"{rows}행 x {cols}열 표 읽기 완료"

    if tool_name == "write_cell":
        row = result.get("row", "?")
        col = result.get("col", "?")
        return f"({row},{col}) 셀에 작성 완료"

    if tool_name == "fill_field":
        field = result.get("field", "")
        return f"필드 '{field}' 채우기 완료"

    if tool_name == "fill_all_empty_cells":
        count = result.get("cells_to_fill", 0)
        return f"{count}개 셀 채우기 대상"

    if tool_name == "undo":
        return result.get("message", "되돌리기 완료")

    if tool_name == "save_document":
        return "문서 저장 완료"

    return f"{tool_name} 완료"
