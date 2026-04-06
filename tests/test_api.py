"""FastAPI REST API 및 WebSocket 테스트.

httpx AsyncClient로 FastAPI 앱을 직접 테스트한다.
COM 의존 로직은 mock 처리.
"""

from __future__ import annotations

import json
from unittest.mock import AsyncMock, MagicMock, patch

import pytest
from fastapi.testclient import TestClient

from src.api.schemas import (
    ChatMessage,
    ErrorResponse,
    FileUploadResponse,
    FormatReportResponse,
    ModelListResponse,
    SuccessResponse,
)
from src.server import create_app


# ------------------------------------------------------------------
# 픽스처
# ------------------------------------------------------------------


@pytest.fixture
def app():
    """테스트용 FastAPI 앱."""
    test_app = create_app()

    # mock state 주입
    test_app.state.llm_router = MagicMock()
    test_app.state.llm_router.default_model = "test-model"
    test_app.state.llm_router.list_models.return_value = [
        MagicMock(id="claude-sonnet", provider="anthropic", model="claude-sonnet-4", description="Claude", available=True),
        MagicMock(id="gpt-4o", provider="openai", model="gpt-4o", description="GPT-4o", available=False),
    ]

    test_app.state.doc_manager = MagicMock()
    test_app.state.chat_agent = MagicMock()
    test_app.state.format_checker = MagicMock()

    return test_app


@pytest.fixture
def client(app) -> TestClient:
    return TestClient(app)


# ------------------------------------------------------------------
# 모델 API
# ------------------------------------------------------------------


class TestModelAPI:
    """모델 관리 API 테스트."""

    def test_list_models(self, client: TestClient) -> None:
        resp = client.get("/api/models")
        assert resp.status_code == 200
        data = resp.json()
        assert "models" in data
        assert "default_model" in data
        assert len(data["models"]) == 2

    def test_set_default_model(self, client: TestClient, app) -> None:
        resp = client.post("/api/models/default", json={"model_id": "gpt-4o"})
        assert resp.status_code == 200
        data = resp.json()
        assert data["success"] is True

    def test_set_default_model_invalid(self, client: TestClient, app) -> None:
        # default_model setter에서 ValueError를 발생시키도록 mock
        mock_router = app.state.llm_router
        type(mock_router).default_model = property(
            fget=lambda self: "test-model",
            fset=MagicMock(side_effect=ValueError("등록되지 않은 모델")),
        )
        resp = client.post("/api/models/default", json={"model_id": "nonexistent"})
        assert resp.status_code == 400


# ------------------------------------------------------------------
# 세션 API
# ------------------------------------------------------------------


class TestSessionAPI:
    """세션 관리 API 테스트."""

    def test_get_history(self, client: TestClient, app) -> None:
        mock_session = MagicMock()
        mock_session.current_snapshot_idx = 1
        app.state.doc_manager.get_session.return_value = mock_session
        app.state.doc_manager.get_history.return_value = [
            MagicMock(index=0, description="초기 상태", created_at=MagicMock(isoformat=lambda: "2024-01-01")),
            MagicMock(index=1, description="작업 1", created_at=MagicMock(isoformat=lambda: "2024-01-02")),
        ]

        resp = client.get("/api/sessions/test123/history")
        assert resp.status_code == 200
        data = resp.json()
        assert len(data["snapshots"]) == 2
        assert data["current_idx"] == 1

    def test_get_history_not_found(self, client: TestClient, app) -> None:
        app.state.doc_manager.get_session.side_effect = KeyError("not found")
        resp = client.get("/api/sessions/nonexistent/history")
        assert resp.status_code == 404

    def test_undo(self, client: TestClient, app) -> None:
        app.state.doc_manager.undo.return_value = True
        resp = client.post("/api/sessions/test123/undo")
        assert resp.status_code == 200
        assert resp.json()["success"] is True

    def test_undo_nothing(self, client: TestClient, app) -> None:
        app.state.doc_manager.undo.return_value = False
        resp = client.post("/api/sessions/test123/undo")
        assert resp.status_code == 200
        assert resp.json()["success"] is False

    def test_redo(self, client: TestClient, app) -> None:
        app.state.doc_manager.redo.return_value = True
        resp = client.post("/api/sessions/test123/redo")
        assert resp.status_code == 200
        assert resp.json()["success"] is True

    def test_close_session(self, client: TestClient, app) -> None:
        app.state.doc_manager.close_session.return_value = "/path/to/final.hwp"
        resp = client.delete("/api/sessions/test123")
        assert resp.status_code == 200
        data = resp.json()
        assert data["success"] is True
        assert "final_path" in data["data"]

    def test_close_session_not_found(self, client: TestClient, app) -> None:
        app.state.doc_manager.close_session.side_effect = KeyError("not found")
        resp = client.delete("/api/sessions/nonexistent")
        assert resp.status_code == 404


# ------------------------------------------------------------------
# Pydantic 스키마 테스트
# ------------------------------------------------------------------


class TestSchemas:
    """Pydantic 모델 직렬화 테스트."""

    def test_file_upload_response(self) -> None:
        r = FileUploadResponse(
            session_id="abc123",
            file_name="test.hwp",
            tables_count=3,
            cells_to_fill=15,
            document_schema={"tables": []},
        )
        d = r.model_dump()
        assert d["session_id"] == "abc123"
        assert d["tables_count"] == 3
        assert "document_schema" in d

    def test_chat_message(self) -> None:
        m = ChatMessage(role="user", content="안녕", timestamp="2024-01-01")
        assert m.tool_calls is None

    def test_format_report_response(self) -> None:
        r = FormatReportResponse(
            passed=False,
            total_checks=10,
            passed_checks=7,
            warnings=[{"location": "문단 1", "rule": "font_name"}],
        )
        assert r.passed is False
        assert len(r.warnings) == 1

    def test_success_response(self) -> None:
        r = SuccessResponse(message="OK", data={"key": "value"})
        assert r.success is True

    def test_error_response(self) -> None:
        r = ErrorResponse(error="오류", detail="상세")
        assert r.error == "오류"


# ------------------------------------------------------------------
# WebSocket 핸들러 단위 테스트
# ------------------------------------------------------------------


class TestWebSocketHandler:
    """WebSocket 메시지 변환 테스트."""

    def test_event_to_ws_text_delta(self) -> None:
        from src.ai.chat_agent import ChatEvent
        from src.api.websocket_handler import ChatWebSocketHandler

        event = ChatEvent(type="text_delta", data="안녕하세요")
        msg = ChatWebSocketHandler._event_to_ws_message(event)
        assert msg["type"] == "text_delta"
        assert msg["content"] == "안녕하세요"

    def test_event_to_ws_tool_start(self) -> None:
        from src.ai.chat_agent import ChatEvent
        from src.api.websocket_handler import ChatWebSocketHandler

        event = ChatEvent(type="tool_start", data={"tool": "analyze_document", "args": {}})
        msg = ChatWebSocketHandler._event_to_ws_message(event)
        assert msg["type"] == "tool_start"
        assert msg["tool"] == "analyze_document"
        assert "분석" in msg["description"]

    def test_event_to_ws_tool_result(self) -> None:
        from src.ai.chat_agent import ChatEvent
        from src.api.websocket_handler import ChatWebSocketHandler

        event = ChatEvent(type="tool_result", data={
            "tool": "write_cell",
            "result": {"success": True, "row": 1, "col": 2},
        })
        msg = ChatWebSocketHandler._event_to_ws_message(event)
        assert msg["type"] == "tool_result"
        assert msg["success"] is True
        assert "(1,2)" in msg["description"]

    def test_event_to_ws_done(self) -> None:
        from src.ai.chat_agent import ChatEvent
        from src.api.websocket_handler import ChatWebSocketHandler

        event = ChatEvent(type="done", data={"model": "claude-sonnet"})
        msg = ChatWebSocketHandler._event_to_ws_message(event)
        assert msg["type"] == "done"

    def test_event_to_ws_error(self) -> None:
        from src.ai.chat_agent import ChatEvent
        from src.api.websocket_handler import ChatWebSocketHandler

        event = ChatEvent(type="error", data={"message": "테스트 에러"})
        msg = ChatWebSocketHandler._event_to_ws_message(event)
        assert msg["type"] == "error"
        assert "테스트 에러" in msg["message"]

    def test_result_description_analyze(self) -> None:
        from src.api.websocket_handler import _build_result_description

        desc = _build_result_description("analyze_document", {"total_tables": 3, "total_cells_to_fill": 10})
        assert "표 3개" in desc
        assert "셀 10개" in desc

    def test_result_description_error(self) -> None:
        from src.api.websocket_handler import _build_result_description

        desc = _build_result_description("read_table", {"error": "표를 찾을 수 없습니다"})
        assert "오류" in desc
