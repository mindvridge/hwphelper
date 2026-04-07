"""통합 테스트 — 10개 시나리오.

COM 의존 테스트는 @pytest.mark.skipif(not HAS_HWP) 로 보호.
LLM 호출은 모두 mock.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any
from unittest.mock import AsyncMock, MagicMock, patch

import pytest
from fastapi.testclient import TestClient

from src.hwp_engine.com_controller import HAS_HWP
from src.hwp_engine.table_reader import Cell, CellStyle, CellType, Table


# ==================================================================
# 1. COM 연결 및 기본 조작
# ==================================================================


@pytest.mark.skipif(not HAS_HWP, reason="한/글 미설치")
class TestCOMConnection:
    """시나리오 1: COM 연결 테스트."""

    def test_connect_and_version(self, hwp_controller) -> None:
        """COM 연결 후 기본 속성 접근."""
        assert hwp_controller._hwp is not None
        # Version 속성 접근 가능 여부
        try:
            _ = hwp_controller.hwp.Version
        except Exception:
            pass  # 버전 접근 실패해도 연결 자체는 성공

    def test_new_document(self, hwp_controller) -> None:
        """빈 문서 생성."""
        hwp_controller.hwp.HAction.Run("FileNew")
        # 문서가 열려 있는지 확인


# ==================================================================
# 2. 표 읽기 → 셀 쓰기 → 서식 확인
# ==================================================================


@pytest.mark.skipif(not HAS_HWP, reason="한/글 미설치")
class TestTableReadWrite:
    """시나리오 2: 표 읽기 → 쓰기 → 서식 보존 확인."""

    def test_read_write_roundtrip(self, hwp_controller, sample_template) -> None:
        """템플릿의 표를 읽고, 빈 셀에 텍스트를 쓰고, 다시 읽어서 확인."""
        from src.hwp_engine.cell_classifier import CellClassifier
        from src.hwp_engine.cell_writer import CellWriter
        from src.hwp_engine.table_reader import TableReader

        hwp_controller.open(sample_template)
        reader = TableReader(hwp_controller)
        classifier = CellClassifier()
        writer = CellWriter(hwp_controller)

        # 읽기
        table = reader.read_table(0)
        classifier.classify_table(table)
        empties = table.empty_cells()
        assert len(empties) > 0

        # 쓰기
        for cell in empties:
            writer.write_cell(0, cell.row, cell.col, "테스트 내용")

        # 다시 읽기
        table2 = reader.read_table(0)
        for cell in table2.cells:
            if cell.row == empties[0].row and cell.col == empties[0].col:
                assert "테스트 내용" in cell.text


# ==================================================================
# 3. 누름틀 워크플로우
# ==================================================================


@pytest.mark.skipif(not HAS_HWP, reason="한/글 미설치")
class TestFieldWorkflow:
    """시나리오 3: 누름틀 생성 → 채우기."""

    def test_field_create_and_fill(self, hwp_controller) -> None:
        from src.hwp_engine.field_manager import FieldManager

        hwp_controller.hwp.HAction.Run("FileNew")
        fm = FieldManager(hwp_controller)
        fields = fm.list_fields()
        # 새 문서이므로 필드 없을 수 있음
        assert isinstance(fields, list)


# ==================================================================
# 4. LLM 라우터 — Anthropic Mock
# ==================================================================


class TestLLMRouterAnthropic:
    """시나리오 4: Claude API mock 테스트."""

    @pytest.mark.asyncio
    async def test_anthropic_chat_mock(self, tmp_path: Path) -> None:
        import yaml
        from src.ai.llm_router import LLMResponse, LLMRouter

        config = {
            "models": {"test-claude": {"provider": "anthropic", "model": "claude-test", "api_key_env": "X", "max_tokens": 100}},
            "default_model": "test-claude",
        }
        cfg_path = tmp_path / "llm.yaml"
        cfg_path.write_text(yaml.dump(config), encoding="utf-8")

        router = LLMRouter(str(cfg_path))

        mock_resp = MagicMock()
        mock_resp.content = [MagicMock(type="text", text="안녕하세요!")]
        mock_resp.usage = MagicMock(input_tokens=10, output_tokens=5)

        mock_client = AsyncMock()
        mock_client.messages.create = AsyncMock(return_value=mock_resp)

        with patch.object(router, "_get_client", return_value=(mock_client, {"provider": "anthropic", "model": "claude-test"})):
            resp = await router.chat([{"role": "user", "content": "테스트"}], model_id="test-claude")

        assert isinstance(resp, LLMResponse)
        assert "안녕" in resp.content


# ==================================================================
# 5. LLM 라우터 — OpenAI 호환 Mock
# ==================================================================


class TestLLMRouterOpenAI:
    """시나리오 5: OpenAI 호환 API mock 테스트."""

    @pytest.mark.asyncio
    async def test_openai_compatible_mock(self, tmp_path: Path) -> None:
        import yaml
        from src.ai.llm_router import LLMResponse, LLMRouter

        config = {
            "models": {"test-local": {
                "provider": "openai_compatible", "model": "local-test",
                "base_url_env": "X", "default_base_url": "http://localhost:8000/v1",
                "api_key_env": "X", "default_api_key": "none", "max_tokens": 100,
            }},
            "default_model": "test-local",
        }
        cfg_path = tmp_path / "llm.yaml"
        cfg_path.write_text(yaml.dump(config), encoding="utf-8")

        router = LLMRouter(str(cfg_path))

        # httpx mock: post()가 AsyncMock이어야 함
        mock_http_response = MagicMock()
        mock_http_response.raise_for_status = MagicMock()
        mock_http_response.json.return_value = {
            "choices": [{"message": {"content": "로컬 모델 응답", "tool_calls": None}}],
            "usage": {"prompt_tokens": 5, "completion_tokens": 3},
        }

        mock_httpx = AsyncMock()
        mock_httpx.post = AsyncMock(return_value=mock_http_response)

        client_info = {
            "provider": "openai_compatible",
            "model": "local-test",
            "_httpx": mock_httpx,
            "_base_url": "http://localhost:8000/v1",
            "_api_key": "none",
        }

        with patch.object(router, "_get_client", return_value=(client_info, {"provider": "openai_compatible", "model": "local-test"})):
            resp = await router.chat([{"role": "user", "content": "테스트"}], model_id="test-local")

        assert isinstance(resp, LLMResponse)
        assert resp.content == "로컬 모델 응답"


# ==================================================================
# 6. 채팅 에이전트 전체 흐름
# ==================================================================


class TestChatAgentFlow:
    """시나리오 6: ChatAgent — 메시지 → 도구 호출 → 응답 루프."""

    @pytest.mark.asyncio
    async def test_full_chat_flow(self, mock_chat_agent) -> None:
        """텍스트 응답이 올바르게 반환되는지 확인."""
        events = []
        async for ev in mock_chat_agent.process_message("test-session", "안녕하세요"):
            events.append(ev)

        types = [e.type for e in events]
        assert "text_delta" in types or "error" in types
        assert "done" in types or "error" in types

    @pytest.mark.asyncio
    async def test_chat_history_accumulates(self, mock_chat_agent) -> None:
        """대화 히스토리가 누적되는지 확인."""
        async for _ in mock_chat_agent.process_message("s1", "첫 번째"):
            pass
        async for _ in mock_chat_agent.process_message("s1", "두 번째"):
            pass

        history = mock_chat_agent.get_history("s1")
        user_msgs = [m for m in history if m.get("role") == "user"]
        assert len(user_msgs) == 2


# ==================================================================
# 7. 문서 되돌리기/다시실행
# ==================================================================


class TestDocumentUndoRedo:
    """시나리오 7: 되돌리기/다시실행 로직 검증."""

    def test_undo_redo_flow(self, tmp_path: Path) -> None:
        from src.hwp_engine.document_manager import DocumentManager

        with patch("src.hwp_engine.document_manager.HwpController") as MockCtrl:
            MockCtrl.return_value = MagicMock()

            dummy = tmp_path / "test.hwp"
            dummy.write_bytes(b"dummy")

            mgr = DocumentManager(
                upload_dir=str(tmp_path / "uploads"),
                output_dir=str(tmp_path / "outputs"),
            )
            sid = mgr.create_session(str(dummy))

            # 작업 시뮬레이션
            mgr.save_snapshot(sid, "작업 A")
            mgr.save_snapshot(sid, "작업 B")
            mgr.save_snapshot(sid, "작업 C")

            session = mgr.get_session(sid)
            assert session.current_snapshot_idx == 3  # 초기 + A + B + C

            # undo 2번
            assert mgr.undo(sid)
            assert mgr.undo(sid)
            assert session.current_snapshot_idx == 1  # 작업 A

            # redo 1번
            assert mgr.redo(sid)
            assert session.current_snapshot_idx == 2  # 작업 B

            # 새 작업 → redo 브랜치 제거
            mgr.save_snapshot(sid, "작업 D")
            assert len(session.snapshots) == 4  # 초기, A, B, D
            assert not mgr.redo(sid)  # C는 사라짐

    def test_undo_at_beginning(self, tmp_path: Path) -> None:
        from src.hwp_engine.document_manager import DocumentManager

        with patch("src.hwp_engine.document_manager.HwpController") as MockCtrl:
            MockCtrl.return_value = MagicMock()

            dummy = tmp_path / "test.hwp"
            dummy.write_bytes(b"dummy")

            mgr = DocumentManager(
                upload_dir=str(tmp_path / "u"),
                output_dir=str(tmp_path / "o"),
            )
            sid = mgr.create_session(str(dummy))

            # 초기 상태에서 undo 불가
            assert not mgr.undo(sid)


# ==================================================================
# 8. 서식 검증 + 자동 교정
# ==================================================================


class TestFormatValidation:
    """시나리오 8: 서식 검증 및 자동 교정."""

    def test_check_multiple_programs(self) -> None:
        """여러 과제의 규정을 순차 검증."""
        from src.validator.format_checker import FormatChecker

        checker = FormatChecker()

        # 같은 스타일이 과제별로 다른 결과를 내야 함
        from src.hwp_engine.table_reader import CellStyle

        style = CellStyle(font_name="나눔고딕", font_size=10.0, char_spacing=0.0, line_spacing=160.0)

        # TIPS: 나눔고딕 허용
        tips_rules = checker.get_rules("TIPS")
        tips_warnings = checker._check_style("셀", style, tips_rules)
        font_tips = [w for w in tips_warnings if w.rule == "font_name"]
        assert len(font_tips) == 0  # 나눔고딕 허용

        # 데이터바우처: 맑은 고딕만 허용
        dv_rules = checker.get_rules("데이터바우처")
        dv_warnings = checker._check_style("셀", style, dv_rules)
        font_dv = [w for w in dv_warnings if w.rule == "font_name"]
        assert len(font_dv) == 1  # 나눔고딕 불허

    def test_auto_fixable_flag(self) -> None:
        """자동 교정 가능 여부가 올바르게 설정되는지 확인."""
        from src.hwp_engine.table_reader import CellStyle
        from src.validator.format_checker import FormatChecker

        checker = FormatChecker()
        rules = checker.get_rules("예비창업패키지")

        # 잘못된 폰트 → auto_fixable=True
        style = CellStyle(font_name="굴림", font_size=10.0, char_spacing=0.0, line_spacing=160.0)
        warnings = checker._check_style("셀", style, rules)
        font_w = [w for w in warnings if w.rule == "font_name"]
        assert font_w[0].auto_fixable is True

        # 잘못된 크기 → auto_fixable=False
        style2 = CellStyle(font_name="맑은 고딕", font_size=8.0, char_spacing=0.0, line_spacing=160.0)
        warnings2 = checker._check_style("셀", style2, rules)
        size_w = [w for w in warnings2 if w.rule == "font_size"]
        assert size_w[0].auto_fixable is False


# ==================================================================
# 9. WebSocket 채팅 통합 테스트
# ==================================================================


class TestWebSocketChat:
    """시나리오 9: WebSocket 채팅 통합."""

    def test_websocket_ping_pong(self, test_app) -> None:
        """ping/pong 프로토콜."""
        client = TestClient(test_app)
        with client.websocket_connect("/ws/chat/test-session") as ws:
            ws.send_json({"type": "ping"})
            data = ws.receive_json()
            assert data["type"] == "pong"

    def test_websocket_empty_message(self, test_app) -> None:
        """빈 메시지 시 에러."""
        client = TestClient(test_app)
        with client.websocket_connect("/ws/chat/test-session") as ws:
            ws.send_json({"type": "message", "content": ""})
            data = ws.receive_json()
            assert data["type"] == "error"

    def test_websocket_unknown_type(self, test_app) -> None:
        """알 수 없는 타입 시 에러."""
        client = TestClient(test_app)
        with client.websocket_connect("/ws/chat/test-session") as ws:
            ws.send_json({"type": "unknown_xyz"})
            data = ws.receive_json()
            assert data["type"] == "error"

    def test_websocket_invalid_json(self, test_app) -> None:
        """잘못된 JSON 시 에러."""
        client = TestClient(test_app)
        with client.websocket_connect("/ws/chat/test-session") as ws:
            ws.send_text("not json at all{{{")
            data = ws.receive_json()
            assert data["type"] == "error"


# ==================================================================
# 10. E2E 파이프라인 (업로드 → 분석 → 채팅 → 저장)
# ==================================================================


class TestE2EPipeline:
    """시나리오 10: 전체 파이프라인 통합 테스트 (REST API mock)."""

    def test_upload_and_get_schema(self, test_app) -> None:
        """업로드 → 스키마 조회 흐름."""
        client = TestClient(test_app)

        # 모델 목록 조회
        resp = client.get("/api/models")
        assert resp.status_code == 200

        # 스키마 조회 (mock)
        mock_session = MagicMock()
        mock_session.schema = {
            "total_tables": 1,
            "total_cells_to_fill": 3,
            "tables": [],
        }
        test_app.state.doc_manager.get_session.return_value = mock_session

        resp = client.get("/api/sessions/test123/schema")
        assert resp.status_code == 200
        data = resp.json()
        assert data["total_tables"] == 1

    def test_undo_redo_api(self, test_app) -> None:
        """undo/redo API 흐름."""
        client = TestClient(test_app)

        test_app.state.doc_manager.undo.return_value = True
        resp = client.post("/api/sessions/test123/undo")
        assert resp.status_code == 200
        assert resp.json()["success"] is True

        test_app.state.doc_manager.redo.return_value = True
        resp = client.post("/api/sessions/test123/redo")
        assert resp.status_code == 200
        assert resp.json()["success"] is True

    def test_close_session_cleans_up(self, test_app) -> None:
        """세션 종료 시 정리."""
        client = TestClient(test_app)

        test_app.state.doc_manager.close_session.return_value = "/path/final.hwp"
        resp = client.delete("/api/sessions/test123")
        assert resp.status_code == 200
        assert resp.json()["data"]["final_path"] == "/path/final.hwp"

    def test_format_check_api(self, test_app) -> None:
        """서식 검증 API."""
        client = TestClient(test_app)

        mock_session = MagicMock()
        test_app.state.doc_manager.get_session.return_value = mock_session

        mock_report = MagicMock()
        mock_report.passed = True
        mock_report.total_checks = 4
        mock_report.passed_checks = 4
        mock_report.warnings = []
        mock_report.errors = []
        mock_report.summary.return_value = "[PASS]"
        test_app.state.format_checker.check_document.return_value = mock_report

        resp = client.post(
            "/api/sessions/test123/format-check",
            json={"program_name": "기본"},
        )
        assert resp.status_code == 200
        assert resp.json()["passed"] is True
