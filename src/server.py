"""FastAPI 앱 엔트리포인트.

시작::

    uvicorn src.server:app --host 0.0.0.0 --port 8080 --reload

또는::

    python -m src.server
"""

from __future__ import annotations

import os
import sys
from contextlib import asynccontextmanager
from pathlib import Path

import structlog
from dotenv import load_dotenv
from fastapi import FastAPI, WebSocket
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from .ai.chat_agent import ChatAgent
from .ai.llm_router import LLMRouter
from .api.routes import router
from .api.websocket_handler import ChatWebSocketHandler
from .hwp_engine.document_manager import DocumentManager
from .utils.debug_utils import setup_logging
from .validator.format_checker import FormatChecker

logger = structlog.get_logger()

load_dotenv()


# ------------------------------------------------------------------
# 라이프사이클
# ------------------------------------------------------------------


@asynccontextmanager
async def lifespan(app: FastAPI):
    """앱 시작/종료 시 리소스를 초기화/정리한다."""
    # --- startup ---
    debug = os.environ.get("DEBUG", "false").lower() == "true"
    setup_logging(debug=debug)

    # PyInstaller exe 환경에서는 exe 폴더를 작업 디렉토리로 사용
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
        os.chdir(exe_dir)
        logger.info("작업 디렉토리 변경", cwd=exe_dir)

    # 필요한 디렉토리 생성
    for d in ["uploads", "outputs", "templates", "data"]:
        Path(d).mkdir(exist_ok=True)

    # LLM 라우터
    config_path = os.environ.get("LLM_CONFIG_PATH", "config/llm_config.yaml")
    llm_router = LLMRouter(config_path=config_path)

    # 기본 모델 오버라이드
    default_model = os.environ.get("DEFAULT_MODEL", "")
    if default_model:
        try:
            llm_router.default_model = default_model
        except ValueError:
            logger.warning("환경변수 DEFAULT_MODEL 무효", model=default_model)

    # COM 전용 스레드 (모든 COM 작업은 이 스레드에서 실행)
    import concurrent.futures
    com_executor = concurrent.futures.ThreadPoolExecutor(max_workers=1, thread_name_prefix="com")

    # 문서 매니저
    upload_dir = os.environ.get("UPLOAD_DIR", "./uploads")
    output_dir = os.environ.get("OUTPUT_DIR", "./outputs")
    doc_manager = DocumentManager(upload_dir=upload_dir, output_dir=output_dir)

    # 채팅 에이전트
    chat_agent = ChatAgent(llm_router=llm_router, doc_manager=doc_manager, com_executor=com_executor)

    # 서식 검증기
    rules_path = os.environ.get("FORMAT_RULES_PATH", "config/format_rules.yaml")
    format_checker = FormatChecker(rules_path=rules_path)

    # app.state에 공유 객체 주입
    app.state.llm_router = llm_router
    app.state.doc_manager = doc_manager
    app.state.chat_agent = chat_agent
    app.state.format_checker = format_checker
    app.state.com_executor = com_executor

    logger.info(
        "서버 초기화 완료",
        models=len(llm_router.list_models()),
        default_model=llm_router.default_model,
        programs=format_checker.available_programs,
    )

    yield

    # --- shutdown ---
    # 활성 세션 정리
    for sid in list(doc_manager.active_sessions):
        try:
            doc_manager.close_session(sid)
        except Exception:
            pass
    logger.info("서버 종료 완료")


# ------------------------------------------------------------------
# FastAPI 앱
# ------------------------------------------------------------------


def create_app() -> FastAPI:
    """FastAPI 앱을 생성한다."""
    app = FastAPI(
        title="HWP-AI AutoFill",
        description="정부지원사업 계획서 HWP 템플릿 AI 자동 채우기",
        version="1.0.0",
        lifespan=lifespan,
    )

    # CORS
    cors_origins = os.environ.get("CORS_ORIGINS", "http://localhost:5173").split(",")
    app.add_middleware(
        CORSMiddleware,
        allow_origins=[o.strip() for o in cors_origins],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    # REST API
    app.include_router(router)

    # WebSocket 채팅
    @app.websocket("/ws/chat/{session_id}")
    async def websocket_chat(websocket: WebSocket, session_id: str) -> None:
        handler = ChatWebSocketHandler(app.state.chat_agent)
        await handler.handle(websocket, session_id)

    # 프론트엔드 정적 파일 (빌드 후)
    # PyInstaller 번들 내부 경로도 검색
    frontend_candidates = [
        Path("frontend/dist"),  # 개발 환경
    ]
    if getattr(sys, "frozen", False):
        frontend_candidates.insert(0, Path(sys._MEIPASS) / "frontend" / "dist")  # type: ignore[attr-defined]

    for frontend_dist in frontend_candidates:
        if frontend_dist.exists() and (frontend_dist / "index.html").exists():
            app.mount("/", StaticFiles(directory=str(frontend_dist), html=True), name="static")
            break

    return app


app = create_app()


# ------------------------------------------------------------------
# 직접 실행
# ------------------------------------------------------------------

if __name__ == "__main__":
    import uvicorn

    host = os.environ.get("SERVER_HOST", "0.0.0.0")
    port = int(os.environ.get("SERVER_PORT", "8080"))
    uvicorn.run("src.server:app", host=host, port=port, reload=True)
