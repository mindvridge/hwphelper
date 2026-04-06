"""REST API 엔드포인트.

모든 엔드포인트는 /api 접두사를 사용한다.
서버 시작 시 app_state에 주입된 공유 객체(LLMRouter, DocumentManager, ChatAgent 등)를
request.app.state에서 가져와 사용한다.
"""

from __future__ import annotations

import asyncio
import os
from pathlib import Path
from typing import Any

import base64

import structlog
from fastapi import APIRouter, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse

from .schemas import (
    DocumentHistoryResponse,
    ErrorResponse,
    FileUploadResponse,
    FormatCheckRequest,
    FormatReportResponse,
    ModelListResponse,
    SetDefaultModelRequest,
    SessionInfoResponse,
    SuccessResponse,
)

logger = structlog.get_logger()
router = APIRouter(prefix="/api")


# ------------------------------------------------------------------
# 유틸리티
# ------------------------------------------------------------------


def _get_state(request: Request) -> dict[str, Any]:
    """app.state에서 공유 객체를 가져온다."""
    state = request.app.state
    return {
        "llm_router": getattr(state, "llm_router", None),
        "doc_manager": getattr(state, "doc_manager", None),
        "chat_agent": getattr(state, "chat_agent", None),
        "format_checker": getattr(state, "format_checker", None),
    }


# ------------------------------------------------------------------
# 파일 업로드
# ------------------------------------------------------------------


@router.post("/upload", response_model=FileUploadResponse)
async def upload_document(request: Request, file: UploadFile = File(...)) -> FileUploadResponse:
    """HWP/HWPX 파일을 업로드하고 세션을 생성한다.

    파일 저장 + COM 세션 생성만 수행. 표 분석은 채팅에서 수행.
    """
    import concurrent.futures

    s = _get_state(request)
    doc_mgr = s["doc_manager"]
    if doc_mgr is None:
        raise HTTPException(500, "DocumentManager가 초기화되지 않았습니다.")

    # 파일 저장
    upload_dir = Path("./uploads")
    upload_dir.mkdir(exist_ok=True)
    filename = file.filename or "unknown.hwp"
    temp_path = upload_dir / filename
    content = await file.read()
    temp_path.write_bytes(content)

    # COM 세션 생성 (전용 COM 스레드에서)
    hwp_visible = os.environ.get("HWP_VISIBLE", "false").lower() == "true"
    abs_temp = str(temp_path.resolve())
    com_executor = getattr(request.app.state, "com_executor", None)

    def _create_session() -> str:
        import pythoncom
        pythoncom.CoInitialize()
        return doc_mgr.create_session(abs_temp, visible=hwp_visible)

    try:
        loop = asyncio.get_event_loop()
        session_id = await loop.run_in_executor(com_executor, _create_session)
    except Exception as e:
        logger.exception("세션 생성 실패", file=filename)
        raise HTTPException(400, f"문서 열기 실패: {e}")

    logger.info("업로드 완료", session_id=session_id, file=filename)
    return FileUploadResponse(
        session_id=session_id,
        file_name=filename,
        tables_count=0,
        cells_to_fill=0,
        document_schema={},
    )


# ------------------------------------------------------------------
# 모델 관리
# ------------------------------------------------------------------


@router.get("/models", response_model=ModelListResponse)
async def list_models(request: Request) -> ModelListResponse:
    """사용 가능한 LLM 모델 목록을 반환한다."""
    s = _get_state(request)
    llm = s["llm_router"]
    if llm is None:
        return ModelListResponse(models=[], default_model="")

    models = llm.list_models()
    return ModelListResponse(
        models=[
            {
                "id": m.id,
                "provider": m.provider,
                "model": m.model,
                "description": m.description,
                "available": m.available,
            }
            for m in models
        ],
        default_model=llm.default_model,
    )


@router.post("/models/default", response_model=SuccessResponse)
async def set_default_model(request: Request, body: SetDefaultModelRequest) -> SuccessResponse:
    """기본 LLM 모델을 변경한다."""
    s = _get_state(request)
    llm = s["llm_router"]
    if llm is None:
        raise HTTPException(500, "LLMRouter가 초기화되지 않았습니다.")

    try:
        llm.default_model = body.model_id
    except ValueError as e:
        raise HTTPException(400, str(e))

    return SuccessResponse(message=f"기본 모델 변경: {body.model_id}")


# ------------------------------------------------------------------
# 세션 관리
# ------------------------------------------------------------------


@router.get("/sessions/{session_id}/schema")
async def get_schema(request: Request, session_id: str) -> dict[str, Any]:
    """현재 문서의 셀 스키마를 반환한다."""
    s = _get_state(request)
    doc_mgr = s["doc_manager"]
    if doc_mgr is None:
        raise HTTPException(500, "DocumentManager가 초기화되지 않았습니다.")

    try:
        session = doc_mgr.get_session(session_id)
    except KeyError:
        raise HTTPException(404, f"세션을 찾을 수 없습니다: {session_id}")

    if session.schema:
        return session.schema

    # 스키마가 없으면 재분석
    try:
        from src.hwp_engine.cell_classifier import CellClassifier
        from src.hwp_engine.schema_generator import SchemaGenerator
        from src.hwp_engine.table_reader import TableReader

        reader = TableReader(session.hwp_ctrl)
        classifier = CellClassifier()
        generator = SchemaGenerator()
        tables = reader.read_all_tables()
        for t in tables:
            classifier.classify_table(t)
        schema = generator.generate(tables, session.hwp_ctrl.file_path or "")
        session.schema = schema
        return schema
    except Exception as e:
        raise HTTPException(500, f"스키마 생성 실패: {e}")


@router.get("/sessions/{session_id}/history", response_model=DocumentHistoryResponse)
async def get_history(request: Request, session_id: str) -> DocumentHistoryResponse:
    """편집 히스토리를 반환한다."""
    s = _get_state(request)
    doc_mgr = s["doc_manager"]
    if doc_mgr is None:
        raise HTTPException(500, "DocumentManager가 초기화되지 않았습니다.")

    try:
        session = doc_mgr.get_session(session_id)
        history = doc_mgr.get_history(session_id)
    except KeyError:
        raise HTTPException(404, f"세션을 찾을 수 없습니다: {session_id}")

    return DocumentHistoryResponse(
        snapshots=[
            {"index": h.index, "description": h.description, "created_at": h.created_at.isoformat()}
            for h in history
        ],
        current_idx=session.current_snapshot_idx,
    )


@router.post("/sessions/{session_id}/undo", response_model=SuccessResponse)
async def undo(request: Request, session_id: str) -> SuccessResponse:
    """마지막 작업을 되돌린다."""
    s = _get_state(request)
    doc_mgr = s["doc_manager"]
    if doc_mgr is None:
        raise HTTPException(500, "DocumentManager가 초기화되지 않았습니다.")

    try:
        ok = doc_mgr.undo(session_id)
    except KeyError:
        raise HTTPException(404, f"세션을 찾을 수 없습니다: {session_id}")

    if not ok:
        return SuccessResponse(success=False, message="더 이상 되돌릴 수 없습니다.")
    return SuccessResponse(message="되돌리기 완료")


@router.post("/sessions/{session_id}/redo", response_model=SuccessResponse)
async def redo(request: Request, session_id: str) -> SuccessResponse:
    """되돌리기를 취소한다."""
    s = _get_state(request)
    doc_mgr = s["doc_manager"]
    if doc_mgr is None:
        raise HTTPException(500, "DocumentManager가 초기화되지 않았습니다.")

    try:
        ok = doc_mgr.redo(session_id)
    except KeyError:
        raise HTTPException(404, f"세션을 찾을 수 없습니다: {session_id}")

    if not ok:
        return SuccessResponse(success=False, message="더 이상 다시 실행할 수 없습니다.")
    return SuccessResponse(message="다시 실행 완료")


@router.get("/sessions/{session_id}/download")
async def download_document(
    request: Request, session_id: str, format: str = "hwp"
) -> FileResponse:
    """현재 문서를 다운로드한다."""
    s = _get_state(request)
    doc_mgr = s["doc_manager"]
    if doc_mgr is None:
        raise HTTPException(500, "DocumentManager가 초기화되지 않았습니다.")

    try:
        session = doc_mgr.get_session(session_id)
    except KeyError:
        raise HTTPException(404, f"세션을 찾을 수 없습니다: {session_id}")

    # 저장
    ctrl = session.hwp_ctrl
    try:
        ctrl.save()
    except Exception:
        logger.debug("다운로드 전 저장 실패 (무시)")

    if format in ("hwpx", "pdf"):
        # 변환 저장
        output_dir = Path(session.working_path).parent
        ext = f".{format}"
        export_path = output_dir / f"export{ext}"
        try:
            ctrl.save_as(str(export_path), fmt=format)
        except Exception as e:
            raise HTTPException(500, f"변환 실패: {e}")
        return FileResponse(
            str(export_path),
            media_type="application/octet-stream",
            filename=f"document{ext}",
        )

    # 기본 HWP
    return FileResponse(
        session.working_path,
        media_type="application/octet-stream",
        filename=Path(session.working_path).name,
    )


@router.get("/sessions/{session_id}/preview")
async def preview_document(request: Request, session_id: str) -> FileResponse:
    """문서를 PDF로 변환하여 미리보기용으로 반환한다."""
    s = _get_state(request)
    doc_mgr = s["doc_manager"]
    if doc_mgr is None:
        raise HTTPException(500, "DocumentManager가 초기화되지 않았습니다.")

    try:
        session = doc_mgr.get_session(session_id)
    except KeyError:
        raise HTTPException(404, f"세션을 찾을 수 없습니다: {session_id}")

    ctrl = session.hwp_ctrl
    output_dir = Path(session.working_path).parent
    pdf_path = output_dir / "preview.pdf"

    try:
        ctrl.export_pdf(str(pdf_path))
    except Exception as e:
        raise HTTPException(500, f"PDF 변환 실패: {e}")

    return FileResponse(str(pdf_path), media_type="application/pdf", filename="preview.pdf")


@router.delete("/sessions/{session_id}", response_model=SuccessResponse)
async def close_session(request: Request, session_id: str) -> SuccessResponse:
    """세션을 종료하고 리소스를 정리한다."""
    s = _get_state(request)
    doc_mgr = s["doc_manager"]
    if doc_mgr is None:
        raise HTTPException(500, "DocumentManager가 초기화되지 않았습니다.")

    try:
        final_path = doc_mgr.close_session(session_id)
    except KeyError:
        raise HTTPException(404, f"세션을 찾을 수 없습니다: {session_id}")

    # 채팅 히스토리 정리
    chat_agent = s["chat_agent"]
    if chat_agent:
        chat_agent.clear_history(session_id)

    return SuccessResponse(message="세션 종료", data={"final_path": final_path})


# ------------------------------------------------------------------
# 서식 검증
# ------------------------------------------------------------------


@router.post("/sessions/{session_id}/format-check", response_model=FormatReportResponse)
async def check_format(
    request: Request, session_id: str, body: FormatCheckRequest
) -> FormatReportResponse:
    """문서의 서식 규정 준수 여부를 검증한다."""
    s = _get_state(request)
    doc_mgr = s["doc_manager"]
    fmt_checker = s["format_checker"]

    if doc_mgr is None or fmt_checker is None:
        raise HTTPException(500, "서비스가 초기화되지 않았습니다.")

    try:
        session = doc_mgr.get_session(session_id)
    except KeyError:
        raise HTTPException(404, f"세션을 찾을 수 없습니다: {session_id}")

    try:
        if body.auto_fix:
            doc_mgr.save_snapshot(session_id, "서식 자동 수정 전")
            fixes = fmt_checker.auto_fix(session.hwp_ctrl, body.program_name)
            report = fmt_checker.check_document(session.hwp_ctrl, body.program_name)
        else:
            report = fmt_checker.check_document(session.hwp_ctrl, body.program_name)
    except Exception as e:
        raise HTTPException(500, f"서식 검증 실패: {e}")

    return FormatReportResponse(
        passed=report.passed,
        total_checks=report.total_checks,
        passed_checks=report.passed_checks,
        warnings=[
            {
                "location": w.location,
                "rule": w.rule,
                "current_value": w.current_value,
                "expected": w.expected,
                "auto_fixable": w.auto_fixable,
            }
            for w in report.warnings
        ],
        errors=[
            {"location": e.location, "rule": e.rule, "message": e.message}
            for e in report.errors
        ],
        summary=report.summary(),
    )


# ------------------------------------------------------------------
# 이미지
# ------------------------------------------------------------------


@router.post("/image/generate")
async def generate_image(request: Request, prompt: str = Form(...), size: str = Form("1024x1024")) -> dict:
    """AI로 이미지를 생성한다."""
    from src.ai.image_generator import ImageGenerator

    gen = ImageGenerator()
    result = await gen.generate(prompt=prompt, size=size)
    if result.error:
        raise HTTPException(500, f"이미지 생성 실패: {result.error}")
    return {"url": result.url, "base64": result.base64_data, "prompt": prompt}


@router.post("/reference/extract")
async def extract_reference(file: UploadFile = File(...)) -> dict:
    """참고파일에서 텍스트를 추출한다 (txt, docx, pdf, csv, md, json)."""
    filename = file.filename or ""
    content = await file.read()
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    text = ""

    try:
        if ext in ("txt", "md", "csv", "json"):
            for enc in ("utf-8", "cp949", "euc-kr"):
                try:
                    text = content.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue

        elif ext == "docx":
            import io
            import zipfile
            import xml.etree.ElementTree as ET
            with zipfile.ZipFile(io.BytesIO(content)) as z:
                if "word/document.xml" in z.namelist():
                    xml_content = z.read("word/document.xml")
                    tree = ET.fromstring(xml_content)
                    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
                    paragraphs = []
                    for p in tree.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"):
                        texts = [t.text for t in p.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t") if t.text]
                        if texts:
                            paragraphs.append("".join(texts))
                    text = "\n".join(paragraphs)

        elif ext == "pdf":
            text = "(PDF 텍스트 추출은 추후 지원 예정입니다. txt 또는 docx로 변환 후 업로드해주세요.)"

        else:
            for enc in ("utf-8", "cp949"):
                try:
                    text = content.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
    except Exception as e:
        text = f"(텍스트 추출 실패: {e})"

    # 최대 5000자
    if len(text) > 5000:
        text = text[:5000] + "\n\n...(이하 생략)..."

    return {"filename": filename, "text": text, "length": len(text)}


@router.post("/image/upload")
async def upload_image(file: UploadFile = File(...)) -> dict:
    """이미지를 업로드하고 base64로 변환한다 (LLM 비전용)."""
    content = await file.read()
    b64 = base64.b64encode(content).decode("utf-8")
    mime = file.content_type or "image/png"
    return {
        "base64": b64,
        "mime_type": mime,
        "data_url": f"data:{mime};base64,{b64}",
        "filename": file.filename,
    }
