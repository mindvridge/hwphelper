"""Pydantic 모델 — API 요청/응답 스키마."""

from __future__ import annotations

from typing import Any

from pydantic import BaseModel, Field


# ------------------------------------------------------------------
# 파일 업로드
# ------------------------------------------------------------------


class FileUploadResponse(BaseModel):
    """문서 업로드 응답."""

    session_id: str
    file_name: str
    tables_count: int
    cells_to_fill: int
    document_schema: dict[str, Any] = {}


# ------------------------------------------------------------------
# 채팅
# ------------------------------------------------------------------


class ChatMessage(BaseModel):
    """채팅 메시지."""

    role: str  # "user" | "assistant"
    content: str
    tool_calls: list[dict[str, Any]] | None = None
    timestamp: str = ""


class ChatRequest(BaseModel):
    """채팅 요청."""

    session_id: str
    message: str
    model_id: str | None = None


class ChatResponse(BaseModel):
    """채팅 응답 (REST 폴백용)."""

    session_id: str
    message: str
    metadata: dict[str, Any] = {}


# ------------------------------------------------------------------
# 모델
# ------------------------------------------------------------------


class ModelListResponse(BaseModel):
    """사용 가능한 모델 목록."""

    models: list[dict[str, Any]]
    default_model: str


class SetDefaultModelRequest(BaseModel):
    """기본 모델 변경 요청."""

    model_id: str


# ------------------------------------------------------------------
# 세션 / 히스토리
# ------------------------------------------------------------------


class DocumentHistoryResponse(BaseModel):
    """편집 히스토리."""

    snapshots: list[dict[str, Any]]
    current_idx: int


class SessionInfoResponse(BaseModel):
    """세션 정보."""

    session_id: str
    file_name: str
    table_count: int
    snapshot_count: int
    created_at: str


# ------------------------------------------------------------------
# 서식 검증
# ------------------------------------------------------------------


class FormatCheckRequest(BaseModel):
    """서식 검증 요청."""

    program_name: str
    auto_fix: bool = False


class FormatReportResponse(BaseModel):
    """서식 검증 보고서."""

    passed: bool
    total_checks: int
    passed_checks: int
    warnings: list[dict[str, Any]]
    errors: list[dict[str, Any]] = []
    summary: str = ""


# ------------------------------------------------------------------
# 공통
# ------------------------------------------------------------------


class ErrorResponse(BaseModel):
    """에러 응답."""

    error: str
    detail: str = ""


class SuccessResponse(BaseModel):
    """성공 응답."""

    success: bool = True
    message: str = ""
    data: dict[str, Any] = {}
