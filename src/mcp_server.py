"""MCP 서버 — Claude Desktop / Cursor에서 HWP 문서를 직접 편집."""

from __future__ import annotations

import concurrent.futures
import contextlib
import io
import json
import logging
import os
import sys
from pathlib import Path
from typing import Any

from dotenv import load_dotenv

load_dotenv()

# structlog 로그를 stderr로, WARNING 이상만
import structlog

logging.basicConfig(stream=sys.stderr, level=logging.WARNING, format="%(message)s")
structlog.configure(
    processors=[structlog.dev.ConsoleRenderer()],
    wrapper_class=structlog.make_filtering_bound_logger(logging.WARNING),
    logger_factory=structlog.PrintLoggerFactory(file=sys.stderr),
)

from mcp.server.fastmcp import FastMCP

# ------------------------------------------------------------------
# MCP 서버
# ------------------------------------------------------------------

mcp = FastMCP(
    "hwp-ai",
    instructions=(
        "한/글(HWP) 문서 자동화 도구입니다.\n"
        "1. analyze_template: HWP 파일 열기 + 표 분석\n"
        "2. write_cell: 셀에 내용 쓰기\n"
        "3. save_document: 저장\n"
        "첫 호출 시 한/글 시작에 5~10초 소요됩니다."
    ),
)

# ------------------------------------------------------------------
# 상태
# ------------------------------------------------------------------

_ctrl: Any = None
_file_path: str | None = None
_tables_cache: list[Any] | None = None
_schema_cache: dict[str, Any] | None = None
_executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)


@contextlib.contextmanager
def _suppress_stdout():
    """stdout 출력을 일시적으로 차단 (pyhwpx 잡음 방지)."""
    old = sys.stdout
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
    try:
        yield
    finally:
        sys.stdout.close()
        sys.stdout = old


def _run_com(func: Any, *args: Any) -> Any:
    """COM 작업을 전용 스레드에서 실행."""
    future = _executor.submit(func, *args)
    return future.result(timeout=120)


def _ensure_ctrl() -> Any:
    global _ctrl
    if _ctrl is None:
        import pythoncom
        pythoncom.CoInitialize()

        with _suppress_stdout():
            from src.hwp_engine.com_controller import HwpController
            visible = os.environ.get("HWP_VISIBLE", "true").lower() == "true"
            _ctrl = HwpController(visible=visible)
            _ctrl.connect()
    return _ctrl


def _invalidate_cache() -> None:
    global _tables_cache, _schema_cache
    _tables_cache = None
    _schema_cache = None


# ------------------------------------------------------------------
# COM 작업 함수들
# ------------------------------------------------------------------

def _do_analyze(file_path: str) -> dict:
    global _file_path, _tables_cache, _schema_cache
    with _suppress_stdout():
        from src.hwp_engine.cell_classifier import CellClassifier
        from src.hwp_engine.schema_generator import SchemaGenerator
        from src.hwp_engine.table_reader import TableReader

        ctrl = _ensure_ctrl()
        abs_path = str(Path(file_path).resolve())
        if _file_path == abs_path and _schema_cache:
            return _schema_cache

        ctrl.open(abs_path)
        _file_path = abs_path
        _invalidate_cache()

        reader = TableReader(ctrl)
        classifier = CellClassifier()
        generator = SchemaGenerator()
        tables = reader.read_all_tables()
        for t in tables:
            classifier.classify_table(t)
        _tables_cache = tables
        _schema_cache = generator.generate(tables, Path(abs_path).name)

    return {
        "document": _schema_cache["document_name"],
        "total_tables": _schema_cache["total_tables"],
        "total_cells_to_fill": _schema_cache["total_cells_to_fill"],
        "tables": [
            {"table_idx": t["table_idx"], "size": f"{t['rows']}x{t['cols']}", "cells_to_fill": t["cells_to_fill"]}
            for t in _schema_cache["tables"]
        ],
    }


def _do_read_table(table_idx: int) -> dict:
    with _suppress_stdout():
        from src.hwp_engine.cell_classifier import CellClassifier
        from src.hwp_engine.table_reader import TableReader
        ctrl = _ensure_ctrl()
        reader = TableReader(ctrl)
        classifier = CellClassifier()
        table = reader.read_table(table_idx)
        classifier.classify_table(table)
    return table.to_dict()


def _do_read_cell(table_idx: int, row: int, col: int) -> dict:
    with _suppress_stdout():
        from src.hwp_engine.table_reader import TableReader
        ctrl = _ensure_ctrl()
        reader = TableReader(ctrl)
        table = reader.read_table(table_idx)
        cell = table.get_cell(row, col)
        if cell is None:
            return {"error": f"셀 ({row},{col}) 없음"}
        style = reader.read_cell_style(table_idx, row, col)
    return {"row": cell.row, "col": cell.col, "text": cell.text, "cell_type": cell.cell_type.value,
            "style": {"font_name": style.font_name, "font_size": style.font_size, "bold": style.bold}}


def _do_write_cell(table_idx: int, row: int, col: int, text: str) -> dict:
    with _suppress_stdout():
        from src.hwp_engine.cell_writer import CellWriter
        ctrl = _ensure_ctrl()
        CellWriter(ctrl).write_cell(table_idx, row, col, text)
        _invalidate_cache()
    return {"success": True, "message": f"표{table_idx} ({row},{col}) 작성 완료"}


def _do_fill_field(field_name: str, text: str) -> dict:
    with _suppress_stdout():
        from src.hwp_engine.field_manager import FieldManager
        ctrl = _ensure_ctrl()
        ok = FieldManager(ctrl).fill_field(field_name, text)
    return {"success": ok, "field": field_name}


def _do_save(file_path: str, fmt: str) -> dict:
    with _suppress_stdout():
        ctrl = _ensure_ctrl()
        if file_path:
            ctrl.save_as(file_path, fmt=fmt)
            return {"success": True, "path": file_path}
        elif fmt != "hwp" and _file_path:
            p = str(Path(_file_path).with_suffix(f".{fmt}"))
            ctrl.save_as(p, fmt=fmt)
            return {"success": True, "path": p}
        else:
            ctrl.save()
            return {"success": True, "path": _file_path}


def _do_validate(program_name: str, auto_fix: bool) -> dict:
    with _suppress_stdout():
        from src.validator.format_checker import FormatChecker
        ctrl = _ensure_ctrl()
        checker = FormatChecker(os.environ.get("FORMAT_RULES_PATH", "config/format_rules.yaml"))
        if auto_fix:
            checker.auto_fix(ctrl, program_name)
        report = checker.check_document(ctrl, program_name)
    return {"passed": report.passed, "summary": report.summary(),
            "warnings": [{"location": w.location, "rule": w.rule, "current": w.current_value, "expected": w.expected} for w in report.warnings]}


# ------------------------------------------------------------------
# MCP 도구
# ------------------------------------------------------------------

@mcp.tool()
def analyze_template(file_path: str) -> str:
    """HWP 파일을 열고 표 구조를 분석합니다. 첫 호출 시 한/글 시작에 5~10초 걸립니다.

    Args:
        file_path: HWP 파일의 절대 경로 (예: C:\\Users\\pc\\문서\\양식.hwp)
    """
    try:
        return json.dumps(_run_com(_do_analyze, file_path), ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


@mcp.tool()
def read_table(table_idx: int) -> str:
    """특정 표의 전체 내용을 읽습니다.

    Args:
        table_idx: 표 번호 (0부터)
    """
    try:
        return json.dumps(_run_com(_do_read_table, table_idx), ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


@mcp.tool()
def read_cell(table_idx: int, row: int, col: int) -> str:
    """특정 셀의 텍스트와 서식을 읽습니다.

    Args:
        table_idx: 표 번호
        row: 행
        col: 열
    """
    try:
        return json.dumps(_run_com(_do_read_cell, table_idx, row, col), ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


@mcp.tool()
def write_cell(table_idx: int, row: int, col: int, text: str) -> str:
    """셀에 텍스트를 삽입합니다 (서식 유지).

    Args:
        table_idx: 표 번호
        row: 행
        col: 열
        text: 삽입할 텍스트
    """
    try:
        return json.dumps(_run_com(_do_write_cell, table_idx, row, col, text), ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


@mcp.tool()
def fill_field(field_name: str, text: str) -> str:
    """누름틀 필드에 텍스트를 삽입합니다.

    Args:
        field_name: 필드 이름
        text: 텍스트
    """
    try:
        return json.dumps(_run_com(_do_fill_field, field_name, text), ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


@mcp.tool()
def fill_all_empty_cells(program_name: str, company_name: str, company_desc: str = "", business_idea: str = "") -> str:
    """빈 셀 목록과 컨텍스트를 반환합니다. write_cell로 하나씩 채우세요.

    Args:
        program_name: 정부과제명
        company_name: 기업명
        company_desc: 기업 소개
        business_idea: 사업 아이디어
    """
    if not _schema_cache:
        return json.dumps({"error": "먼저 analyze_template으로 문서를 열어주세요."}, ensure_ascii=False)
    cells = []
    for t in _schema_cache.get("tables", []):
        for c in t["cells"]:
            if c.get("needs_fill"):
                cells.append({"table_idx": t["table_idx"], "row": c["row"], "col": c["col"], "context": c.get("context", {})})
    return json.dumps({"total": len(cells), "program": program_name, "company": company_name, "cells": cells,
                        "instruction": "위 셀의 context를 참고하여 write_cell로 채워주세요."}, ensure_ascii=False, indent=2)


@mcp.tool()
def validate_format(program_name: str, auto_fix: bool = False) -> str:
    """서식 규정을 검증합니다.

    Args:
        program_name: 정부과제명
        auto_fix: 자동 수정 여부
    """
    try:
        return json.dumps(_run_com(_do_validate, program_name, auto_fix), ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


@mcp.tool()
def save_document(file_path: str = "", format: str = "hwp") -> str:
    """문서를 저장합니다.

    Args:
        file_path: 저장 경로 (빈 문자열이면 원본)
        format: hwp, hwpx, pdf
    """
    try:
        return json.dumps(_run_com(_do_save, file_path, format), ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


# ------------------------------------------------------------------

if __name__ == "__main__":
    mcp.run()
