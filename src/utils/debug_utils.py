"""디버그 유틸리티 — 문서 구조 시각화, 문서 비교, 연결 진단.

개발 중 디버깅과 문제 해결에 사용하는 도구 모음.

사용법::

    from src.utils.debug_utils import dump_table_structure, test_com_connection

    # 표 구조 Rich 출력
    dump_table_structure(hwp_ctrl, 0)

    # COM 연결 테스트
    test_com_connection()

    # LLM 연결 테스트
    await test_llm_connection("claude-sonnet")
"""

from __future__ import annotations

import asyncio
import difflib
import json
import logging
from pathlib import Path
from typing import Any

import structlog

logger = structlog.get_logger()


# ------------------------------------------------------------------
# 로깅 설정
# ------------------------------------------------------------------


def setup_logging(debug: bool = False) -> None:
    """structlog 로깅을 설정한다."""
    structlog.configure(
        processors=[
            structlog.contextvars.merge_contextvars,
            structlog.processors.add_log_level,
            structlog.processors.StackInfoRenderer(),
            structlog.dev.set_exc_info,
            structlog.processors.TimeStamper(fmt="iso"),
            structlog.dev.ConsoleRenderer() if debug else structlog.processors.JSONRenderer(),
        ],
        wrapper_class=structlog.make_filtering_bound_logger(
            logging.DEBUG if debug else logging.INFO
        ),
        context_class=dict,
        logger_factory=structlog.PrintLoggerFactory(),
        cache_logger_on_first_use=True,
    )


# ------------------------------------------------------------------
# 표 구조 시각화
# ------------------------------------------------------------------


def dump_table_structure(hwp_ctrl: Any, table_idx: int) -> None:
    """Rich로 표 구조를 시각화하여 콘솔에 출력한다."""
    from rich.console import Console
    from rich.table import Table as RichTable

    from src.hwp_engine.cell_classifier import CellClassifier
    from src.hwp_engine.table_reader import TableReader

    console = Console()
    reader = TableReader(hwp_ctrl)
    classifier = CellClassifier()

    table = reader.read_table(table_idx)
    classifier.classify_table(table)

    rt = RichTable(title=f"표 {table_idx} ({table.rows}x{table.cols})")

    # 열 헤더
    for c in range(table.cols):
        rt.add_column(f"Col {c}", width=20)

    # 그리드 구성
    grid: list[list[str]] = [[""] * table.cols for _ in range(table.rows)]
    for cell in table.cells:
        style_map = {
            "label": "[bold blue]",
            "empty": "[bold red]",
            "placeholder": "[bold yellow]",
            "prefilled": "[dim]",
        }
        prefix = style_map.get(cell.cell_type.value, "")
        suffix = "[/]" if prefix else ""
        text = cell.text[:18] or "(빈셀)"
        grid[cell.row][cell.col] = f"{prefix}[{cell.cell_type.value.upper()}] {text}{suffix}"

    for row in grid:
        rt.add_row(*row)

    console.print(rt)

    # 요약
    from src.hwp_engine.schema_generator import SchemaGenerator

    gen = SchemaGenerator()
    schema = gen.generate_table_schema(table)
    console.print(f"  채울 셀: [bold]{schema['cells_to_fill']}[/bold]개")


def dump_all_tables(hwp_ctrl: Any) -> None:
    """문서 내 모든 표를 시각화한다."""
    from src.hwp_engine.table_reader import TableReader

    reader = TableReader(hwp_ctrl)
    count = reader.get_table_count()
    for i in range(count):
        dump_table_structure(hwp_ctrl, i)


# ------------------------------------------------------------------
# 문서 비교
# ------------------------------------------------------------------


def compare_documents(path1: str, path2: str) -> str:
    """두 HWP 문서의 텍스트를 추출하여 diff를 반환한다.

    COM을 사용하여 두 문서의 텍스트를 추출하고 줄 단위로 비교한다.
    """
    from src.hwp_engine.com_controller import HwpController

    texts: list[list[str]] = []

    for path in [path1, path2]:
        with HwpController(visible=False) as ctrl:
            ctrl.open(path)
            hwp = ctrl.hwp
            hwp.HAction.Run("SelectAll")
            text = hwp.GetTextFile("TEXT", "") or ""
            texts.append(text.splitlines())

    diff = difflib.unified_diff(
        texts[0],
        texts[1],
        fromfile=Path(path1).name,
        tofile=Path(path2).name,
        lineterm="",
    )
    return "\n".join(diff)


def compare_documents_rich(path1: str, path2: str) -> None:
    """두 문서의 diff를 Rich로 출력한다."""
    from rich.console import Console
    from rich.syntax import Syntax

    console = Console()
    diff_text = compare_documents(path1, path2)
    if not diff_text:
        console.print("[green]두 문서가 동일합니다.[/green]")
        return

    syntax = Syntax(diff_text, "diff", theme="monokai")
    console.print(syntax)


# ------------------------------------------------------------------
# 연결 진단
# ------------------------------------------------------------------


def test_com_connection() -> dict[str, Any]:
    """한/글 COM 연결을 진단한다.

    Returns
    -------
    dict
        진단 결과. keys: success, version, error.
    """
    from rich.console import Console

    console = Console()

    result: dict[str, Any] = {"success": False, "version": "", "error": ""}

    # 1. pywin32 확인
    try:
        import win32com.client  # noqa: F401
        console.print("  [green]OK[/green] pywin32 설치됨")
    except ImportError:
        result["error"] = "pywin32가 설치되어 있지 않습니다."
        console.print(f"  [red]FAIL[/red] {result['error']}")
        return result

    # 2. pyhwpx 확인
    try:
        import pyhwpx  # noqa: F401
        console.print("  [green]OK[/green] pyhwpx 설치됨")
    except ImportError:
        console.print("  [yellow]WARN[/yellow] pyhwpx 미설치 (win32com 직접 사용)")

    # 3. COM 연결
    try:
        from src.hwp_engine.com_controller import HwpController

        ctrl = HwpController(visible=False)
        ctrl.connect()
        console.print("  [green]OK[/green] 한/글 COM 연결 성공")

        # 버전 확인
        try:
            version = ctrl.hwp.Version
            result["version"] = str(version)
            console.print(f"  [green]OK[/green] 한/글 버전: {version}")
        except Exception:
            console.print("  [yellow]WARN[/yellow] 버전 확인 실패")

        ctrl.quit()
        result["success"] = True

    except Exception as e:
        result["error"] = str(e)
        console.print(f"  [red]FAIL[/red] COM 연결 실패: {e}")

    return result


async def test_llm_connection(model_id: str = "") -> dict[str, Any]:
    """LLM API 연결을 테스트한다.

    Parameters
    ----------
    model_id : str
        테스트할 모델 ID. 빈 문자열이면 기본 모델.

    Returns
    -------
    dict
        테스트 결과. keys: success, model, response, error, latency_ms.
    """
    import time

    from rich.console import Console

    from src.ai.llm_router import LLMRouter

    console = Console()
    result: dict[str, Any] = {"success": False, "model": "", "response": "", "error": "", "latency_ms": 0}

    router = LLMRouter()
    mid = model_id or router.default_model
    result["model"] = mid

    console.print(f"  모델: [cyan]{mid}[/cyan]")

    try:
        start = time.time()
        response = await router.chat(
            messages=[{"role": "user", "content": "안녕하세요. 연결 테스트입니다. 한 문장으로 답해주세요."}],
            model_id=mid,
            max_tokens=50,
        )
        latency = (time.time() - start) * 1000

        from src.ai.llm_router import LLMResponse
        assert isinstance(response, LLMResponse)

        result["success"] = True
        result["response"] = response.content[:100]
        result["latency_ms"] = round(latency)
        console.print(f"  [green]OK[/green] 응답: {response.content[:60]}...")
        console.print(f"  [green]OK[/green] 지연: {latency:.0f}ms, 토큰: {response.usage.input_tokens}+{response.usage.output_tokens}")

    except Exception as e:
        result["error"] = str(e)
        console.print(f"  [red]FAIL[/red] {e}")

    return result


# ------------------------------------------------------------------
# JSON 덤프
# ------------------------------------------------------------------


def dump_table_schema(schema: dict[str, Any], filepath: str | None = None) -> str:
    """표 스키마를 JSON으로 덤프한다 (디버그용)."""
    output = json.dumps(schema, ensure_ascii=False, indent=2)
    if filepath:
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(output)
    return output


def dump_session_info(doc_manager: Any, session_id: str) -> None:
    """세션 정보를 Rich로 출력한다."""
    from rich.console import Console
    from rich.panel import Panel

    console = Console()
    try:
        session = doc_manager.get_session(session_id)
        history = doc_manager.get_history(session_id)

        info = (
            f"세션 ID: [cyan]{session.session_id}[/cyan]\n"
            f"원본: {session.original_path}\n"
            f"작업파일: {session.working_path}\n"
            f"스냅샷: {len(session.snapshots)}개 (현재: #{session.current_snapshot_idx})\n"
            f"생성: {session.created_at.isoformat()}"
        )
        console.print(Panel(info, title="세션 정보"))

        for h in history:
            marker = "→ " if h.index == session.current_snapshot_idx else "  "
            console.print(f"  {marker}#{h.index} {h.description} ({h.created_at.isoformat()})")

    except KeyError:
        console.print(f"[red]세션을 찾을 수 없습니다: {session_id}[/red]")
