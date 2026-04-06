"""CLI 엔트리포인트 — Typer 기반, 웹 UI 없이 터미널에서도 사용 가능.

사용법::

    hwp-ai serve                          # 웹 서버 시작
    hwp-ai analyze template.hwp           # 템플릿 분석
    hwp-ai generate template.hwp -p 예비창업패키지 -c "AI스타트업"
    hwp-ai validate template.hwp -p 예비창업패키지 --fix
    hwp-ai ingest past_proposal.txt -p 예비창업패키지
    hwp-ai list-fields template.hwp
    hwp-ai add-fields template.hwp --schema fields.json
    hwp-ai models
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Optional

import typer
from dotenv import load_dotenv
from rich.console import Console
from rich.panel import Panel
from rich.table import Table as RichTable

load_dotenv()

app = typer.Typer(
    name="hwp-ai",
    help="HWP-AI AutoFill — 정부과제 계획서 자동 채우기 CLI",
    no_args_is_help=True,
    rich_markup_mode="rich",
)
console = Console()


# ------------------------------------------------------------------
# 1. serve
# ------------------------------------------------------------------


@app.command()
def serve(
    host: str = typer.Option("0.0.0.0", help="서버 호스트"),
    port: int = typer.Option(8080, help="서버 포트"),
    dev: bool = typer.Option(False, help="개발 모드 (핫리로드)"),
    open_browser: bool = typer.Option(True, "--open/--no-open", help="브라우저 자동 열기"),
) -> None:
    """[bold blue]서버[/bold blue] 웹 서버를 시작합니다 (FastAPI + React)."""
    import uvicorn

    console.print(
        Panel(
            f"[bold]HWP-AI AutoFill[/bold] 서버 시작\n"
            f"주소: [cyan]http://{host if host != '0.0.0.0' else 'localhost'}:{port}[/cyan]\n"
            f"모드: {'[yellow]개발[/yellow]' if dev else '[green]운영[/green]'}",
            title="서버",
            border_style="blue",
        )
    )

    if open_browser:
        import threading
        import webbrowser

        def _open() -> None:
            import time
            time.sleep(1.5)
            webbrowser.open(f"http://localhost:{port}")

        threading.Thread(target=_open, daemon=True).start()

    uvicorn.run("src.server:app", host=host, port=port, reload=dev)


# ------------------------------------------------------------------
# 2. analyze
# ------------------------------------------------------------------


@app.command()
def analyze(
    file: str = typer.Argument(..., help="HWP/HWPX 파일 경로"),
) -> None:
    """[bold green]분석[/bold green] HWP 템플릿의 표 구조를 분석합니다."""
    _check_file(file)

    from src.hwp_engine.cell_classifier import CellClassifier
    from src.hwp_engine.com_controller import HwpController
    from src.hwp_engine.schema_generator import SchemaGenerator
    from src.hwp_engine.table_reader import TableReader

    with console.status("문서를 분석하고 있습니다..."):
        with HwpController(visible=False) as ctrl:
            ctrl.open(file)
            reader = TableReader(ctrl)
            classifier = CellClassifier()
            generator = SchemaGenerator()

            tables = reader.read_all_tables()
            for t in tables:
                classifier.classify_table(t)

            schema = generator.generate(tables, Path(file).name)

    # 결과 표시
    console.print(Panel(
        f"파일: [cyan]{Path(file).name}[/cyan]\n"
        f"표 개수: [bold]{schema['total_tables']}[/bold]\n"
        f"채울 셀: [bold red]{schema['total_cells_to_fill']}[/bold red]",
        title="분석 결과",
        border_style="green",
    ))

    for tbl in schema["tables"]:
        rt = RichTable(title=f"표 {tbl['table_idx']} ({tbl['rows']}x{tbl['cols']})")
        rt.add_column("행", style="dim", width=4)
        rt.add_column("열", style="dim", width=4)
        rt.add_column("타입", width=12)
        rt.add_column("텍스트", max_width=50)

        for cell in tbl["cells"]:
            type_style = {
                "label": "[blue]LABEL[/blue]",
                "empty": "[red]EMPTY[/red]",
                "placeholder": "[yellow]PLACEHOLDER[/yellow]",
                "prefilled": "[dim]PREFILLED[/dim]",
            }.get(cell["cell_type"], cell["cell_type"])

            rt.add_row(
                str(cell["row"]),
                str(cell["col"]),
                type_style,
                cell["text"][:50] or "-",
            )

        console.print(rt)


# ------------------------------------------------------------------
# 3. generate
# ------------------------------------------------------------------


@app.command()
def generate(
    file: str = typer.Argument(..., help="HWP/HWPX 파일 경로"),
    program: str = typer.Option(..., "-p", "--program", help="정부과제명"),
    company: str = typer.Option(..., "-c", "--company", help="기업/기관명"),
    desc: str = typer.Option("", "-d", "--desc", help="기업 소개"),
    idea: str = typer.Option("", "-i", "--idea", help="사업 아이디어"),
    model: str = typer.Option("", "-m", "--model", help="LLM 모델 ID"),
    output: str = typer.Option("", "-o", "--output", help="출력 파일 경로"),
    concurrency: int = typer.Option(3, help="동시 LLM 호출 수"),
) -> None:
    """[bold magenta]생성[/bold magenta] AI로 빈 셀을 자동 채웁니다."""
    import asyncio

    _check_file(file)

    async def _run() -> None:
        from src.ai.cell_generator import CellGenerator
        from src.ai.llm_router import LLMRouter
        from src.hwp_engine.cell_classifier import CellClassifier
        from src.hwp_engine.cell_writer import CellWriter
        from src.hwp_engine.com_controller import HwpController
        from src.hwp_engine.schema_generator import SchemaGenerator
        from src.hwp_engine.table_reader import TableReader

        router = LLMRouter()
        gen = CellGenerator(llm_router=router)
        company_info = f"기업명: {company}"
        if desc:
            company_info += f"\n소개: {desc}"
        if idea:
            company_info += f"\n사업 아이디어: {idea}"

        with HwpController(visible=False) as ctrl:
            ctrl.open(file)
            reader = TableReader(ctrl)
            classifier = CellClassifier()
            generator_obj = SchemaGenerator()
            writer = CellWriter(ctrl)

            # 분석
            with console.status("문서 분석 중..."):
                tables = reader.read_all_tables()
                for t in tables:
                    classifier.classify_table(t)
                schema = generator_obj.generate(tables, Path(file).name)

            total = schema["total_cells_to_fill"]
            if total == 0:
                console.print("[yellow]채울 빈 셀이 없습니다.[/yellow]")
                return

            console.print(f"채울 셀: [bold]{total}[/bold]개")

            # 생성
            from rich.progress import Progress

            with Progress(console=console) as progress:
                task = progress.add_task("셀 생성 중...", total=total)

                async def on_progress(current: int, total_: int, cell: dict) -> None:
                    progress.update(task, completed=current)

                fills = await gen.generate_all(
                    schema=schema,
                    program_name=program,
                    company_info=company_info,
                    model_id=model or None,
                    concurrency=concurrency,
                    on_progress=on_progress,
                )

            # 쓰기
            with console.status("셀에 내용을 삽입하고 있습니다..."):
                success = 0
                for fill in fills:
                    # 어느 표에 속하는지 찾기
                    for tbl in schema["tables"]:
                        for cell in tbl["cells"]:
                            if cell["row"] == fill.row and cell["col"] == fill.col and cell["needs_fill"]:
                                try:
                                    writer.write_cell(tbl["table_idx"], fill.row, fill.col, fill.text)
                                    success += 1
                                except Exception:
                                    pass
                                break

            # 저장
            out_path = output or str(Path(file).with_stem(Path(file).stem + "_filled"))
            ctrl.save_as(out_path)

            console.print(Panel(
                f"생성: [bold green]{len(fills)}[/bold green]개\n"
                f"삽입 성공: [bold green]{success}[/bold green]개\n"
                f"저장: [cyan]{out_path}[/cyan]",
                title="생성 완료",
                border_style="magenta",
            ))

    asyncio.run(_run())


# ------------------------------------------------------------------
# 4. validate
# ------------------------------------------------------------------


@app.command()
def validate(
    file: str = typer.Argument(..., help="HWP/HWPX 파일 경로"),
    program: str = typer.Option("기본", "-p", "--program", help="정부과제명"),
    fix: bool = typer.Option(False, "--fix", help="위반 사항 자동 수정"),
    output: str = typer.Option("", "-o", "--output", help="수정 결과 저장 경로"),
) -> None:
    """[bold yellow]검증[/bold yellow] 서식 규정 준수 여부를 검사합니다."""
    _check_file(file)

    from src.hwp_engine.com_controller import HwpController
    from src.validator.format_checker import FormatChecker

    checker = FormatChecker()

    with HwpController(visible=False) as ctrl:
        ctrl.open(file)

        if fix:
            with console.status("서식 자동 수정 중..."):
                fixes = checker.auto_fix(ctrl, program)
            console.print(f"[green]자동 수정 {len(fixes)}건 적용[/green]")
            for f in fixes:
                status = "[green]OK[/green]" if f.success else "[red]실패[/red]"
                console.print(f"  {status} {f.location}: {f.rule} {f.old_value} → {f.new_value}")

        with console.status("서식 검증 중..."):
            report = checker.check_document(ctrl, program)

        # 결과 표시
        status_badge = "[bold green]PASS[/bold green]" if report.passed else "[bold red]FAIL[/bold red]"
        console.print(Panel(
            f"상태: {status_badge}\n"
            f"검사: {report.passed_checks}/{report.total_checks} 통과\n"
            f"경고: [yellow]{len(report.warnings)}[/yellow]건\n"
            f"에러: [red]{len(report.errors)}[/red]건",
            title=f"서식 검증 — {program}",
            border_style="yellow",
        ))

        if report.warnings:
            rt = RichTable(title="경고 목록")
            rt.add_column("위치", width=15)
            rt.add_column("규칙", width=15)
            rt.add_column("현재값", width=15)
            rt.add_column("규정값", width=20)
            rt.add_column("자동수정", width=8)
            for w in report.warnings:
                rt.add_row(
                    w.location, w.rule, w.current_value, w.expected,
                    "[green]O[/green]" if w.auto_fixable else "[red]X[/red]",
                )
            console.print(rt)

        if fix and output:
            ctrl.save_as(output)
            console.print(f"수정 결과 저장: [cyan]{output}[/cyan]")


# ------------------------------------------------------------------
# 5. ingest
# ------------------------------------------------------------------


@app.command()
def ingest(
    file: str = typer.Argument(..., help="텍스트 파일 경로 (.txt)"),
    program: str = typer.Option("", "-p", "--program", help="정부과제명"),
) -> None:
    """[bold cyan]색인[/bold cyan] 과거 계획서를 RAG 벡터 DB에 등록합니다."""
    _check_file(file)

    from src.ai.rag_engine import RAGEngine

    engine = RAGEngine()
    with console.status("문서를 색인하고 있습니다..."):
        count = engine.ingest_document(file, program_name=program)

    console.print(f"[green]색인 완료[/green]: {count}개 청크 등록 (과제: {program or '미지정'})")


# ------------------------------------------------------------------
# 6. list-fields
# ------------------------------------------------------------------


@app.command("list-fields")
def list_fields(
    file: str = typer.Argument(..., help="HWP/HWPX 파일 경로"),
) -> None:
    """[bold blue]필드[/bold blue] 누름틀 필드 목록을 확인합니다."""
    _check_file(file)

    from src.hwp_engine.com_controller import HwpController
    from src.hwp_engine.field_manager import FieldManager

    with HwpController(visible=False) as ctrl:
        ctrl.open(file)
        fm = FieldManager(ctrl)
        fields = fm.list_fields()

    if not fields:
        console.print("[yellow]누름틀 필드가 없습니다.[/yellow]")
        return

    rt = RichTable(title=f"누름틀 필드 ({len(fields)}개)")
    rt.add_column("#", style="dim", width=4)
    rt.add_column("필드명", style="cyan")
    rt.add_column("현재 값", max_width=50)

    for i, f in enumerate(fields, 1):
        rt.add_row(str(i), f.name, f.value or "[dim]-[/dim]")

    console.print(rt)


# ------------------------------------------------------------------
# 7. add-fields
# ------------------------------------------------------------------


@app.command("add-fields")
def add_fields(
    file: str = typer.Argument(..., help="HWP/HWPX 파일 경로"),
    schema: str = typer.Option(..., "--schema", help="필드 매핑 JSON 파일 경로"),
    output: str = typer.Option("", "-o", "--output", help="출력 파일 경로"),
) -> None:
    """[bold blue]필드 생성[/bold blue] JSON 스키마로 누름틀 필드를 자동 생성합니다.

    JSON 형식::

        {"table_idx": 0, "fields": {"사업명": [0, 1], "기관명": [1, 1]}}
    """
    _check_file(file)
    _check_file(schema)

    from src.hwp_engine.com_controller import HwpController
    from src.hwp_engine.field_manager import FieldManager

    with open(schema, encoding="utf-8") as f:
        mapping_data = json.load(f)

    table_idx = mapping_data.get("table_idx", 0)
    field_mapping: dict[str, tuple[int, int]] = {}
    for name, pos in mapping_data.get("fields", {}).items():
        field_mapping[name] = (pos[0], pos[1])

    with HwpController(visible=False) as ctrl:
        ctrl.open(file)
        fm = FieldManager(ctrl)

        with console.status("필드를 생성하고 있습니다..."):
            fm.create_field_template(table_idx, field_mapping)

        out_path = output or str(Path(file).with_stem(Path(file).stem + "_fields"))
        ctrl.save_as(out_path)

    console.print(f"[green]필드 {len(field_mapping)}개 생성 완료[/green]: [cyan]{out_path}[/cyan]")


# ------------------------------------------------------------------
# 8. models
# ------------------------------------------------------------------


@app.command()
def models() -> None:
    """[bold]모델[/bold] 사용 가능한 LLM 모델 목록을 표시합니다."""
    from src.ai.llm_router import LLMRouter

    router = LLMRouter()
    model_list = router.list_models()

    rt = RichTable(title="사용 가능한 LLM 모델")
    rt.add_column("ID", style="cyan")
    rt.add_column("프로바이더", width=12)
    rt.add_column("모델")
    rt.add_column("상태", width=8)
    rt.add_column("설명", max_width=40)
    rt.add_column("비용/1K", width=10)

    for m in model_list:
        status = "[green]활성[/green]" if m.available else "[red]미설정[/red]"
        cost = f"${m.estimated_cost_per_1k:.4f}" if m.estimated_cost_per_1k else "-"
        rt.add_row(m.id, m.provider, m.model, status, m.description, cost)

    console.print(rt)
    console.print(f"\n기본 모델: [bold cyan]{router.default_model}[/bold cyan]")


# ------------------------------------------------------------------
# 유틸리티
# ------------------------------------------------------------------


def _check_file(path: str) -> None:
    """파일 존재 여부를 확인한다."""
    if not Path(path).exists():
        console.print(f"[red]파일을 찾을 수 없습니다: {path}[/red]")
        raise typer.Exit(1)


# ------------------------------------------------------------------
# 엔트리포인트
# ------------------------------------------------------------------


def main() -> None:
    """CLI 메인 함수."""
    app()


if __name__ == "__main__":
    main()
