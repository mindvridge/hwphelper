"""빠른 문서 채우기 — CLI 원라이너.

사용법::

    python scripts/quick_fill.py \\
        --template "양식/예비창업패키지.hwp" \\
        --program "예비창업패키지" \\
        --company "마인드브이알" \\
        --desc "AI 기반 심리상담, 모의면접 서비스" \\
        --idea "HWP 문서 AI 자동 작성 시스템" \\
        --model "claude-sonnet" \\
        --output "결과/filled.hwp"

    # 최소 옵션
    python scripts/quick_fill.py -t template.hwp -p 예비창업패키지 -c "회사명"
"""

from __future__ import annotations

import argparse
import asyncio
import sys
from pathlib import Path

from dotenv import load_dotenv


def main() -> None:
    load_dotenv()

    parser = argparse.ArgumentParser(
        description="HWP 문서 AI 자동 채우기",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("-t", "--template", required=True, help="HWP/HWPX 템플릿 파일 경로")
    parser.add_argument("-p", "--program", required=True, help="정부과제명 (예: 예비창업패키지)")
    parser.add_argument("-c", "--company", required=True, help="기업/기관명")
    parser.add_argument("-d", "--desc", default="", help="기업 소개")
    parser.add_argument("-i", "--idea", default="", help="사업 아이디어")
    parser.add_argument("-m", "--model", default="", help="LLM 모델 ID (기본: 설정 파일의 default)")
    parser.add_argument("-o", "--output", default="", help="출력 파일 경로")
    parser.add_argument("--concurrency", type=int, default=3, help="동시 LLM 호출 수")
    parser.add_argument("--validate", action="store_true", help="채우기 후 서식 검증")
    parser.add_argument("--fix", action="store_true", help="서식 위반 자동 교정")
    parser.add_argument("--dry-run", action="store_true", help="실제 쓰기 없이 생성 결과만 출력")
    args = parser.parse_args()

    # 파일 확인
    template = Path(args.template)
    if not template.exists():
        print(f"오류: 파일을 찾을 수 없습니다: {template}")
        sys.exit(1)

    asyncio.run(_run(args, template))


async def _run(args: argparse.Namespace, template: Path) -> None:
    from rich.console import Console
    from rich.panel import Panel
    from rich.progress import Progress

    console = Console()

    console.print(Panel(
        f"[bold]HWP-AI 자동 채우기[/bold]\n"
        f"템플릿: [cyan]{template.name}[/cyan]\n"
        f"과제: [cyan]{args.program}[/cyan]\n"
        f"기업: [cyan]{args.company}[/cyan]"
        + (f"\n소개: {args.desc}" if args.desc else "")
        + (f"\n아이디어: {args.idea}" if args.idea else "")
        + (f"\n모델: [cyan]{args.model or '(기본)'}[/cyan]" if True else ""),
        border_style="magenta",
    ))

    from src.ai.cell_generator import CellGenerator
    from src.ai.llm_router import LLMRouter
    from src.hwp_engine.cell_classifier import CellClassifier
    from src.hwp_engine.cell_writer import CellWriter
    from src.hwp_engine.com_controller import HwpController
    from src.hwp_engine.schema_generator import SchemaGenerator
    from src.hwp_engine.table_reader import TableReader
    from src.validator.format_checker import FormatChecker

    router = LLMRouter()
    gen = CellGenerator(llm_router=router)

    company_info = f"기업명: {args.company}"
    if args.desc:
        company_info += f"\n소개: {args.desc}"
    if args.idea:
        company_info += f"\n사업 아이디어: {args.idea}"

    with HwpController(visible=False) as ctrl:
        # ── 1단계: 분석 ─────────────────────────────────

        with console.status("[bold]문서 분석 중..."):
            ctrl.open(str(template))
            reader = TableReader(ctrl)
            classifier = CellClassifier()
            schema_gen = SchemaGenerator()

            tables = reader.read_all_tables()
            for t in tables:
                classifier.classify_table(t)
            schema = schema_gen.generate(tables, template.name)

        total = schema["total_cells_to_fill"]
        console.print(
            f"표 [bold]{schema['total_tables']}[/bold]개, "
            f"채울 셀 [bold red]{total}[/bold red]개 감지"
        )

        if total == 0:
            console.print("[yellow]채울 빈 셀이 없습니다.[/yellow]")
            return

        # ── 2단계: AI 생성 ──────────────────────────────

        with Progress(console=console) as progress:
            task = progress.add_task("[magenta]셀 생성 중...", total=total)

            async def on_progress(current: int, total_: int, cell: dict) -> None:
                label = cell.get("context", {}).get("row_label", "")
                progress.update(task, completed=current, description=f"[magenta]{label or '셀 생성'}...")

            fills = await gen.generate_all(
                schema=schema,
                program_name=args.program,
                company_info=company_info,
                model_id=args.model or None,
                concurrency=args.concurrency,
                on_progress=on_progress,
            )

        console.print(f"[green]{len(fills)}[/green]개 셀 콘텐츠 생성 완료")

        # dry-run 시 결과만 출력
        if args.dry_run:
            from rich.table import Table as RichTable

            rt = RichTable(title="생성 결과 (dry-run)")
            rt.add_column("행", width=4)
            rt.add_column("열", width=4)
            rt.add_column("내용", max_width=60)
            for f in fills:
                rt.add_row(str(f.row), str(f.col), f.text[:60])
            console.print(rt)
            return

        # ── 3단계: 셀 쓰기 ──────────────────────────────

        writer = CellWriter(ctrl)
        success_count = 0

        with console.status("[bold]문서에 내용 삽입 중..."):
            for fill in fills:
                for tbl in schema["tables"]:
                    for cell in tbl["cells"]:
                        if cell["row"] == fill.row and cell["col"] == fill.col and cell["needs_fill"]:
                            try:
                                writer.write_cell(tbl["table_idx"], fill.row, fill.col, fill.text)
                                success_count += 1
                            except Exception:
                                pass
                            break

        console.print(f"[green]{success_count}[/green]/{len(fills)} 셀 삽입 완료")

        # ── 4단계: 서식 검증 (선택) ─────────────────────

        if args.validate or args.fix:
            checker = FormatChecker()

            if args.fix:
                with console.status("서식 자동 교정 중..."):
                    fixes = checker.auto_fix(ctrl, args.program)
                console.print(f"[green]{len(fixes)}[/green]건 자동 교정")

            with console.status("서식 검증 중..."):
                report = checker.check_document(ctrl, args.program)

            status = "[green]PASS[/green]" if report.passed else "[red]FAIL[/red]"
            console.print(f"서식 검증: {status} ({report.passed_checks}/{report.total_checks})")
            if report.warnings:
                console.print(f"  경고 {len(report.warnings)}건")

        # ── 5단계: 저장 ─────────────────────────────────

        output_path = args.output or str(template.with_stem(template.stem + "_filled"))
        ctrl.save_as(output_path)

        console.print(Panel(
            f"[bold green]완료![/bold green]\n"
            f"저장: [cyan]{output_path}[/cyan]\n"
            f"생성: {len(fills)}개 | 삽입: {success_count}개",
            border_style="green",
        ))


if __name__ == "__main__":
    main()
