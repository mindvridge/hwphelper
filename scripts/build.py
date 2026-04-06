"""프로덕션 빌드 스크립트.

사용법::

    python scripts/build.py              # 프론트엔드 빌드
    python scripts/build.py --all        # 프론트엔드 + Python 패키지
    python scripts/build.py --exe        # PyInstaller 단일 실행 파일 (선택)
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

from rich.console import Console
from rich.panel import Panel

console = Console()
ROOT = Path(__file__).resolve().parent.parent
FRONTEND = ROOT / "frontend"
DIST = FRONTEND / "dist"


def build_frontend() -> bool:
    """프론트엔드를 프로덕션 빌드한다."""
    console.print(Panel("[bold]1. 프론트엔드 빌드[/bold]", border_style="blue"))

    if not (FRONTEND / "package.json").exists():
        console.print("[red]frontend/package.json이 없습니다.[/red]")
        return False

    # npm install (node_modules 없으면)
    if not (FRONTEND / "node_modules").exists():
        console.print("  npm install 실행 중...")
        result = subprocess.run(
            ["npm", "install"],
            cwd=str(FRONTEND),
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            console.print(f"[red]npm install 실패:[/red]\n{result.stderr[:500]}")
            return False

    # npm run build
    console.print("  npm run build 실행 중...")
    result = subprocess.run(
        ["npm", "run", "build"],
        cwd=str(FRONTEND),
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        console.print(f"[red]빌드 실패:[/red]\n{result.stderr[:500]}")
        return False

    # 결과 확인
    if DIST.exists() and (DIST / "index.html").exists():
        size = sum(f.stat().st_size for f in DIST.rglob("*") if f.is_file())
        console.print(f"  [green]OK[/green] dist/ 생성 완료 ({size / 1024:.0f} KB)")
        return True
    else:
        console.print("[red]dist/index.html이 생성되지 않았습니다.[/red]")
        return False


def build_python_package() -> bool:
    """Python 패키지를 빌드한다."""
    console.print(Panel("[bold]2. Python 패키지 빌드[/bold]", border_style="blue"))

    result = subprocess.run(
        [sys.executable, "-m", "build"],
        cwd=str(ROOT),
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        console.print(f"[red]Python 빌드 실패:[/red]\n{result.stderr[:500]}")
        return False

    dist_dir = ROOT / "dist"
    if dist_dir.exists():
        wheels = list(dist_dir.glob("*.whl"))
        for w in wheels:
            console.print(f"  [green]OK[/green] {w.name} ({w.stat().st_size / 1024:.0f} KB)")
        return True
    return False


def build_exe() -> bool:
    """PyInstaller로 단일 실행 파일을 생성한다."""
    console.print(Panel("[bold]3. 실행 파일 생성 (PyInstaller)[/bold]", border_style="blue"))

    if not shutil.which("pyinstaller"):
        console.print("[yellow]PyInstaller가 설치되어 있지 않습니다.[/yellow]")
        console.print("  pip install pyinstaller 후 다시 실행하세요.")
        return False

    # frontend/dist 포함
    add_data = ""
    if DIST.exists():
        add_data = f"--add-data={DIST};frontend/dist"

    cmd = [
        "pyinstaller",
        "--onefile",
        "--name=hwp-ai",
        "--console",
        f"--add-data=config;config",
    ]
    if add_data:
        cmd.append(add_data)
    cmd.append(str(ROOT / "src" / "cli.py"))

    console.print(f"  실행: {' '.join(cmd[:5])}...")
    result = subprocess.run(cmd, cwd=str(ROOT), capture_output=True, text=True)

    if result.returncode != 0:
        console.print(f"[red]PyInstaller 실패:[/red]\n{result.stderr[:500]}")
        return False

    exe_path = ROOT / "dist" / "hwp-ai.exe"
    if exe_path.exists():
        console.print(f"  [green]OK[/green] {exe_path} ({exe_path.stat().st_size / 1024 / 1024:.1f} MB)")
        return True
    return False


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="HWP-AI AutoFill 빌드")
    parser.add_argument("--all", action="store_true", help="프론트엔드 + Python 패키지")
    parser.add_argument("--exe", action="store_true", help="PyInstaller 실행 파일")
    parser.add_argument("--frontend-only", action="store_true", help="프론트엔드만")
    args = parser.parse_args()

    success = True

    # 프론트엔드는 항상 빌드
    if not build_frontend():
        success = False

    if args.all:
        if not build_python_package():
            success = False

    if args.exe:
        if not build_exe():
            success = False

    # 요약
    console.print()
    if success:
        console.print(Panel(
            "[bold green]빌드 완료![/bold green]\n\n"
            "서버 시작: [cyan]hwp-ai serve[/cyan]\n"
            "또는: [cyan]uvicorn src.server:app --port 8080[/cyan]",
            border_style="green",
        ))
    else:
        console.print("[red]일부 빌드가 실패했습니다. 위 로그를 확인하세요.[/red]")
        sys.exit(1)


if __name__ == "__main__":
    main()
