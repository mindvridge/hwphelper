"""환경 진단 스크립트 — 실행에 필요한 모든 환경을 점검한다.

사용법::

    python scripts/setup_check.py
    uv run python scripts/setup_check.py
"""

from __future__ import annotations

import os
import platform
import shutil
import socket
import subprocess
import sys
from pathlib import Path


def main() -> None:
    from rich.console import Console
    from rich.panel import Panel
    from rich.table import Table

    console = Console()
    results: list[tuple[str, bool, str]] = []  # (항목, 통과, 설명)

    console.print(Panel("[bold]HWP-AI AutoFill 환경 진단[/bold]", border_style="blue"))

    # ── 1. Python 버전 ──────────────────────────────────

    py_ver = platform.python_version()
    py_ok = sys.version_info >= (3, 11)
    results.append(("Python 버전", py_ok, f"{py_ver} {'(OK)' if py_ok else '(3.11+ 필요)'}"))

    # ── 2. OS ────────────────────────────────────────────

    os_name = platform.system()
    os_ok = os_name == "Windows"
    os_ver = platform.version()
    results.append(("운영체제", os_ok, f"{os_name} {os_ver} {'(OK)' if os_ok else '(Windows 필요)'}"))

    # ── 3. 한/글 설치 확인 ───────────────────────────────

    hwp_ok = False
    hwp_msg = "미설치"
    try:
        import win32com.client as win32  # type: ignore

        hwp = win32.Dispatch("HWPFrame.HwpObject")
        try:
            ver = hwp.Version
            hwp_msg = f"v{ver}"
        except Exception:
            hwp_msg = "버전 확인 불가"
        hwp_ok = True

        # 종료
        try:
            hwp.Clear(1)
            hwp.Quit()
        except Exception:
            pass
    except ImportError:
        hwp_msg = "pywin32 미설치"
    except Exception as e:
        hwp_msg = f"COM 오류: {e}"

    results.append(("한/글 설치", hwp_ok, hwp_msg))

    # ── 4. COM 보안 모듈 ────────────────────────────────

    sec_ok = False
    sec_msg = "확인 불가"
    if hwp_ok:
        try:
            hwp2 = win32.Dispatch("HWPFrame.HwpObject")  # type: ignore[possibly-undefined]
            hwp2.RegisterModule("FilePathCheckerModuleExample", "FilePathCheckerModule")
            sec_ok = True
            sec_msg = "등록 성공"
            try:
                hwp2.Clear(1)
                hwp2.Quit()
            except Exception:
                pass
        except Exception:
            sec_msg = "등록 실패 — 한/글 보안 설정을 확인하세요"
    else:
        sec_msg = "한/글 미설치로 건너뜀"

    results.append(("보안 모듈", sec_ok or not hwp_ok, sec_msg))

    # ── 5. pyhwpx ───────────────────────────────────────

    try:
        import pyhwpx  # noqa: F401
        results.append(("pyhwpx", True, "설치됨"))
    except ImportError:
        results.append(("pyhwpx", False, "미설치 — pip install pyhwpx"))

    # ── 6. Node.js ──────────────────────────────────────

    node_path = shutil.which("node")
    if node_path:
        try:
            ver = subprocess.check_output(["node", "--version"], text=True).strip()
            major = int(ver.lstrip("v").split(".")[0])
            node_ok = major >= 18
            results.append(("Node.js", node_ok, f"{ver} {'(OK)' if node_ok else '(18+ 필요)'}"))
        except Exception:
            results.append(("Node.js", False, "버전 확인 실패"))
    else:
        results.append(("Node.js", False, "미설치"))

    # ── 7. npm ──────────────────────────────────────────

    npm_path = shutil.which("npm")
    if npm_path:
        try:
            ver = subprocess.check_output(["npm", "--version"], text=True).strip()
            results.append(("npm", True, f"v{ver}"))
        except Exception:
            results.append(("npm", False, "버전 확인 실패"))
    else:
        results.append(("npm", False, "미설치"))

    # ── 8. uv ───────────────────────────────────────────

    uv_path = shutil.which("uv")
    if uv_path:
        try:
            ver = subprocess.check_output(["uv", "--version"], text=True).strip()
            results.append(("uv", True, ver))
        except Exception:
            results.append(("uv", True, "설치됨"))
    else:
        results.append(("uv", False, "미설치 — pip install uv"))

    # ── 9. API 키 확인 ──────────────────────────────────

    api_keys = {
        "ANTHROPIC_API_KEY": "Claude",
        "OPENAI_API_KEY": "GPT",
        "DEEPSEEK_API_KEY": "DeepSeek",
        "QWEN_API_KEY": "Qwen",
        "LOCAL_LLM_BASE_URL": "로컬 LLM",
    }
    from dotenv import load_dotenv
    load_dotenv()

    key_count = 0
    for env_var, label in api_keys.items():
        val = os.environ.get(env_var, "")
        if val:
            masked = val[:8] + "..." if len(val) > 10 else "(설정됨)"
            results.append((f"API 키: {label}", True, masked))
            key_count += 1
        else:
            results.append((f"API 키: {label}", False, "미설정"))

    # ── 10. 포트 사용 가능 여부 ─────────────────────────

    for port in [8080, 5173]:
        port_ok = _check_port(port)
        status = "사용 가능" if port_ok else "사용 중"
        results.append((f"포트 {port}", port_ok, status))

    # ── 11. 프론트엔드 빌드 ─────────────────────────────

    dist_path = Path("frontend/dist/index.html")
    results.append(("프론트엔드 빌드", dist_path.exists(), "빌드됨" if dist_path.exists() else "미빌드 — cd frontend && npm run build"))

    # ── 12. 설정 파일 ──────────────────────────────────

    for cfg in ["config/llm_config.yaml", "config/format_rules.yaml", ".env"]:
        exists = Path(cfg).exists()
        results.append((f"설정: {cfg}", exists, "존재" if exists else "없음"))

    # ── 결과 출력 ───────────────────────────────────────

    table = Table(title="환경 진단 결과")
    table.add_column("항목", width=22)
    table.add_column("상태", width=6)
    table.add_column("설명", max_width=50)

    passed = 0
    for name, ok, desc in results:
        status = "[green]PASS[/green]" if ok else "[red]FAIL[/red]"
        table.add_row(name, status, desc)
        if ok:
            passed += 1

    console.print(table)

    total = len(results)
    console.print(f"\n[bold]{passed}/{total} 통과[/bold]")

    if passed < total:
        console.print("\n[yellow]위 FAIL 항목을 해결한 후 다시 실행하세요.[/yellow]")
    else:
        console.print("\n[green]모든 환경이 준비되었습니다![/green]")

    # 핵심 항목만 체크
    critical_ok = all(ok for name, ok, _ in results if name in ("Python 버전", "운영체제"))
    if not critical_ok:
        sys.exit(1)


def _check_port(port: int) -> bool:
    """포트가 사용 가능한지 확인한다."""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(1)
            s.bind(("127.0.0.1", port))
            return True
    except OSError:
        return False


if __name__ == "__main__":
    main()
