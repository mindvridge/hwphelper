"""HWP-AI AutoFill — 윈도우 데스크톱 앱."""

import io
import logging
import os
import socket
import sys
import threading
import time
import traceback

# 콘솔 없는 환경에서 stdout/stderr 안전 처리
if sys.stdout is None or not hasattr(sys.stdout, "write"):
    sys.stdout = io.StringIO()
if sys.stderr is None or not hasattr(sys.stderr, "write"):
    sys.stderr = io.StringIO()

# 에러 로그 수집
_error_log: list[str] = []
_server_log: list[str] = []


class LogCapture(logging.Handler):
    def emit(self, record: logging.LogRecord) -> None:
        msg = self.format(record)
        _server_log.append(msg)
        if len(_server_log) > 200:
            _server_log.pop(0)


# 로그 캡처 설정
log_handler = LogCapture()
log_handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s", datefmt="%H:%M:%S"))
logging.root.addHandler(log_handler)
logging.root.setLevel(logging.INFO)


def get_base_path() -> str:
    if getattr(sys, "frozen", False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def get_exe_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def find_free_port(start: int = 8090, end: int = 8120) -> int:
    for port in range(start, end + 1):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.settimeout(1)
                s.bind(("127.0.0.1", port))
                return port
        except OSError:
            continue
    return start


def server_ready(port: int) -> bool:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(1)
            s.connect(("127.0.0.1", port))
            return True
    except OSError:
        return False


def start_server(host: str, port: int) -> None:
    try:
        import uvicorn
        uvicorn.run("src.server:app", host=host, port=port, log_level="info")
    except Exception as e:
        _error_log.append(f"서버 시작 실패: {e}\n{traceback.format_exc()}")


def show_error(msg: str) -> None:
    import ctypes
    ctypes.windll.user32.MessageBoxW(0, msg, "HWP-AI AutoFill", 0x10)


LOADING_HTML = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    background: #0f172a; color: #e2e8f0;
    font-family: 'Segoe UI', sans-serif;
    display: flex; align-items: center; justify-content: center; height: 100vh;
  }
  .container { text-align: center; }
  h1 { font-size: 24px; font-weight: 600; margin-bottom: 8px; }
  p { font-size: 14px; color: #94a3b8; margin-bottom: 32px; }
  .spinner {
    width: 40px; height: 40px; margin: 0 auto 20px;
    border: 3px solid #1e293b; border-top: 3px solid #818cf8;
    border-radius: 50%; animation: spin 0.8s linear infinite;
  }
  @keyframes spin { to { transform: rotate(360deg); } }
  #status { font-size: 13px; color: #64748b; margin-bottom: 16px; }
  #error-btn {
    display: none; margin-top: 16px; padding: 8px 20px;
    background: #dc2626; color: white; border: none; border-radius: 6px;
    font-size: 13px; cursor: pointer;
  }
  #error-btn:hover { background: #b91c1c; }
  #log-box {
    display: none; margin-top: 16px; text-align: left;
    background: #1e293b; border-radius: 8px; padding: 12px;
    font-family: 'Consolas', monospace; font-size: 11px;
    color: #94a3b8; max-height: 300px; overflow-y: auto;
    width: 600px; white-space: pre-wrap; word-break: break-all;
  }
</style>
</head>
<body>
<div class="container">
  <h1>HWP-AI AutoFill</h1>
  <p>정부과제 계획서 자동 작성 도우미</p>
  <div class="spinner" id="spinner"></div>
  <div id="status">서버를 시작하는 중...</div>
  <button id="error-btn" onclick="toggleLog()">에러 로그 보기</button>
  <div id="log-box"></div>
</div>
<script>
  function showError(msg) {
    document.getElementById('spinner').style.display = 'none';
    document.getElementById('status').textContent = msg;
    document.getElementById('status').style.color = '#f87171';
    document.getElementById('error-btn').style.display = 'inline-block';
  }
  function setLog(text) {
    document.getElementById('log-box').textContent = text;
  }
  function toggleLog() {
    var box = document.getElementById('log-box');
    box.style.display = box.style.display === 'none' ? 'block' : 'none';
  }
</script>
</body>
</html>
"""


def main() -> None:
    base = get_base_path()
    exe_dir = get_exe_dir()

    if base not in sys.path:
        sys.path.insert(0, base)

    os.chdir(exe_dir)

    os.environ.setdefault("LLM_CONFIG_PATH", os.path.join(base, "config", "llm_config.yaml"))
    os.environ.setdefault("FORMAT_RULES_PATH", os.path.join(base, "config", "format_rules.yaml"))

    env_path = os.path.join(exe_dir, ".env")
    if os.path.exists(env_path):
        from dotenv import load_dotenv
        load_dotenv(env_path)

    for d in ["uploads", "outputs"]:
        os.makedirs(os.path.join(exe_dir, d), exist_ok=True)

    preferred = int(os.environ.get("SERVER_PORT", "8090"))
    port = find_free_port(preferred, preferred + 30)

    # 서버 시작 (백그라운드)
    server_thread = threading.Thread(target=start_server, args=("127.0.0.1", port), daemon=True)
    server_thread.start()

    import webview

    window = webview.create_window(
        title="HWP-AI AutoFill",
        html=LOADING_HTML,
        width=1280,
        height=820,
        min_size=(900, 600),
        text_select=True,
    )

    def on_loaded() -> None:
        """서버 준비되면 메인 화면으로 전환, 실패 시 에러 표시."""
        for i in range(60):
            if server_ready(port):
                time.sleep(0.3)
                window.load_url(f"http://127.0.0.1:{port}")
                return
            time.sleep(0.5)

            # 에러 발생 체크
            if _error_log:
                log_text = "\\n".join(_error_log + ["", "--- 서버 로그 ---"] + _server_log[-30:])
                log_text = log_text.replace("'", "\\'").replace("\n", "\\n")
                window.evaluate_js(f'showError("서버 시작 실패"); setLog(\'{log_text}\');')
                return

        # 타임아웃
        log_text = "\\n".join(["30초 타임아웃", ""] + _server_log[-30:])
        log_text = log_text.replace("'", "\\'").replace("\n", "\\n")
        window.evaluate_js(f'showError("서버 시작 시간 초과 (30초)"); setLog(\'{log_text}\');')

    threading.Thread(target=on_loaded, daemon=True).start()
    webview.start(debug=False)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        show_error(f"오류가 발생했습니다:\n\n{traceback.format_exc()}")
