"""전체 자동채우기 테스트 — WebSocket으로 서버에 요청."""
import sys
import os
import json
import shutil
import asyncio
import time

sys.stdout.reconfigure(encoding="utf-8")
os.chdir("c:/hwphelper")

import httpx
import websockets


async def main():
    base = "http://localhost:8090"
    src = "[별첨 1] 2026년 예비창업패키지 사업계획서 양식.hwp"
    dst = "test_autofill.hwp"
    shutil.copy2(src, dst)

    # 1. 파일 업로드
    print("=== 1. 파일 업로드 ===")
    async with httpx.AsyncClient(timeout=120) as c:
        with open(dst, "rb") as f:
            resp = await c.post(
                f"{base}/api/upload",
                files={"file": (dst, f, "application/octet-stream")},
            )
        data = resp.json()
        session_id = data.get("session_id", "")
        print(f"세션: {session_id}")

    if not session_id:
        print("세션 생성 실패")
        return

    # 잠시 대기 (COM 초기화)
    await asyncio.sleep(3)

    # 2. WebSocket 채팅
    print("\n=== 2. 자동채우기 시작 ===")
    ws_url = f"ws://localhost:8090/ws/chat/{session_id}"

    try:
        async with websockets.connect(ws_url, max_size=10 * 1024 * 1024) as ws:
            msg = {
                "type": "message",
                "content": (
                    "AI 기반 정부과제 사업계획서 자동작성 솔루션 DocuMind를 개발하는 "
                    "스타트업입니다. 대표자는 풀스택 개발 10년 경력이며, "
                    "사업계획서를 작성해주세요."
                ),
                "model_id": "qwen3.5-27b",
            }
            await ws.send(json.dumps(msg))
            print(f"전송: {msg['content'][:60]}...")

            start = time.time()
            text_parts = []
            event_count = 0
            tools_used = []

            while time.time() - start < 600:
                try:
                    raw = await asyncio.wait_for(ws.recv(), timeout=180)
                    data = json.loads(raw)
                    etype = data.get("type", "")
                    event_count += 1

                    if etype == "text_delta":
                        text_parts.append(data.get("content", ""))

                    elif etype == "tool_start":
                        tool = data.get("tool", "")
                        tools_used.append(tool)
                        print(f"  🔧 {tool} 시작")

                    elif etype == "tool_result":
                        tool = data.get("tool", "")
                        result = data.get("result", {})
                        result_str = json.dumps(result, ensure_ascii=False)[:120]
                        print(f"  ✅ {tool}: {result_str}")

                    elif etype == "progress":
                        desc = data.get("description", "")
                        cur = data.get("current", 0)
                        tot = data.get("total", 0)
                        print(f"  📊 [{cur}/{tot}] {desc}")

                    elif etype == "document_updated":
                        print("  📝 문서 업데이트")

                    elif etype == "done":
                        print("\n=== 완료 ===")
                        break

                    elif etype == "error":
                        print(f"  ❌ 오류: {data.get('message', '')}")
                        break

                except asyncio.TimeoutError:
                    print("  ⏰ 90초 타임아웃")
                    break

            elapsed = time.time() - start
            full_text = "".join(text_parts)
            print(f"\n소요 시간: {elapsed:.1f}초")
            print(f"총 이벤트: {event_count}")
            print(f"사용 도구: {tools_used}")
            print(f"\n--- 응답 텍스트 ---")
            print(full_text[:500] if full_text else "(텍스트 없음)")

    except Exception as e:
        print(f"WebSocket 오류: {e}")


asyncio.run(main())
