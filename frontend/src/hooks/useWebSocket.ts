import { useCallback, useEffect, useRef, useState } from "react";
import type { WSIncoming } from "../types";

const MAX_RETRIES = 5;
const RETRY_MS = 3000;
const PING_MS = 30_000;

export function useWebSocket(sessionId: string) {
  const wsRef = useRef<WebSocket | null>(null);
  const retries = useRef(0);
  const pingTimer = useRef<ReturnType<typeof setInterval>>();
  const [isConnected, setIsConnected] = useState(false);
  const [lastMessage, setLastMessage] = useState<WSIncoming | null>(null);

  // ── 연결 ──────────────────────────────────────────────

  const connect = useCallback(() => {
    if (!sessionId) return;
    // 이미 연결 중이면 무시
    if (
      wsRef.current &&
      (wsRef.current.readyState === WebSocket.OPEN ||
        wsRef.current.readyState === WebSocket.CONNECTING)
    )
      return;

    const proto = location.protocol === "https:" ? "wss:" : "ws:";
    const ws = new WebSocket(`${proto}//${location.host}/ws/chat/${sessionId}`);

    ws.onopen = () => {
      setIsConnected(true);
      retries.current = 0;
      // 핑/퐁 하트비트
      pingTimer.current = setInterval(() => {
        if (ws.readyState === WebSocket.OPEN) {
          ws.send(JSON.stringify({ type: "ping" }));
        }
      }, PING_MS);
    };

    ws.onclose = () => {
      setIsConnected(false);
      clearInterval(pingTimer.current);
      // 자동 재연결
      if (retries.current < MAX_RETRIES) {
        retries.current += 1;
        setTimeout(connect, RETRY_MS);
      }
    };

    ws.onerror = () => {
      ws.close();
    };

    ws.onmessage = (e) => {
      try {
        const data = JSON.parse(e.data) as WSIncoming;
        if (data.type !== "pong") {
          setLastMessage(data);
        }
      } catch {
        /* 무시 */
      }
    };

    wsRef.current = ws;
  }, [sessionId]);

  // ── 전송 ──────────────────────────────────────────────

  const send = useCallback((data: Record<string, unknown>) => {
    if (wsRef.current?.readyState === WebSocket.OPEN) {
      wsRef.current.send(JSON.stringify(data));
    }
  }, []);

  // ── 해제 ──────────────────────────────────────────────

  const disconnect = useCallback(() => {
    clearInterval(pingTimer.current);
    retries.current = MAX_RETRIES; // 재연결 방지
    wsRef.current?.close();
    wsRef.current = null;
  }, []);

  useEffect(() => {
    connect();
    return () => disconnect();
  }, [connect, disconnect]);

  return { isConnected, lastMessage, send, disconnect };
}
