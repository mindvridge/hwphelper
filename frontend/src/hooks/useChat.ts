import { useCallback, useEffect, useRef, useState } from "react";
import type { ChatMessage, ProgressInfo, ToolCallInfo, WSIncoming } from "../types";
import { useWebSocket } from "./useWebSocket";

export function useChat(sessionId: string) {
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [isStreaming, setIsStreaming] = useState(false);
  const [progress, setProgress] = useState<ProgressInfo | null>(null);
  const [currentModel, setCurrentModel] = useState<string>("");
  const streamRef = useRef<string>("");
  const toolsRef = useRef<ToolCallInfo[]>([]);
  const { isConnected, lastMessage, send } = useWebSocket(sessionId);

  // ── WebSocket 메시지 처리 ──────────────────────────────

  useEffect(() => {
    if (!lastMessage) return;
    const msg = lastMessage;

    switch (msg.type) {
      case "text_delta": {
        const delta = (msg.content as string) ?? "";
        streamRef.current += delta;
        // 마지막 assistant 메시지를 업데이트
        setMessages((prev) => {
          const last = prev[prev.length - 1];
          if (last?.role === "assistant" && last.id.startsWith("stream-")) {
            return [
              ...prev.slice(0, -1),
              { ...last, content: streamRef.current },
            ];
          }
          // 아직 assistant 메시지가 없으면 생성
          return [
            ...prev,
            {
              id: `stream-${Date.now()}`,
              role: "assistant",
              content: streamRef.current,
              timestamp: new Date().toISOString(),
              toolCalls: [],
            },
          ];
        });
        break;
      }

      case "tool_start": {
        const tool: ToolCallInfo = {
          id: `tool-${Date.now()}`,
          name: (msg.tool as string) ?? "",
          status: "running",
          description: (msg.description as string) ?? "",
        };
        toolsRef.current = [...toolsRef.current, tool];
        // assistant 메시지에 toolCalls 추가
        setMessages((prev) => {
          const last = prev[prev.length - 1];
          if (last?.role === "assistant") {
            return [
              ...prev.slice(0, -1),
              { ...last, toolCalls: [...toolsRef.current] },
            ];
          }
          return [
            ...prev,
            {
              id: `stream-${Date.now()}`,
              role: "assistant",
              content: streamRef.current,
              timestamp: new Date().toISOString(),
              toolCalls: [...toolsRef.current],
            },
          ];
        });
        break;
      }

      case "tool_result": {
        const toolName = (msg.tool as string) ?? "";
        const success = msg.success as boolean;
        const desc = (msg.description as string) ?? "";
        toolsRef.current = toolsRef.current.map((t) =>
          t.name === toolName && t.status === "running"
            ? { ...t, status: success ? "success" : "error", description: desc, result: msg.result as Record<string, unknown> }
            : t,
        );
        setMessages((prev) => {
          const last = prev[prev.length - 1];
          if (last?.role === "assistant") {
            return [
              ...prev.slice(0, -1),
              { ...last, toolCalls: [...toolsRef.current] },
            ];
          }
          return prev;
        });
        break;
      }

      case "document_updated":
        // 문서 변경 → 부모에서 스키마 재로드 트리거
        window.dispatchEvent(new CustomEvent("document-updated"));
        break;

      case "progress":
        setProgress({
          current: (msg.current as number) ?? 0,
          total: (msg.total as number) ?? 0,
          description: (msg.description as string) ?? "",
        });
        break;

      case "done": {
        setIsStreaming(false);
        setProgress(null);
        const usage = msg.usage as ChatMessage["usage"];
        // 최종 메시지에 사용량 정보 추가
        setMessages((prev) => {
          const last = prev[prev.length - 1];
          if (last?.role === "assistant") {
            return [...prev.slice(0, -1), { ...last, usage }];
          }
          return prev;
        });
        // 스트림 상태 초기화
        streamRef.current = "";
        toolsRef.current = [];
        break;
      }

      case "error":
        setIsStreaming(false);
        setMessages((prev) => [
          ...prev,
          {
            id: `err-${Date.now()}`,
            role: "assistant",
            content: `오류가 발생했습니다: ${(msg.message as string) ?? "알 수 없는 오류"}`,
            timestamp: new Date().toISOString(),
          },
        ]);
        streamRef.current = "";
        toolsRef.current = [];
        break;
    }
  }, [lastMessage]);

  // ── 메시지 전송 ──────────────────────────────────────

  const sendMessage = useCallback(
    (content: string, modelId?: string) => {
      const userMsg: ChatMessage = {
        id: `user-${Date.now()}`,
        role: "user",
        content,
        timestamp: new Date().toISOString(),
      };
      setMessages((prev) => [...prev, userMsg]);
      setIsStreaming(true);
      streamRef.current = "";
      toolsRef.current = [];

      send({
        type: "message",
        content,
        model_id: modelId ?? (currentModel || undefined),
      });
    },
    [send, currentModel],
  );

  const sendMessageWithImage = useCallback(
    (content: string, imageDataUrl: string, modelId?: string) => {
      const userMsg: ChatMessage = {
        id: `user-${Date.now()}`,
        role: "user",
        content,
        imageUrl: imageDataUrl,
        timestamp: new Date().toISOString(),
      };
      setMessages((prev) => [...prev, userMsg]);
      setIsStreaming(true);
      streamRef.current = "";
      toolsRef.current = [];

      send({
        type: "message",
        content,
        image: imageDataUrl,
        model_id: modelId ?? (currentModel || undefined),
      });
    },
    [send, currentModel],
  );

  const addLocalMessage = useCallback(
    (role: "user" | "assistant", content: string) => {
      setMessages((prev) => [
        ...prev,
        {
          id: `local-${Date.now()}`,
          role,
          content,
          timestamp: new Date().toISOString(),
        },
      ]);
    },
    [],
  );

  const clearMessages = useCallback(() => {
    setMessages([]);
    streamRef.current = "";
    toolsRef.current = [];
  }, []);

  return {
    messages,
    sendMessage,
    sendMessageWithImage,
    addLocalMessage,
    clearMessages,
    isStreaming,
    isConnected,
    progress,
    currentModel,
    setCurrentModel,
  };
}
