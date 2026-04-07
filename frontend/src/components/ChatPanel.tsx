import { Bot, Loader2 } from "lucide-react";
import { useEffect, useRef } from "react";
import type { ChatMessage, FileUploadResponse, ProgressInfo } from "../types";
import ChatInput from "./ChatInput";
import ChatMessageBubble from "./ChatMessage";
import ProgressIndicator from "./ProgressIndicator";

interface Props {
  messages: ChatMessage[];
  onSend: (message: string) => void;
  onSendWithImage?: (message: string, imageDataUrl: string) => void;
  onLocalMessage?: (role: "user" | "assistant", content: string) => void;
  onFileUpload?: (data: FileUploadResponse) => void;
  onStop?: () => void;
  hasHwpSession: boolean;
  isStreaming: boolean;
  progress: ProgressInfo | null;
}

export default function ChatPanel({
  messages,
  onSend,
  onSendWithImage,
  onLocalMessage,
  onFileUpload,
  onStop,
  hasHwpSession,
  isStreaming,
  progress,
}: Props) {
  const bottomRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, isStreaming]);

  return (
    <div className="flex flex-col h-full bg-slate-900">
      <div className="flex-1 overflow-y-auto px-4 py-6 space-y-4">
        {messages.length === 0 && (
          <div className="flex flex-col items-center justify-center h-full text-slate-500 gap-4">
            <div className="w-16 h-16 rounded-2xl bg-slate-800 flex items-center justify-center">
              <Bot className="w-8 h-8 text-indigo-400" />
            </div>
            <div className="text-center">
              <p className="text-lg font-medium text-slate-300">
                HWP-AI AutoFill
              </p>
              <p className="text-sm mt-1">
                HWP 문서를 업로드하고 대화를 시작하세요
              </p>
              <p className="text-xs mt-1 text-slate-600">
                표 구조를 분석하고 AI가 내용을 채워넣습니다
              </p>
            </div>
          </div>
        )}

        {messages.map((msg) => (
          <ChatMessageBubble key={msg.id} message={msg} />
        ))}

        {isStreaming &&
          messages.length > 0 &&
          messages[messages.length - 1]?.role !== "assistant" && (
            <div className="flex items-center gap-2 text-slate-500 px-2">
              <Loader2 className="w-4 h-4 animate-spin" />
              <span className="text-sm">응답 생성 중...</span>
            </div>
          )}

        {progress && <ProgressIndicator progress={progress} />}
        <div ref={bottomRef} />
      </div>

      <ChatInput
        onSend={onSend}
        onSendWithImage={onSendWithImage}
        onLocalMessage={onLocalMessage}
        onFileUpload={onFileUpload}
        onStop={onStop}
        hasHwpSession={hasHwpSession}
        isStreaming={isStreaming}
      />
    </div>
  );
}
