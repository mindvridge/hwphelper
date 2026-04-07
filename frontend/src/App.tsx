import { useCallback, useState } from "react";
import ChatPanel from "./components/ChatPanel";
import ModelSelector from "./components/ModelSelector";
import { useChat } from "./hooks/useChat";
import type { FileUploadResponse } from "./types";

export default function App() {
  const [sessionId, setSessionId] = useState("");
  const chat = useChat(sessionId);

  const handleFileUpload = useCallback((data: FileUploadResponse) => {
    setSessionId(data.session_id);
  }, []);

  return (
    <div className="flex flex-col h-screen bg-slate-950 text-slate-200">
      {/* 헤더 */}
      <header className="flex items-center gap-3 px-4 py-2.5 border-b border-slate-800 bg-slate-900/80 backdrop-blur-sm flex-shrink-0">
        <h1 className="text-sm font-bold text-slate-300 tracking-tight">
          HWP-AI AutoFill
        </h1>

        <div className="flex-1" />

        <ModelSelector
          value={chat.currentModel}
          onChange={chat.setCurrentModel}
        />

        <div
          className={`w-2 h-2 rounded-full ${
            chat.isConnected ? "bg-emerald-400" : "bg-slate-600"
          }`}
          title={chat.isConnected ? "연결됨" : "연결 끊김"}
        />
      </header>

      {/* 채팅 */}
      <div className="flex-1 overflow-hidden">
        <ChatPanel
          messages={chat.messages}
          onSend={chat.sendMessage}
          onSendWithImage={chat.sendMessageWithImage}
          onLocalMessage={chat.addLocalMessage}
          onFileUpload={handleFileUpload}
          onStop={chat.stopGeneration}
          hasHwpSession={!!sessionId}
          isStreaming={chat.isStreaming}
          progress={chat.progress}
        />
      </div>
    </div>
  );
}
