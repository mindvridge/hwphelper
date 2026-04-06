import { ChevronDown, ChevronRight, CheckCircle2, XCircle, Loader2 } from "lucide-react";
import { useState } from "react";
import ReactMarkdown from "react-markdown";
import type { ChatMessage, ToolCallInfo } from "../types";

interface Props {
  message: ChatMessage;
}

/* ── 도구 실행 카드 ────────────────────────────────────── */

function ToolCard({ tool }: { tool: ToolCallInfo }) {
  const [open, setOpen] = useState(false);
  const icon =
    tool.status === "running" ? (
      <Loader2 className="w-4 h-4 animate-spin text-blue-400" />
    ) : tool.status === "success" ? (
      <CheckCircle2 className="w-4 h-4 text-emerald-400" />
    ) : (
      <XCircle className="w-4 h-4 text-red-400" />
    );

  return (
    <div className="rounded-md border border-slate-700 bg-slate-800/60 text-sm my-1.5">
      <button
        onClick={() => setOpen(!open)}
        className="flex items-center gap-2 w-full px-3 py-2 text-left hover:bg-slate-700/40 rounded-md transition-colors"
      >
        {icon}
        <span className="text-slate-300 font-mono text-xs">{tool.name}</span>
        <span className="text-slate-400 text-xs flex-1 truncate">
          {tool.description}
        </span>
        {open ? (
          <ChevronDown className="w-3.5 h-3.5 text-slate-500" />
        ) : (
          <ChevronRight className="w-3.5 h-3.5 text-slate-500" />
        )}
      </button>
      {/* 이미지 생성 결과 */}
      {tool.result?.image_url && (
        <div className="px-3 pb-2">
          <img
            src={tool.result.image_url as string}
            alt={tool.result.prompt as string || "생성된 이미지"}
            className="max-w-xs rounded-lg border border-slate-600 cursor-pointer"
            onClick={() => window.open(tool.result?.image_url as string, "_blank")}
          />
        </div>
      )}
      {open && tool.result && (
        <pre className="px-3 pb-2 text-xs text-slate-400 overflow-x-auto max-h-40">
          {JSON.stringify(tool.result, null, 2)}
        </pre>
      )}
    </div>
  );
}

/* ── 메시지 버블 ───────────────────────────────────────── */

export default function ChatMessageBubble({ message }: Props) {
  const isUser = message.role === "user";

  return (
    <div className={`flex ${isUser ? "justify-end" : "justify-start"}`}>
      <div
        className={`max-w-[85%] rounded-2xl px-4 py-3 shadow-sm ${
          isUser
            ? "bg-indigo-600 text-white"
            : "bg-slate-800 border border-slate-700 text-slate-200"
        }`}
      >
        {/* 도구 호출 */}
        {message.toolCalls && message.toolCalls.length > 0 && (
          <div className="mb-2">
            {message.toolCalls.map((tc) => (
              <ToolCard key={tc.id} tool={tc} />
            ))}
          </div>
        )}

        {/* 첨부 이미지 */}
        {message.imageUrl && (
          <img
            src={message.imageUrl}
            alt="첨부 이미지"
            className="max-w-xs rounded-lg mb-2 border border-slate-600 cursor-pointer"
            onClick={() => window.open(message.imageUrl, "_blank")}
          />
        )}

        {/* 텍스트 */}
        {message.content && (
          isUser ? (
            <p className="whitespace-pre-wrap leading-relaxed">{message.content}</p>
          ) : (
            <div className="prose prose-sm prose-invert max-w-none leading-relaxed">
              <ReactMarkdown>{message.content}</ReactMarkdown>
            </div>
          )
        )}

        {/* 메타 */}
        <div
          className={`flex items-center gap-2 text-[10px] mt-1.5 ${
            isUser ? "text-indigo-300" : "text-slate-500"
          }`}
        >
          <span>
            {new Date(message.timestamp).toLocaleTimeString("ko-KR", {
              hour: "2-digit",
              minute: "2-digit",
            })}
          </span>
          {message.usage && (
            <span>
              {message.usage.input_tokens + message.usage.output_tokens} tokens
              {message.usage.estimated_cost_usd > 0 &&
                ` / $${message.usage.estimated_cost_usd.toFixed(4)}`}
            </span>
          )}
        </div>
      </div>
    </div>
  );
}
