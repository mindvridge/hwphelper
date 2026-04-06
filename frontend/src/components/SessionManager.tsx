import {
  Clock,
  Download,
  Redo2,
  Undo2,
  X,
} from "lucide-react";
import { useEffect, useState } from "react";
import type { HistorySnapshot } from "../types";
import {
  closeSession,
  downloadDocument,
  getHistory,
  redoAction,
  undoAction,
} from "../lib/api";

interface Props {
  sessionId: string;
  onSessionClosed?: () => void;
}

export default function SessionManager({ sessionId, onSessionClosed }: Props) {
  const [snapshots, setSnapshots] = useState<HistorySnapshot[]>([]);
  const [currentIdx, setCurrentIdx] = useState(-1);

  const loadHistory = () => {
    if (!sessionId) return;
    getHistory(sessionId)
      .then((h) => {
        setSnapshots(h.snapshots);
        setCurrentIdx(h.current_idx);
      })
      .catch(() => {});
  };

  useEffect(() => {
    loadHistory();
    window.addEventListener("document-updated", loadHistory);
    return () => window.removeEventListener("document-updated", loadHistory);
  }, [sessionId]); // eslint-disable-line react-hooks/exhaustive-deps

  const handleUndo = async () => {
    await undoAction(sessionId);
    loadHistory();
    window.dispatchEvent(new CustomEvent("document-updated"));
  };

  const handleRedo = async () => {
    await redoAction(sessionId);
    loadHistory();
    window.dispatchEvent(new CustomEvent("document-updated"));
  };

  const handleDownload = async (format: string) => {
    const blob = await downloadDocument(sessionId, format);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `document.${format}`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleClose = async () => {
    if (!confirm("세션을 종료하시겠습니까?")) return;
    await closeSession(sessionId);
    onSessionClosed?.();
  };

  if (!sessionId) return null;

  return (
    <div className="flex flex-col gap-3">
      {/* 액션 버튼 */}
      <div className="flex items-center gap-1">
        <button
          onClick={handleUndo}
          disabled={currentIdx <= 0}
          className="p-1.5 rounded-md text-slate-400 hover:text-slate-200 hover:bg-slate-800 disabled:opacity-30 transition-colors"
          title="되돌리기"
        >
          <Undo2 className="w-4 h-4" />
        </button>
        <button
          onClick={handleRedo}
          disabled={currentIdx >= snapshots.length - 1}
          className="p-1.5 rounded-md text-slate-400 hover:text-slate-200 hover:bg-slate-800 disabled:opacity-30 transition-colors"
          title="다시실행"
        >
          <Redo2 className="w-4 h-4" />
        </button>

        <div className="flex-1" />

        <button
          onClick={() => handleDownload("hwp")}
          className="p-1.5 rounded-md text-slate-400 hover:text-slate-200 hover:bg-slate-800 transition-colors"
          title="HWP 다운로드"
        >
          <Download className="w-4 h-4" />
        </button>
        <button
          onClick={handleClose}
          className="p-1.5 rounded-md text-slate-400 hover:text-red-400 hover:bg-slate-800 transition-colors"
          title="세션 종료"
        >
          <X className="w-4 h-4" />
        </button>
      </div>

      {/* 다운로드 형식 */}
      <div className="flex gap-1">
        {["hwp", "hwpx", "pdf"].map((fmt) => (
          <button
            key={fmt}
            onClick={() => handleDownload(fmt)}
            className="flex-1 text-[10px] py-1 rounded bg-slate-800 text-slate-400 hover:text-slate-200 hover:bg-slate-700 transition-colors uppercase"
          >
            {fmt}
          </button>
        ))}
      </div>

      {/* 히스토리 타임라인 */}
      {snapshots.length > 0 && (
        <div className="space-y-0.5 max-h-48 overflow-y-auto">
          <p className="text-[10px] text-slate-600 font-medium mb-1 flex items-center gap-1">
            <Clock className="w-3 h-3" /> 편집 히스토리
          </p>
          {snapshots.map((snap) => (
            <div
              key={snap.index}
              className={`text-[10px] px-2 py-1 rounded ${
                snap.index === currentIdx
                  ? "bg-indigo-900/30 text-indigo-300"
                  : "text-slate-500 hover:bg-slate-800"
              }`}
            >
              <span className="text-slate-600 mr-1">#{snap.index}</span>
              {snap.description}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
