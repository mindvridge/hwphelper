import type { ProgressInfo } from "../types";

interface Props {
  progress: ProgressInfo;
}

export default function ProgressIndicator({ progress }: Props) {
  const pct = progress.total > 0 ? Math.round((progress.current / progress.total) * 100) : 0;

  return (
    <div className="mx-2 p-3 bg-slate-800 border border-slate-700 rounded-xl">
      <div className="flex justify-between text-xs text-slate-400 mb-1.5">
        <span>{progress.description || "처리 중..."}</span>
        <span>
          {progress.current}/{progress.total}
        </span>
      </div>
      <div className="w-full bg-slate-700 rounded-full h-1.5">
        <div
          className="bg-indigo-500 h-1.5 rounded-full transition-all duration-300"
          style={{ width: `${pct}%` }}
        />
      </div>
    </div>
  );
}
