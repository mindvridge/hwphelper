import { AlertTriangle, CheckCircle2, ShieldAlert, Wrench } from "lucide-react";
import { useState } from "react";
import type { FormatReportData } from "../types";
import { checkFormat } from "../lib/api";

interface Props {
  sessionId: string;
}

const PROGRAMS = ["예비창업패키지", "창업성장기술개발", "데이터바우처", "TIPS", "기본"];

export default function FormatReport({ sessionId }: Props) {
  const [program, setProgram] = useState("기본");
  const [report, setReport] = useState<FormatReportData | null>(null);
  const [loading, setLoading] = useState(false);

  const runCheck = async (autoFix = false) => {
    if (!sessionId) return;
    setLoading(true);
    try {
      const r = await checkFormat(sessionId, program, autoFix);
      setReport(r);
    } catch {
      /* 무시 */
    } finally {
      setLoading(false);
    }
  };

  if (!sessionId) return null;

  return (
    <div className="space-y-2">
      <p className="text-[10px] text-slate-600 font-medium flex items-center gap-1">
        <ShieldAlert className="w-3 h-3" /> 서식 검증
      </p>

      {/* 과제 선택 */}
      <select
        value={program}
        onChange={(e) => setProgram(e.target.value)}
        className="w-full text-xs bg-slate-800 border border-slate-700 rounded-md px-2 py-1.5 text-slate-300 focus:outline-none focus:ring-1 focus:ring-indigo-500"
      >
        {PROGRAMS.map((p) => (
          <option key={p} value={p}>
            {p}
          </option>
        ))}
      </select>

      {/* 버튼 */}
      <div className="flex gap-1">
        <button
          onClick={() => runCheck(false)}
          disabled={loading}
          className="flex-1 text-xs py-1.5 rounded-md bg-slate-800 text-slate-300 hover:bg-slate-700 disabled:opacity-50 transition-colors"
        >
          검사
        </button>
        <button
          onClick={() => runCheck(true)}
          disabled={loading}
          className="flex items-center gap-1 text-xs py-1.5 px-2 rounded-md bg-indigo-900/40 text-indigo-300 hover:bg-indigo-900/60 disabled:opacity-50 transition-colors"
        >
          <Wrench className="w-3 h-3" /> 자동 교정
        </button>
      </div>

      {/* 결과 */}
      {report && (
        <div className="space-y-1.5">
          {/* 배지 */}
          <div
            className={`flex items-center gap-1.5 text-xs px-2 py-1.5 rounded-md ${
              report.passed
                ? "bg-emerald-900/30 text-emerald-400"
                : "bg-red-900/30 text-red-400"
            }`}
          >
            {report.passed ? (
              <CheckCircle2 className="w-3.5 h-3.5" />
            ) : (
              <AlertTriangle className="w-3.5 h-3.5" />
            )}
            {report.passed ? "통과" : "위반 발견"} — {report.passed_checks}/
            {report.total_checks}
          </div>

          {/* 경고 목록 */}
          {report.warnings.map((w, i) => (
            <div
              key={i}
              className="text-[10px] px-2 py-1.5 rounded bg-slate-800 border-l-2 border-amber-600"
            >
              <span className="text-slate-400">{w.location}</span>
              <span className="text-slate-600 mx-1">·</span>
              <span className="text-amber-400">{w.rule}</span>
              <p className="text-slate-500 mt-0.5">
                현재: {w.current_value} → 규정: {w.expected}
              </p>
            </div>
          ))}

          {report.errors.map((e, i) => (
            <div
              key={`e-${i}`}
              className="text-[10px] px-2 py-1.5 rounded bg-slate-800 border-l-2 border-red-600"
            >
              <span className="text-slate-400">{e.location}</span>
              <span className="text-slate-600 mx-1">·</span>
              <span className="text-red-400">{e.message}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
