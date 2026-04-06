import { FileText, ChevronDown, ChevronRight } from "lucide-react";
import { useCallback, useEffect, useState } from "react";
import type { DocumentSchema } from "../types";
import { getSchema } from "../lib/api";
import TableSchemaView from "./TableSchemaView";

interface Props {
  sessionId: string;
  schema: DocumentSchema | null;
  onSchemaLoad?: (schema: DocumentSchema) => void;
  onCellClick?: (tableIdx: number, row: number, col: number) => void;
}

export default function DocumentPreview({
  sessionId,
  schema: externalSchema,
  onSchemaLoad,
  onCellClick,
}: Props) {
  const [schema, setSchema] = useState<DocumentSchema | null>(externalSchema);
  const [expandedTables, setExpandedTables] = useState<Set<number>>(new Set([0]));

  // 외부 스키마가 변경되면 반영
  useEffect(() => {
    if (externalSchema) setSchema(externalSchema);
  }, [externalSchema]);

  // 세션이 있으면 스키마 로드
  useEffect(() => {
    if (!sessionId || schema) return;
    getSchema(sessionId)
      .then((s) => {
        setSchema(s);
        onSchemaLoad?.(s);
      })
      .catch(() => {});
  }, [sessionId]); // eslint-disable-line react-hooks/exhaustive-deps

  // document-updated 이벤트 수신 시 리로드
  useEffect(() => {
    const handler = () => {
      if (!sessionId) return;
      getSchema(sessionId)
        .then((s) => {
          setSchema(s);
          onSchemaLoad?.(s);
        })
        .catch(() => {});
    };
    window.addEventListener("document-updated", handler);
    return () => window.removeEventListener("document-updated", handler);
  }, [sessionId, onSchemaLoad]);

  const toggleTable = useCallback((idx: number) => {
    setExpandedTables((prev) => {
      const next = new Set(prev);
      if (next.has(idx)) next.delete(idx);
      else next.add(idx);
      return next;
    });
  }, []);

  if (!sessionId) {
    return (
      <div className="h-full flex items-center justify-center text-slate-600 p-6">
        <div className="text-center">
          <FileText className="w-10 h-10 mx-auto mb-3 text-slate-700" />
          <p className="text-sm">문서를 업로드하면</p>
          <p className="text-sm">표 구조가 표시됩니다</p>
        </div>
      </div>
    );
  }

  if (!schema) {
    return (
      <div className="p-4 text-sm text-slate-500">스키마를 불러오는 중...</div>
    );
  }

  return (
    <div className="h-full flex flex-col overflow-hidden">
      {/* 헤더 */}
      <div className="px-3 py-2.5 border-b border-slate-700 flex-shrink-0">
        <p className="text-xs font-medium text-slate-400 truncate">
          {schema.document_name || "문서"}
        </p>
        <p className="text-[11px] text-slate-600 mt-0.5">
          표 {schema.total_tables}개 · 빈 셀 {schema.total_cells_to_fill}개
        </p>
      </div>

      {/* 표 목록 */}
      <div className="flex-1 overflow-y-auto">
        {schema.tables.map((table) => (
          <div key={table.table_idx} className="border-b border-slate-800">
            <button
              onClick={() => toggleTable(table.table_idx)}
              className="w-full flex items-center gap-2 px-3 py-2 text-left hover:bg-slate-800/50 transition-colors"
            >
              {expandedTables.has(table.table_idx) ? (
                <ChevronDown className="w-3.5 h-3.5 text-slate-500 flex-shrink-0" />
              ) : (
                <ChevronRight className="w-3.5 h-3.5 text-slate-500 flex-shrink-0" />
              )}
              <span className="text-xs text-slate-300">
                표 {table.table_idx + 1}
              </span>
              <span className="text-[10px] text-slate-600">
                {table.rows}x{table.cols}
              </span>
              {table.cells_to_fill > 0 && (
                <span className="ml-auto text-[10px] px-1.5 py-0.5 rounded bg-red-900/30 text-red-400">
                  {table.cells_to_fill}
                </span>
              )}
            </button>
            {expandedTables.has(table.table_idx) && (
              <div className="px-2 pb-2">
                <TableSchemaView
                  table={table}
                  onCellClick={onCellClick}
                />
              </div>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}
