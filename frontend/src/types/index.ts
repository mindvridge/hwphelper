// ------------------------------------------------------------------
// 채팅
// ------------------------------------------------------------------

export interface ChatMessage {
  id: string;
  role: "user" | "assistant" | "system";
  content: string;
  imageUrl?: string;
  toolCalls?: ToolCallInfo[];
  timestamp: string;
  model?: string;
  usage?: TokenUsage;
}

export interface ToolCallInfo {
  id: string;
  name: string;
  status: "running" | "success" | "error";
  description: string;
  result?: Record<string, unknown>;
}

export interface TokenUsage {
  input_tokens: number;
  output_tokens: number;
  estimated_cost_usd: number;
}

// ------------------------------------------------------------------
// 문서 스키마
// ------------------------------------------------------------------

export interface DocumentSchema {
  document_name: string;
  total_tables: number;
  total_cells_to_fill: number;
  tables: TableSchema[];
}

export interface TableSchema {
  table_idx: number;
  rows: number;
  cols: number;
  cells_to_fill: number;
  cells: CellSchema[];
}

export interface CellSchema {
  row: number;
  col: number;
  text: string;
  cell_type: "label" | "empty" | "prefilled" | "placeholder" | "unknown";
  needs_fill: boolean;
  row_span?: number;
  col_span?: number;
  context?: Record<string, string>;
}

// ------------------------------------------------------------------
// 진행률
// ------------------------------------------------------------------

export interface ProgressInfo {
  current: number;
  total: number;
  description: string;
}

// ------------------------------------------------------------------
// API 응답
// ------------------------------------------------------------------

export interface FileUploadResponse {
  session_id: string;
  file_name: string;
  tables_count: number;
  cells_to_fill: number;
  document_schema: DocumentSchema;
}

export interface ModelInfo {
  id: string;
  provider: string;
  model: string;
  description: string;
  available: boolean;
}

export interface ModelListResponse {
  models: ModelInfo[];
  default_model: string;
}

export interface HistorySnapshot {
  index: number;
  description: string;
  created_at: string;
}

export interface HistoryResponse {
  snapshots: HistorySnapshot[];
  current_idx: number;
}

export interface FormatWarning {
  location: string;
  rule: string;
  current_value: string;
  expected: string;
  auto_fixable: boolean;
}

export interface FormatReportData {
  passed: boolean;
  total_checks: number;
  passed_checks: number;
  warnings: FormatWarning[];
  errors: { location: string; rule: string; message: string }[];
  summary: string;
}

// ------------------------------------------------------------------
// WebSocket
// ------------------------------------------------------------------

export interface WSIncoming {
  type:
    | "text_delta"
    | "tool_start"
    | "tool_result"
    | "document_updated"
    | "progress"
    | "done"
    | "error"
    | "pong";
  [key: string]: unknown;
}
