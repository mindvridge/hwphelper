import type {
  DocumentSchema,
  FileUploadResponse,
  FormatReportData,
  HistoryResponse,
  ModelListResponse,
} from "../types";

const API = "/api";

async function request<T>(path: string, opts?: RequestInit): Promise<T> {
  const res = await fetch(`${API}${path}`, {
    headers: { "Content-Type": "application/json" },
    ...opts,
  });
  if (!res.ok) {
    const body = await res.json().catch(() => ({ error: res.statusText }));
    throw new Error(body.error ?? body.detail ?? res.statusText);
  }
  return res.json();
}

// ── 파일 ────────────────────────────────────────────────

export async function uploadDocument(file: File): Promise<FileUploadResponse> {
  const fd = new FormData();
  fd.append("file", file);
  const res = await fetch(`${API}/upload`, { method: "POST", body: fd });
  if (!res.ok) {
    const body = await res.json().catch(() => ({ detail: res.statusText }));
    throw new Error(body.detail ?? body.error ?? "업로드 실패");
  }
  return res.json();
}

export async function downloadDocument(
  sessionId: string,
  format = "hwp",
): Promise<Blob> {
  const res = await fetch(
    `${API}/sessions/${sessionId}/download?format=${format}`,
  );
  if (!res.ok) throw new Error("다운로드 실패");
  return res.blob();
}

// ── 모델 ────────────────────────────────────────────────

export async function getModels(): Promise<ModelListResponse> {
  return request("/models");
}

export async function setDefaultModel(modelId: string) {
  return request("/models/default", {
    method: "POST",
    body: JSON.stringify({ model_id: modelId }),
  });
}

// ── 세션 ────────────────────────────────────────────────

export async function getSchema(
  sessionId: string,
): Promise<DocumentSchema> {
  return request(`/sessions/${sessionId}/schema`);
}

export async function getHistory(
  sessionId: string,
): Promise<HistoryResponse> {
  return request(`/sessions/${sessionId}/history`);
}

export async function undoAction(sessionId: string) {
  return request(`/sessions/${sessionId}/undo`, { method: "POST" });
}

export async function redoAction(sessionId: string) {
  return request(`/sessions/${sessionId}/redo`, { method: "POST" });
}

export async function closeSession(sessionId: string) {
  return request(`/sessions/${sessionId}`, { method: "DELETE" });
}

// ── 참고파일 ────────────────────────────────────────────

export async function extractReference(
  file: File,
): Promise<{ filename: string; text: string; length: number }> {
  const fd = new FormData();
  fd.append("file", file);
  const res = await fetch(`${API}/reference/extract`, { method: "POST", body: fd });
  if (!res.ok) throw new Error("텍스트 추출 실패");
  return res.json();
}

// ── 서식 검증 ───────────────────────────────────────────

export async function checkFormat(
  sessionId: string,
  programName: string,
  autoFix = false,
): Promise<FormatReportData> {
  return request(`/sessions/${sessionId}/format-check`, {
    method: "POST",
    body: JSON.stringify({ program_name: programName, auto_fix: autoFix }),
  });
}
