import { FileText, Image, Paperclip, Send, Square } from "lucide-react";
import { useRef, useState } from "react";
import type { FileUploadResponse } from "../types";
import { extractReference, uploadDocument } from "../lib/api";

interface Props {
  onSend: (message: string) => void;
  onSendWithImage?: (message: string, imageDataUrl: string) => void;
  onLocalMessage?: (role: "user" | "assistant", content: string) => void;
  onFileUpload?: (data: FileUploadResponse) => void;
  hasHwpSession: boolean;
  isStreaming: boolean;
  onStop?: () => void;
}

export default function ChatInput({
  onSend,
  onSendWithImage,
  onLocalMessage,
  onFileUpload,
  hasHwpSession,
  isStreaming,
  onStop,
}: Props) {
  const [input, setInput] = useState("");
  const [uploading, setUploading] = useState(false);
  const [pendingImage, setPendingImage] = useState<string | null>(null);
  const [pendingImageName, setPendingImageName] = useState("");
  const [pendingFile, setPendingFile] = useState<File | null>(null);
  const hwpRef = useRef<HTMLInputElement>(null);
  const refRef = useRef<HTMLInputElement>(null);
  const imageRef = useRef<HTMLInputElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  const handleSubmit = async () => {
    if (isStreaming) return;
    const trimmed = input.trim();

    // 참고파일이 첨부되어 있으면 서버에서 텍스트 추출 후 전송
    if (pendingFile) {
      const file = pendingFile;
      setPendingFile(null);
      setInput("");
      if (textareaRef.current) textareaRef.current.style.height = "auto";

      onLocalMessage?.("user", `📄 ${file.name} 분석 중...`);
      try {
        const result = await extractReference(file);
        if (result.text) {
          const msg = trimmed
            ? `[참고파일: ${file.name}]\n${result.text}\n\n${trimmed}`
            : `[참고파일: ${file.name}]\n${result.text}\n\n이 내용을 참고하여 HWP 문서의 빈 셀을 작성해주세요.`;
          onSend(msg);
        } else {
          onLocalMessage?.("assistant", `❌ ${file.name}에서 텍스트를 추출할 수 없습니다.`);
        }
      } catch {
        onLocalMessage?.("assistant", `❌ 참고파일 처리 실패`);
      }
      return;
    }

    if ((!trimmed && !pendingImage) || isStreaming) return;

    if (pendingImage && onSendWithImage) {
      onSendWithImage(trimmed || "이 이미지를 분석해주세요.", pendingImage);
    } else if (trimmed) {
      onSend(trimmed);
    }
    setInput("");
    setPendingImage(null);
    setPendingImageName("");
    if (textareaRef.current) textareaRef.current.style.height = "auto";
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSubmit();
    }
  };

  const handleInput = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setInput(e.target.value);
    const ta = e.target;
    ta.style.height = "auto";
    ta.style.height = Math.min(ta.scrollHeight, 160) + "px";
  };

  // HWP 파일 업로드 (한 번만)
  const handleHwpFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploading(true);
    try {
      onLocalMessage?.("user", `📎 ${file.name} 업로드 중...`);
      const data = await uploadDocument(file);
      onFileUpload?.(data);
      onLocalMessage?.(
        "assistant",
        `✅ **${file.name}** 열기 완료\n\n` +
          `한/글에서 문서가 열렸습니다.\n` +
          `"빈 셀을 채워줘" 또는 원하는 작업을 말씀하세요.\n` +
          `참고자료가 있으면 📄 버튼으로 첨부하세요.`,
      );
    } catch (err) {
      const msg = err instanceof Error ? err.message : "알 수 없는 오류";
      onLocalMessage?.("assistant", `❌ 파일 업로드 실패: ${msg}`);
    } finally {
      setUploading(false);
      if (hwpRef.current) hwpRef.current.value = "";
    }
  };

  // 참고파일 첨부 (txt, pdf, docx 등)
  const handleRefFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setPendingFile(file);
    if (refRef.current) refRef.current.value = "";
  };

  // 이미지 첨부
  const handleImage = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      setPendingImage(reader.result as string);
      setPendingImageName(file.name);
    };
    reader.readAsDataURL(file);
    if (imageRef.current) imageRef.current.value = "";
  };

  return (
    <div className="border-t border-slate-700 bg-slate-900 p-3">
      {/* 이미지 미리보기 */}
      {pendingImage && (
        <div className="max-w-3xl mx-auto mb-2 flex items-center gap-2">
          <img
            src={pendingImage}
            alt="첨부"
            className="w-16 h-16 rounded-lg object-cover border border-slate-600"
          />
          <span className="text-xs text-slate-400 flex-1 truncate">
            {pendingImageName}
          </span>
          <button
            onClick={() => { setPendingImage(null); setPendingImageName(""); }}
            className="text-xs text-red-400 hover:text-red-300"
          >
            제거
          </button>
        </div>
      )}

      {/* 참고파일 미리보기 */}
      {pendingFile && (
        <div className="max-w-3xl mx-auto mb-2 flex items-center gap-2 px-3 py-2 bg-slate-800 rounded-lg">
          <FileText className="w-4 h-4 text-amber-400 flex-shrink-0" />
          <span className="text-xs text-slate-300 flex-1 truncate">
            {pendingFile.name}
          </span>
          <button
            onClick={() => setPendingFile(null)}
            className="text-xs text-red-400 hover:text-red-300"
          >
            제거
          </button>
        </div>
      )}

      <div className="flex items-end gap-1.5 max-w-3xl mx-auto">
        {/* HWP 업로드 (세션 없을 때만) */}
        {!hasHwpSession && (
          <>
            <button
              type="button"
              onClick={() => hwpRef.current?.click()}
              disabled={uploading}
              className="p-2 text-slate-400 hover:text-slate-200 hover:bg-slate-800 rounded-lg transition-colors disabled:opacity-50"
              title="HWP 파일 열기"
            >
              <Paperclip className="w-5 h-5" />
            </button>
            <input
              ref={hwpRef}
              type="file"
              accept=".hwp,.hwpx"
              onChange={handleHwpFile}
              className="hidden"
            />
          </>
        )}

        {/* 참고파일 첨부 (세션 있을 때) */}
        {hasHwpSession && (
          <>
            <button
              type="button"
              onClick={() => refRef.current?.click()}
              className="p-2 text-amber-400 hover:text-amber-300 hover:bg-slate-800 rounded-lg transition-colors"
              title="참고자료 첨부 (txt, pdf 등)"
            >
              <FileText className="w-5 h-5" />
            </button>
            <input
              ref={refRef}
              type="file"
              accept=".txt,.md,.csv,.json,.pdf,.docx"
              onChange={handleRefFile}
              className="hidden"
            />
          </>
        )}

        {/* 이미지 첨부 */}
        <button
          type="button"
          onClick={() => imageRef.current?.click()}
          className="p-2 text-slate-400 hover:text-slate-200 hover:bg-slate-800 rounded-lg transition-colors"
          title="이미지 첨부"
        >
          <Image className="w-5 h-5" />
        </button>
        <input
          ref={imageRef}
          type="file"
          accept="image/*"
          onChange={handleImage}
          className="hidden"
        />

        {/* 입력창 */}
        <textarea
          ref={textareaRef}
          value={input}
          onChange={handleInput}
          onKeyDown={handleKeyDown}
          placeholder={
            pendingFile
              ? `${pendingFile.name} 참고하여 질문하세요...`
              : pendingImage
                ? "이미지에 대해 질문하세요..."
                : hasHwpSession
                  ? "채팅으로 문서를 수정하세요..."
                  : "HWP 파일을 첨부하여 시작하세요..."
          }
          rows={1}
          className="flex-1 resize-none bg-slate-800 border border-slate-600 rounded-xl px-4 py-2.5 text-slate-200 placeholder-slate-500 focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-transparent text-sm leading-relaxed"
        />

        {/* 전송 / 중지 */}
        {isStreaming ? (
          <button
            type="button"
            onClick={onStop}
            className="p-2.5 bg-red-600 hover:bg-red-500 text-white rounded-xl transition-colors"
            title="생성 중지"
          >
            <Square className="w-4 h-4" />
          </button>
        ) : (
          <button
            type="button"
            onClick={handleSubmit}
            disabled={!input.trim() && !pendingImage && !pendingFile}
            className="p-2.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl transition-colors disabled:opacity-40 disabled:cursor-not-allowed"
            title="전송"
          >
            <Send className="w-4 h-4" />
          </button>
        )}
      </div>
    </div>
  );
}
