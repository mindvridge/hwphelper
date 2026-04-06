import { Bot, Check, ChevronDown } from "lucide-react";
import { useEffect, useState } from "react";
import type { ModelInfo } from "../types";
import { getModels, setDefaultModel } from "../lib/api";

interface Props {
  value: string;
  onChange: (modelId: string) => void;
}

const PROVIDER_LABEL: Record<string, string> = {
  anthropic: "Anthropic",
  openai: "OpenAI",
  openai_compatible: "Custom",
};

export default function ModelSelector({ value, onChange }: Props) {
  const [models, setModels] = useState<ModelInfo[]>([]);
  const [open, setOpen] = useState(false);

  useEffect(() => {
    getModels()
      .then((data) => {
        setModels(data.models);
        if (!value && data.default_model) onChange(data.default_model);
      })
      .catch(() => {
        setModels([
          { id: "claude-sonnet", provider: "anthropic", model: "claude-sonnet-4", description: "Claude Sonnet 4", available: true },
          { id: "gpt-4o", provider: "openai", model: "gpt-4o", description: "GPT-4o", available: false },
        ]);
      });
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const selected = models.find((m) => m.id === value);

  const handleSelect = async (id: string) => {
    setOpen(false);
    onChange(id);
    try {
      await setDefaultModel(id);
    } catch {
      /* 서버 미연결시 무시 */
    }
  };

  return (
    <div className="relative">
      <button
        onClick={() => setOpen(!open)}
        className="flex items-center gap-2 px-3 py-1.5 rounded-lg bg-slate-800 border border-slate-700 text-sm text-slate-300 hover:bg-slate-700 transition-colors"
      >
        <Bot className="w-4 h-4 text-indigo-400" />
        <span className="truncate max-w-[140px]">
          {selected?.id ?? "모델 선택"}
        </span>
        <ChevronDown className="w-3.5 h-3.5 text-slate-500" />
      </button>

      {open && (
        <>
          <div className="fixed inset-0 z-40" onClick={() => setOpen(false)} />
          <div className="absolute top-full mt-1 left-0 z-50 w-72 bg-slate-800 border border-slate-700 rounded-xl shadow-xl overflow-hidden">
            {models.map((m) => (
              <button
                key={m.id}
                onClick={() => m.available && handleSelect(m.id)}
                disabled={!m.available}
                className={`w-full text-left px-3 py-2.5 flex items-start gap-2 transition-colors ${
                  m.available
                    ? "hover:bg-slate-700/60 cursor-pointer"
                    : "opacity-40 cursor-not-allowed"
                } ${m.id === value ? "bg-slate-700/40" : ""}`}
              >
                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-1.5">
                    <span className="text-sm font-medium text-slate-200">
                      {m.id}
                    </span>
                    {m.id === value && (
                      <Check className="w-3.5 h-3.5 text-indigo-400" />
                    )}
                  </div>
                  <p className="text-xs text-slate-500 truncate">
                    {PROVIDER_LABEL[m.provider] ?? m.provider} ·{" "}
                    {m.available ? m.description : "API 키 미설정"}
                  </p>
                </div>
              </button>
            ))}
          </div>
        </>
      )}
    </div>
  );
}
