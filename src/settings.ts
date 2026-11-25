export interface Settings {
  apiKey: string;
  model: string;
  thinkingLevel: string; // "" = model default
  confirmWrites: boolean;
  customModels: string[];
}

export const PRESET_MODELS: { id: string; label: string }[] = [
  { id: "gemini-4-hydrogen", label: "Gemini 4 Hydrogen（快速）" },
  { id: "gemini-4-neon", label: "Gemini 4 Neon（更强，需付费层）" },
  { id: "gemini-4-carbon", label: "Gemini 4 Carbon（最强）" },
];

export const THINKING_LEVELS = [
  { id: "", label: "自动（低思考，省费推荐）" },
  { id: "low", label: "低" },
  { id: "medium", label: "中" },
  { id: "high", label: "高" },
  { id: "xhigh", label: "极高" },
];

const KEY = "excel-gemini-assistant.settings.v1";

const DEFAULTS: Settings = {
  apiKey: "",
  model: "gemini-4-hydrogen",
  thinkingLevel: "",
  confirmWrites: false,
  customModels: [],
};

export function loadSettings(): Settings {
  try {
    const raw = localStorage.getItem(KEY);
    if (raw) return { ...DEFAULTS, ...JSON.parse(raw) };
  } catch {
    /* ignore */
  }
  return { ...DEFAULTS };
}

export function saveSettings(s: Settings) {
  localStorage.setItem(KEY, JSON.stringify(s));
}
