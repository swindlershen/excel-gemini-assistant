// Minimal Gemini REST client (browser-side, user's own API key).
// Uses streamGenerateContent over SSE and preserves raw parts (incl. thoughtSignature),
// which Gemini 3+/4 models require to be sent back unchanged during function calling.

const API_BASE = "https://generativelanguage.googleapis.com/v1beta";

export interface FunctionCall {
  name: string;
  args?: Record<string, unknown>;
  id?: string;
}

export interface Part {
  text?: string;
  inlineData?: { mimeType: string; data: string };
  thought?: boolean;
  thoughtSignature?: string;
  functionCall?: FunctionCall;
  functionResponse?: { name: string; response: Record<string, unknown>; id?: string };
}

export interface Content {
  role: "user" | "model";
  parts: Part[];
}

export interface FunctionDeclaration {
  name: string;
  description: string;
  parameters?: Record<string, unknown>;
}

export interface UsageMetadata {
  promptTokenCount?: number;
  candidatesTokenCount?: number;
  thoughtsTokenCount?: number;
  totalTokenCount?: number;
}

export interface StreamChunk {
  parts: Part[];
  finishReason?: string;
  usage?: UsageMetadata;
  blockReason?: string;
}

export interface GenerateOptions {
  apiKey: string;
  model: string;
  contents: Content[];
  systemInstruction?: string;
  tools?: FunctionDeclaration[];
  thinkingLevel?: string; // "", "low", "medium", "high", "xhigh"
  signal?: AbortSignal;
}

export class GeminiError extends Error {
  constructor(message: string, public status?: number) {
    super(message);
  }
}

export function normalizeModelId(model: string): string {
  return model.trim().replace(/^models\//, "");
}

async function readError(res: Response): Promise<GeminiError> {
  let msg = `HTTP ${res.status}`;
  try {
    const data = await res.json();
    const m = Array.isArray(data) ? data[0]?.error?.message : data?.error?.message;
    if (m) msg += `: ${m}`;
  } catch {
    /* ignore */
  }
  return new GeminiError(msg, res.status);
}

async function doFetch(url: string, init: RequestInit): Promise<Response> {
  try {
    return await fetch(url, init);
  } catch (e) {
    if ((e as Error).name === "AbortError") throw e;
    throw new GeminiError(
      "无法连接 Gemini API（generativelanguage.googleapis.com）。请检查网络 / 代理是否可以访问 Google。"
    );
  }
}

export async function* streamGenerate(opts: GenerateOptions): AsyncGenerator<StreamChunk> {
  const model = normalizeModelId(opts.model);
  const url = `${API_BASE}/models/${encodeURIComponent(model)}:streamGenerateContent?alt=sse`;

  const body: Record<string, unknown> = { contents: opts.contents };
  if (opts.systemInstruction) {
    body.systemInstruction = { parts: [{ text: opts.systemInstruction }] };
  }
  if (opts.tools?.length) {
    body.tools = [{ functionDeclarations: opts.tools }];
  }
  if (opts.thinkingLevel) {
    body.generationConfig = { thinkingConfig: { thinkingLevel: opts.thinkingLevel } };
  }

  const res = await doFetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-goog-api-key": opts.apiKey },
    body: JSON.stringify(body),
    signal: opts.signal,
  });
  if (!res.ok || !res.body) throw await readError(res);

  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";

  const parseEvent = (raw: string): StreamChunk | null => {
    const dataLines = raw
      .split(/\r?\n/)
      .filter((l) => l.startsWith("data:"))
      .map((l) => l.slice(5).trimStart());
    if (!dataLines.length) return null;
    const json = JSON.parse(dataLines.join("\n"));
    if (json.error) throw new GeminiError(json.error.message ?? "Unknown error", json.error.code);
    const cand = json.candidates?.[0];
    return {
      parts: (cand?.content?.parts ?? []) as Part[],
      finishReason: cand?.finishReason,
      usage: json.usageMetadata,
      blockReason: json.promptFeedback?.blockReason,
    };
  };

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });
    let idx: number;
    while ((idx = buffer.search(/\r?\n\r?\n/)) >= 0) {
      const raw = buffer.slice(0, idx);
      buffer = buffer.slice(idx).replace(/^\r?\n\r?\n/, "");
      const chunk = parseEvent(raw);
      if (chunk) yield chunk;
    }
  }
  if (buffer.trim()) {
    const chunk = parseEvent(buffer);
    if (chunk) yield chunk;
  }
}

export interface ModelInfo {
  id: string;
  displayName: string;
}

export async function listModels(apiKey: string): Promise<ModelInfo[]> {
  const out: ModelInfo[] = [];
  let pageToken = "";
  do {
    const url = `${API_BASE}/models?pageSize=1000${pageToken ? `&pageToken=${pageToken}` : ""}`;
    const res = await doFetch(url, { headers: { "x-goog-api-key": apiKey } });
    if (!res.ok) throw await readError(res);
    const data = await res.json();
    for (const m of data.models ?? []) {
      const methods: string[] = m.supportedGenerationMethods ?? [];
      const id = normalizeModelId(m.name);
      if (!methods.includes("generateContent")) continue;
      if (/image|tts|embedding|aqa|imagen|veo|transcribe|live|robotics/i.test(id)) continue;
      out.push({ id, displayName: m.displayName ?? id });
    }
    pageToken = data.nextPageToken ?? "";
  } while (pageToken);
  return out;
}
