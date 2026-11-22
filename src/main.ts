import "./style.css";
import { marked } from "marked";
import DOMPurify from "dompurify";
import { Content, GeminiError, Part, UsageMetadata, listModels, streamGenerate } from "./gemini";
import {
  ChangedRange,
  TOOL_DECLARATIONS,
  TOOL_LABELS,
  ToolContext,
  UndoOp,
  WRITE_TOOLS,
  executeTool,
  getSelectionInfo,
  selectAddress,
} from "./tools";
import { PRESET_MODELS, THINKING_LEVELS, Settings, loadSettings, saveSettings } from "./settings";

const MAX_STEPS = 40;

const SYSTEM_PROMPT = `你是嵌入在 Microsoft Excel 侧边栏中的 AI 助手，通过工具直接读取和修改用户当前打开的工作簿。

工作原则：
1. 先了解再动手：不确定结构时先调用 get_workbook_info / read_range，不要猜测单元格内容。
2. 用真实公式建模：计算类结果优先写成引用其它单元格的 Excel 公式，而不是硬编码数值；保持现有公式不被破坏。公式一律用英文函数名和英文逗号（例如 =SUMIFS(C:C,A:A,"北京")）。
3. 谨慎修改：覆盖已有非空数据前先确认这是用户意图；只修改与任务相关的区域。删除行列、清除大范围数据前先说明。
4. 写入后检查：关注 write_range 返回的 errors，发现 #REF!/#NAME?/#VALUE! 等错误要修正。
5. 引用单元格时使用 工作表!地址 形式（例如 Sheet1!B2:D10），方便用户定位。
6. 完成修改后，用简短列表总结改了哪些区域、做了什么。用户可以点击“撤销本轮修改”回滚。
7. 用用户使用的语言回答（默认中文），简洁、专业，适当使用 Markdown 表格。
8. 如果无法完成（例如缺少数据或 API 不支持），直接说明原因，不要编造结果。`;

/* ------------------------------ state ------------------------------ */

let settings: Settings = loadSettings();
let history: Content[] = [];
let busy = false;
let abort: AbortController | null = null;
let pendingConfirm: ((ok: boolean) => void) | null = null;
const inExcel = () => typeof Excel !== "undefined" && Office?.context?.host === Office.HostType.Excel;

const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;
const messagesEl = $("messages");
const inputEl = $<HTMLTextAreaElement>("input");
const sendBtn = $<HTMLButtonElement>("sendBtn");
const stopBtn = $<HTMLButtonElement>("stopBtn");
const statusText = $("statusText");

function el<K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string, text?: string) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
}

function renderMarkdown(target: HTMLElement, text: string) {
  target.innerHTML = DOMPurify.sanitize(marked.parse(text, { async: false, breaks: true }) as string);
  target.querySelectorAll("a").forEach((a) => {
    a.target = "_blank";
    a.rel = "noopener";
  });
}

function scrollToBottom() {
  messagesEl.scrollTop = messagesEl.scrollHeight;
}

/* ---------------------------- model lists --------------------------- */

function allModels(): { id: string; label: string }[] {
  const list = [...PRESET_MODELS];
  for (const id of settings.customModels) if (!list.some((m) => m.id === id)) list.push({ id, label: id });
  if (settings.model && !list.some((m) => m.id === settings.model))
    list.push({ id: settings.model, label: settings.model });
  return list;
}

function fillModelSelect(sel: HTMLSelectElement, value: string) {
  sel.innerHTML = "";
  for (const m of allModels()) {
    const o = el("option", undefined, m.label);
    o.value = m.id;
    sel.appendChild(o);
  }
  sel.value = value;
}

function refreshModelSelectors() {
  fillModelSelect($<HTMLSelectElement>("modelQuick"), settings.model);
  fillModelSelect($<HTMLSelectElement>("modelSelect"), settings.model);
}

/* ------------------------------ settings ---------------------------- */

function openSettings() {
  $<HTMLInputElement>("apiKeyInput").value = settings.apiKey;
  refreshModelSelectors();
  const th = $<HTMLSelectElement>("thinkingSelect");
  th.innerHTML = "";
  for (const t of THINKING_LEVELS) {
    const o = el("option", undefined, t.label);
    o.value = t.id;
    th.appendChild(o);
  }
  th.value = settings.thinkingLevel;
  $<HTMLInputElement>("confirmWrites").checked = settings.confirmWrites;
  $("testStatus").textContent = "";
  $("settingsPanel").classList.remove("hidden");
}

function readSettingsForm(): Settings {
  return {
    ...settings,
    apiKey: $<HTMLInputElement>("apiKeyInput").value.trim(),
    model: $<HTMLSelectElement>("modelSelect").value,
    thinkingLevel: $<HTMLSelectElement>("thinkingSelect").value,
    confirmWrites: $<HTMLInputElement>("confirmWrites").checked,
  };
}

function bindSettings() {
  $("settingsBtn").onclick = openSettings;
  $("closeSettings").onclick = () => $("settingsPanel").classList.add("hidden");
  $("toggleKey").onclick = () => {
    const i = $<HTMLInputElement>("apiKeyInput");
    i.type = i.type === "password" ? "text" : "password";
    $("toggleKey").textContent = i.type === "password" ? "显示" : "隐藏";
  };
  $("addModelBtn").onclick = () => {
    const id = $<HTMLInputElement>("customModelInput").value.trim().replace(/^models\//, "");
    if (!id) return;
    if (!settings.customModels.includes(id)) settings.customModels.push(id);
    settings.model = id;
    $<HTMLInputElement>("customModelInput").value = "";
    refreshModelSelectors();
  };
  $("fetchModelsBtn").onclick = async () => {
    const key = $<HTMLInputElement>("apiKeyInput").value.trim();
    const st = $("fetchStatus");
    if (!key) return void (st.textContent = "请先填写 API Key");
    st.textContent = "获取中…";
    try {
      const models = await listModels(key);
      const current = $<HTMLSelectElement>("modelSelect").value;
      for (const m of models) if (!settings.customModels.includes(m.id)) settings.customModels.push(m.id);
      settings.model = current;
      refreshModelSelectors();
      st.textContent = `已获取 ${models.length} 个模型`;
    } catch (e) {
      st.textContent = (e as Error).message;
    }
  };
  $("testBtn").onclick = async () => {
    const s = readSettingsForm();
    const st = $("testStatus");
    if (!s.apiKey) return void (st.textContent = "请先填写 API Key");
    st.textContent = "测试中…";
    try {
      let text = "";
      for await (const c of streamGenerate({
        apiKey: s.apiKey,
        model: s.model,
        thinkingLevel: s.thinkingLevel,
        contents: [{ role: "user", parts: [{ text: "只回复：OK" }] }],
      }))
        text += c.parts.map((p) => (p.thought ? "" : p.text ?? "")).join("");
      st.textContent = `✅ 连接成功（${s.model}）：${text.trim().slice(0, 40)}`;
    } catch (e) {
      st.textContent = `❌ ${(e as Error).message}`;
    }
  };
  $("saveSettings").onclick = () => {
    settings = readSettingsForm();
    saveSettings(settings);
    refreshModelSelectors();
    $("settingsPanel").classList.add("hidden");
  };
  $<HTMLSelectElement>("modelQuick").onchange = (e) => {
    settings.model = (e.target as HTMLSelectElement).value;
    saveSettings(settings);
  };
}

/* ------------------------------ selection --------------------------- */

async function refreshSelectionPill() {
  if (!inExcel()) return;
  try {
    const info = await Excel.run(async (ctx) => {
      const r = ctx.workbook.getSelectedRange();
      r.load("address");
      await ctx.sync();
      return r.address;
    });
    $("selectionPill").textContent = `选区：${info}`;
  } catch {
    /* ignore */
  }
}

/* -------------------------------- UI -------------------------------- */

function addUserMessage(text: string) {
  $("welcome")?.remove();
  const m = el("div", "msg user");
  m.appendChild(el("div", "bubble", text));
  messagesEl.appendChild(m);
  scrollToBottom();
}

function newAssistantBlock() {
  const m = el("div", "msg assistant");
  messagesEl.appendChild(m);
  return m;
}

function summarizeArgs(name: string, args: Record<string, unknown>): string {
  const sheet = args.sheet ? `${args.sheet}!` : "";
  const addr = args.address ?? args.start_cell ?? args.data_address;
  if (addr) return `${sheet}${addr}`;
  if (args.name) return String(args.name);
  if (args.text) return `“${args.text}”`;
  if (name === "rename_sheet") return `${args.sheet} → ${args.new_name}`;
  return "";
}

interface ToolCard {
  root: HTMLElement;
  setStatus: (s: "running" | "ok" | "error" | "waiting" | "rejected", msg?: string) => void;
  setResult: (r: unknown, changed?: ChangedRange[]) => void;
  confirm: () => Promise<boolean>;
}

function createToolCard(parent: HTMLElement, name: string, args: Record<string, unknown>): ToolCard {
  const root = el("div", "tool-card");
  const head = el("div", "tool-head");
  const icon = el("span", "tool-status running", "●");
  const label = el("span", "tool-label", TOOL_LABELS[name] ?? name);
  const summary = el("span", "tool-summary", summarizeArgs(name, args));
  const toggle = el("span", "tool-toggle", "▸");
  head.append(icon, label, summary, toggle);
  const body = el("div", "tool-body hidden");
  const argsPre = el("pre", undefined, JSON.stringify(args, null, 2));
  body.append(el("div", "tool-section", "参数"), argsPre);
  const links = el("div", "tool-links");
  const actions = el("div", "tool-actions hidden");
  root.append(head, links, actions, body);
  head.onclick = () => {
    body.classList.toggle("hidden");
    toggle.textContent = body.classList.contains("hidden") ? "▸" : "▾";
  };
  parent.appendChild(root);
  scrollToBottom();

  return {
    root,
    setStatus(s, msg) {
      icon.className = `tool-status ${s}`;
      icon.textContent = s === "ok" ? "✓" : s === "error" ? "✕" : s === "rejected" ? "⊘" : "●";
      if (msg) root.title = msg;
    },
    setResult(r, changed) {
      body.append(el("div", "tool-section", "结果"), el("pre", undefined, JSON.stringify(r, null, 2)));
      for (const c of changed ?? []) {
        const a = el("a", "range-link", `${c.sheet}!${c.address}`);
        a.href = "#";
        a.onclick = (e) => {
          e.preventDefault();
          selectAddress(c.sheet, c.address).catch(() => {});
        };
        links.appendChild(a);
      }
    },
    confirm() {
      actions.classList.remove("hidden");
      actions.innerHTML = "";
      const ok = el("button", "primary-btn small", "允许");
      const no = el("button", "small-btn", "拒绝");
      actions.append(el("span", "hint", "需要修改工作簿，是否允许？"), ok, no);
      scrollToBottom();
      return new Promise<boolean>((resolve) => {
        const done = (v: boolean) => {
          pendingConfirm = null;
          actions.classList.add("hidden");
          resolve(v);
        };
        pendingConfirm = done;
        ok.onclick = () => done(true);
        no.onclick = () => done(false);
      });
    },
  };
}

function addError(parent: HTMLElement, msg: string) {
  parent.appendChild(el("div", "error-box", msg));
  scrollToBottom();
}

function setBusy(b: boolean) {
  busy = b;
  sendBtn.classList.toggle("hidden", b);
  stopBtn.classList.toggle("hidden", !b);
  statusText.textContent = b ? "思考中…" : "";
}

/* ------------------------------- agent ------------------------------ */

async function buildContextPrefix(): Promise<string> {
  if (!inExcel()) return "";
  try {
    const info = await getSelectionInfo(50);
    return `[自动附加的上下文] 活动工作表：${info.active_sheet}；当前选区：${info.selection}${
      info.values ? `；选区内容：${JSON.stringify(info.values)}` : ""
    }${info.formulas ? `；选区公式：${JSON.stringify(info.formulas)}` : ""}\n\n`;
  } catch {
    return "";
  }
}

function pushUser(parts: Part[]) {
  const last = history[history.length - 1];
  if (last && last.role === "user") last.parts.push(...parts);
  else history.push({ role: "user", parts });
}

/** Makes sure every model functionCall has a matching functionResponse (e.g. after abort). */
function repairHistory() {
  const last = history[history.length - 1];
  if (last?.role === "model") {
    const calls = last.parts.filter((p) => p.functionCall);
    if (calls.length)
      history.push({
        role: "user",
        parts: calls.map((p) => ({
          functionResponse: { name: p.functionCall!.name, id: p.functionCall!.id, response: { error: "用户已中止" } },
        })),
      });
  }
}

async function runTurn(userText: string) {
  if (!settings.apiKey) {
    openSettings();
    $("testStatus").textContent = "请先填写 Gemini API Key。";
    return;
  }
  addUserMessage(userText);
  const block = newAssistantBlock();
  const undoOps: UndoOp[] = [];
  const tctx: ToolContext = { undo: undoOps };
  const usageTotal: UsageMetadata = { promptTokenCount: 0, candidatesTokenCount: 0, thoughtsTokenCount: 0 };

  setBusy(true);
  abort = new AbortController();

  pushUser([{ text: (await buildContextPrefix()) + userText }]);

  try {
    for (let step = 0; step < MAX_STEPS; step++) {
      const parts: Part[] = [];
      let textEl: HTMLElement | null = null;
      let text = "";
      let finishReason: string | undefined;
      let blockReason: string | undefined;
      let usage: UsageMetadata | undefined;

      statusText.textContent = step === 0 ? "思考中…" : "继续处理…";
      for await (const chunk of streamGenerate({
        apiKey: settings.apiKey,
        model: settings.model,
        thinkingLevel: settings.thinkingLevel,
        systemInstruction: SYSTEM_PROMPT,
        tools: TOOL_DECLARATIONS,
        contents: history,
        signal: abort.signal,
      })) {
        for (const p of chunk.parts) {
          const empty = !p.text && !p.functionCall && !p.thoughtSignature && !p.functionResponse;
          if (!empty) parts.push(p);
          if (p.text && !p.thought) {
            text += p.text;
            if (!textEl) {
              textEl = el("div", "md");
              block.appendChild(textEl);
            }
            renderMarkdown(textEl, text);
            scrollToBottom();
          }
        }
        if (chunk.finishReason) finishReason = chunk.finishReason;
        if (chunk.blockReason) blockReason = chunk.blockReason;
        if (chunk.usage) usage = chunk.usage;
      }
      if (usage) {
        usageTotal.promptTokenCount! += usage.promptTokenCount ?? 0;
        usageTotal.candidatesTokenCount! += usage.candidatesTokenCount ?? 0;
        usageTotal.thoughtsTokenCount! += usage.thoughtsTokenCount ?? 0;
      }

      if (!parts.length) {
        addError(
          block,
          blockReason
            ? `请求被拦截：${blockReason}`
            : `模型没有返回内容${finishReason ? `（finishReason: ${finishReason}）` : ""}。`
        );
        break;
      }
      history.push({ role: "model", parts });

      const calls = parts.filter((p) => p.functionCall).map((p) => p.functionCall!);
      if (!calls.length) {
        if (finishReason && finishReason !== "STOP")
          addError(block, `生成提前结束：${finishReason}`);
        break;
      }

      const responses: Part[] = [];
      for (const call of calls) {
        const args = (call.args ?? {}) as Record<string, unknown>;
        const card = createToolCard(block, call.name, args);
        let response: Record<string, unknown>;
        if (abort.signal.aborted) {
          card.setStatus("rejected");
          response = { error: "用户已中止" };
        } else if (!inExcel()) {
          card.setStatus("error");
          response = { error: "当前不在 Excel 环境中，无法执行工具。" };
        } else {
          let allowed = true;
          if (settings.confirmWrites && WRITE_TOOLS.has(call.name)) {
            card.setStatus("waiting");
            statusText.textContent = "等待确认…";
            allowed = await card.confirm();
          }
          if (!allowed) {
            card.setStatus("rejected");
            response = { error: "用户拒绝了此操作。请询问用户希望如何调整。" };
          } else {
            statusText.textContent = `${TOOL_LABELS[call.name] ?? call.name}…`;
            try {
              const r = await executeTool(call.name, args, tctx);
              card.setStatus("ok");
              card.setResult(r.response, r.changed);
              response = { result: r.response };
            } catch (e) {
              const msg = (e as Error).message || String(e);
              card.setStatus("error", msg);
              card.setResult({ error: msg });
              response = { error: msg };
            }
          }
        }
        responses.push({ functionResponse: { name: call.name, id: call.id, response } });
      }
      history.push({ role: "user", parts: responses });
      if (abort.signal.aborted) break;
      if (step === MAX_STEPS - 1) addError(block, `已达到单轮最多 ${MAX_STEPS} 步，自动停止。可以输入“继续”。`);
    }
  } catch (e) {
    if ((e as Error).name === "AbortError") addError(block, "已停止。");
    else if (e instanceof GeminiError) addError(block, `Gemini API 错误：${e.message}`);
    else addError(block, `出错了：${(e as Error).message ?? e}`);
  } finally {
    repairHistory();
    setBusy(false);
    abort = null;
    renderTurnFooter(block, undoOps, usageTotal);
    refreshSelectionPill();
  }
}

function renderTurnFooter(block: HTMLElement, undoOps: UndoOp[], usage: UsageMetadata) {
  const foot = el("div", "turn-footer");
  if (undoOps.length) {
    const btn = el("button", "small-btn", `↶ 撤销本轮修改（${undoOps.length} 项）`);
    btn.onclick = async () => {
      btn.disabled = true;
      btn.textContent = "撤销中…";
      const errors: string[] = [];
      for (const op of [...undoOps].reverse()) {
        try {
          await op();
        } catch (e) {
          errors.push((e as Error).message);
        }
      }
      btn.textContent = errors.length ? `已撤销（${errors.length} 项失败）` : "✓ 已撤销";
      if (errors.length) btn.title = errors.join("\n");
      pushUser([{ text: "[系统提示] 用户已撤销了你上一轮对工作簿所做的全部修改。" }]);
    };
    foot.appendChild(btn);
  }
  const total = (usage.promptTokenCount ?? 0) + (usage.candidatesTokenCount ?? 0) + (usage.thoughtsTokenCount ?? 0);
  if (total)
    foot.appendChild(
      el(
        "span",
        "usage",
        `${settings.model} · 输入 ${usage.promptTokenCount} · 输出 ${usage.candidatesTokenCount}${
          usage.thoughtsTokenCount ? ` · 思考 ${usage.thoughtsTokenCount}` : ""
        } tokens`
      )
    );
  if (foot.childNodes.length) block.appendChild(foot);
  scrollToBottom();
}

/* ------------------------------ wiring ------------------------------ */

function send() {
  const text = inputEl.value.trim();
  if (!text || busy) return;
  inputEl.value = "";
  autosize();
  runTurn(text);
}

function autosize() {
  inputEl.style.height = "auto";
  inputEl.style.height = Math.min(inputEl.scrollHeight, 160) + "px";
}

function newChat() {
  if (busy) return;
  history = [];
  messagesEl.innerHTML = "";
  location.reload();
}

function init() {
  bindSettings();
  refreshModelSelectors();
  sendBtn.onclick = send;
  stopBtn.onclick = () => {
    abort?.abort();
    pendingConfirm?.(false);
  };
  $("newChatBtn").onclick = newChat;
  inputEl.addEventListener("input", autosize);
  inputEl.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey && !e.isComposing) {
      e.preventDefault();
      send();
    }
  });
  document.querySelectorAll<HTMLButtonElement>(".chip").forEach((c) => {
    c.onclick = () => {
      inputEl.value = c.textContent ?? "";
      send();
    };
  });

  if (inExcel()) {
    refreshSelectionPill();
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => refreshSelectionPill());
  } else {
    $("selectionPill").textContent = "未在 Excel 中运行（仅预览界面）";
  }
  if (!settings.apiKey) openSettings();
}

if (typeof Office !== "undefined") {
  Office.onReady(() => init());
} else {
  init();
}
