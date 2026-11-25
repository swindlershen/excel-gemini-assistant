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
import { Attachment, MAX_TOTAL_BYTES, fileToAttachment } from "./attachments";
import {
  OrderDiscrepancy,
  extractInvoices,
  inspectInvoiceSheet,
  readPendingOrders,
  summarize,
  updateOrderAmount,
  updateOrderRows,
  writeInvoices,
} from "./invoice";

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
8. 如果无法完成（例如缺少数据或 API 不支持），直接说明原因，不要编造结果。

发票 / 票据识别（用户上传图片或 PDF 并要求录入时；尽量一次性写入所有行，减少工具调用次数）：
- 逐张仔细识别，看不清或缺失的字段留空并在“备注”列注明“需核对”，绝不臆造号码或金额。
- 默认录入到名为“收票登记”的工作表：不存在就用 add_sheet 新建并写表头；已存在则先 read_range 读取表头和已有数据，按现有表头列顺序在最后一行之后追加，不要覆盖已有行。
- 默认表头：序号, 发票类型, 发票代码, 发票号码, 开票日期, 购买方名称, 购买方税号, 销售方名称, 销售方税号, 项目名称, 金额(不含税), 税率, 税额, 价税合计, 校验, 来源文件, 备注
  （全电发票没有发票代码则留空；项目名称多行时用“；”连接主要项目。）
- 日期写成 yyyy-mm-dd；金额、税额写纯数字（不要带 ¥ 和逗号）；税率写成小数（如 0.13）并设置百分比格式；发票号码、税号、代码作为文本写入（在前面加英文单引号 ' 以保留前导零）。
- “校验”列写公式核对 金额+税额=价税合计，例如 =IF(ABS(K2+M2-N2)<0.01,"✓","需核对")（按实际列调整）。
- 录入前检查发票号码是否已存在于台账中，重复的不要再录入，并在回复中说明。
- 首次建表时给表头加粗加底色、金额列设置 #,##0.00 格式、自动调整列宽。
- 完成后用表格简要列出本次录入的发票（号码、销售方、价税合计）和需要人工核对的项。
- 用户要求明细时，另建“发票明细”工作表，每个商品/服务一行，并带上发票号码关联。`;

const INVOICE_PROMPT = "请识别附件中的发票，按规则录入“收票登记”工作表。";

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

function attachmentChip(a: Attachment): HTMLElement {
  const chip = el("div", "att");
  if (a.previewUrl) {
    const img = el("img");
    img.src = a.previewUrl;
    chip.appendChild(img);
  } else {
    chip.appendChild(el("span", "att-icon", a.mimeType === "application/pdf" ? "PDF" : "IMG"));
  }
  chip.appendChild(el("span", "att-name", a.name));
  return chip;
}

function addUserMessage(text: string, atts: Attachment[] = []) {
  $("welcome")?.remove();
  const m = el("div", "msg user");
  const bubble = el("div", "bubble", text);
  if (atts.length) {
    const strip = el("div", "att-list");
    atts.forEach((a) => strip.appendChild(attachmentChip(a)));
    bubble.appendChild(strip);
  }
  m.appendChild(bubble);
  messagesEl.appendChild(m);
  scrollToBottom();
}

/* ---------------------------- attachments --------------------------- */

let attachments: Attachment[] = [];

function renderAttachStrip() {
  const strip = $("attachStrip");
  strip.innerHTML = "";
  strip.classList.toggle("hidden", !attachments.length);
  attachments.forEach((a, i) => {
    const chip = attachmentChip(a);
    const x = el("button", "att-remove", "✕");
    x.title = "移除";
    x.onclick = () => {
      attachments.splice(i, 1);
      renderAttachStrip();
    };
    chip.appendChild(x);
    strip.appendChild(chip);
  });
}

async function addFiles(files: File[] | FileList) {
  for (const f of Array.from(files)) {
    try {
      statusText.textContent = `读取 ${f.name || "图片"}…`;
      const a = await fileToAttachment(f);
      const total = attachments.reduce((s, x) => s + x.bytes, 0) + a.bytes;
      if (total > MAX_TOTAL_BYTES) {
        statusText.textContent = "附件总大小超过 18MB，请分批发送";
        break;
      }
      attachments.push(a);
    } catch (e) {
      statusText.textContent = (e as Error).message;
    }
  }
  if (!busy && statusText.textContent?.startsWith("读取")) statusText.textContent = "";
  renderAttachStrip();
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

/** Older attachments are replaced by placeholders once total inline payload would exceed the request limit. */
function pruneInlineData(incomingBytes: number) {
  let total = incomingBytes;
  const inlineParts: Part[] = [];
  for (const c of history)
    for (const p of c.parts)
      if (p.inlineData) {
        total += Math.round((p.inlineData.data.length * 3) / 4);
        inlineParts.push(p);
      }
  for (const p of inlineParts) {
    if (total <= MAX_TOTAL_BYTES) break;
    total -= Math.round((p.inlineData!.data.length * 3) / 4);
    delete p.inlineData;
    p.text = "[较早的附件已从上下文中移除以节省空间]";
  }
}

async function runTurn(userText: string, atts: Attachment[] = []) {
  if (!settings.apiKey) {
    openSettings();
    $("testStatus").textContent = "请先填写 Gemini API Key。";
    return;
  }
  addUserMessage(userText, atts);
  const block = newAssistantBlock();
  const undoOps: UndoOp[] = [];
  const tctx: ToolContext = { undo: undoOps };
  const usageTotal: UsageMetadata = { promptTokenCount: 0, candidatesTokenCount: 0, thoughtsTokenCount: 0 };
  let apiCalls = 0;

  setBusy(true);
  abort = new AbortController();

  const fileNote = atts.length ? `\n\n[附件 ${atts.length} 个，按顺序为：${atts.map((a) => a.name).join("，")}]` : "";
  pruneInlineData(atts.reduce((s, a) => s + a.bytes, 0));
  pushUser([
    { text: (await buildContextPrefix()) + userText + fileNote },
    ...atts.flatMap((a, i) => [
      { text: `附件${i + 1}：${a.name}` },
      { inlineData: { mimeType: a.mimeType, data: a.data } },
    ]),
  ]);

  try {
    for (let step = 0; step < MAX_STEPS; step++) {
      const parts: Part[] = [];
      let textEl: HTMLElement | null = null;
      let text = "";
      let finishReason: string | undefined;
      let blockReason: string | undefined;
      let usage: UsageMetadata | undefined;

      statusText.textContent = step === 0 ? "思考中…" : "继续处理…";
      apiCalls++;
      for await (const chunk of streamGenerate({
        apiKey: settings.apiKey,
        model: settings.model,
        thinkingLevel: settings.thinkingLevel || (/gemini-(3|4)/.test(settings.model) ? "low" : ""),
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
    dropInlineData();
    setBusy(false);
    abort = null;
    renderTurnFooter(block, undoOps, usageTotal, apiCalls);
    refreshSelectionPill();
  }
}

/** Attachments are only needed in the turn they were sent; keeping them would re-bill them on every later call. */
function dropInlineData() {
  for (const c of history)
    for (const p of c.parts)
      if (p.inlineData) {
        delete p.inlineData;
        p.text = "[附件内容已在上一轮处理，如需再次分析请重新上传]";
      }
}

/* --------------------------- invoice fast path --------------------------- */

async function runInvoiceTurn(userText: string, atts: Attachment[]) {
  if (!settings.apiKey) {
    openSettings();
    $("testStatus").textContent = "请先填写 Gemini API Key。";
    return;
  }
  addUserMessage(userText, atts);
  const block = newAssistantBlock();
  const undoOps: UndoOp[] = [];
  const tctx: ToolContext = { undo: undoOps };
  let usage: UsageMetadata = {};
  setBusy(true);
  abort = new AbortController();
  try {
    if (!inExcel()) throw new Error("当前不在 Excel 环境中。");
    statusText.textContent = "读取待收票订单…";
    const layout = await inspectInvoiceSheet();
    const pendingOrders = layout.mode === "project" ? await readPendingOrders() : [];
    statusText.textContent = `识别 ${atts.length} 个文件${pendingOrders.length ? `（匹配 ${pendingOrders.length} 条待收票订单）` : ""}…`;
    const extra = userText === INVOICE_PROMPT ? "" : userText;
    const r = await extractInvoices(
      settings.apiKey,
      settings.model,
      settings.thinkingLevel,
      atts,
      extra,
      pendingOrders,
      abort.signal
    );
    usage = r.usage ?? {};
    if (!r.invoices.length) throw new Error("没有在附件中识别到发票。");
    statusText.textContent = "写入工作表…";
    const res = await writeInvoices(layout, r.invoices, atts, tctx);
    const md = el("div", "md");
    const summary = summarize(res);
    renderMarkdown(md, summary);
    block.appendChild(md);
    if (res.discrepancies?.length) renderDiscrepancies(block, res.discrepancies);
    // Give the chat model a compact text record (with exact discrepancy cell coords) for follow-up corrections.
    const diffNote = res.discrepancies?.length
      ? `\n金额差异明细（如需用户口头修改可直接定位单元格，无需重读整表）：\n` +
        res.discrepancies
          .map(
            (d) =>
              `- 订单 ${d.orderNo}（采购合计 ${d.orderAmount} vs 已收票 ${d.invoicedAmount}）：` +
              d.items.map((it) => `${d.orderSheet}!${d.costCol}${it.row} 序号${it.seq} ${it.name} 当前=${it.cost}`).join("；")
          )
          .join("\n")
      : "";
    pushUser([
      {
        text: `[系统记录] 用户通过“扫描发票”录入了发票，结果如下：\n${summary}${diffNote}\n识别数据：${JSON.stringify(r.invoices)}`,
      },
    ]);
    history.push({ role: "model", parts: [{ text: "已记录。" }] });
  } catch (e) {
    if ((e as Error).name === "AbortError") addError(block, "已停止。");
    else if (e instanceof GeminiError) addError(block, `Gemini API 错误：${e.message}`);
    else addError(block, `出错了：${(e as Error).message ?? e}`);
  } finally {
    setBusy(false);
    abort = null;
    renderTurnFooter(block, undoOps, usage, 1);
    refreshSelectionPill();
  }
}

function renderDiscrepancies(block: HTMLElement, list: OrderDiscrepancy[]) {
  const fmt = (n: number) => n.toLocaleString("zh-CN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  for (const d of list) {
    const card = el("div", "diff-card");
    const head = el("div", "diff-head");
    head.appendChild(el("span", "diff-title", `⚠️ 金额差异：订单 ${d.orderNo || "（未填单号）"}`));
    const jumpAddr =
      d.orderSheetRows.length === 1
        ? `${d.costCol}${d.orderSheetRows[0]}`
        : `${d.costCol}${d.orderSheetRows[0]}:${d.costCol}${d.orderSheetRows[d.orderSheetRows.length - 1]}`;
    const link = el("a", "range-link", `${d.orderSheet}!${jumpAddr}`);
    link.href = "#";
    link.onclick = (e) => {
      e.preventDefault();
      selectAddress(d.orderSheet, jumpAddr).catch(() => {});
    };
    head.appendChild(link);
    card.appendChild(head);

    const sign = d.diff > 0 ? "+" : "";
    card.appendChild(
      el(
        "div",
        "diff-info",
        `${d.itemSummary ? `${d.itemSummary} · ` : ""}采购订单 ${fmt(d.orderAmount)} vs 已收票 ${fmt(d.invoicedAmount)}（差额 ${sign}${fmt(d.diff)}）`
      )
    );

    const actions = el("div", "diff-actions");
    if (d.orderSheetRows.length === 1) {
      const btn = el(
        "button",
        "primary-btn small",
        `按发票金额修改采购订单（${fmt(d.orderAmount)} → ${fmt(d.invoicedAmount)}）`
      );
      btn.onclick = async () => {
        btn.disabled = true;
        btn.textContent = "修改中…";
        try {
          const { status, undo } = await updateOrderAmount(d);
          actions.innerHTML = "";
          actions.appendChild(
            el("span", "diff-ok", `✓ 已更新为 ${fmt(d.invoicedAmount)}${status ? `（${status}）` : ""}`)
          );
          const undoBtn = el("button", "small-btn", "↶ 撤销此修改");
          undoBtn.onclick = async () => {
            undoBtn.disabled = true;
            try {
              await undo();
              actions.innerHTML = "";
              btn.disabled = false;
              btn.textContent = `按发票金额修改采购订单（${fmt(d.orderAmount)} → ${fmt(d.invoicedAmount)}）`;
              actions.appendChild(btn);
            } catch (err) {
              undoBtn.disabled = false;
              undoBtn.textContent = `撤销失败：${(err as Error).message}`;
            }
          };
          actions.appendChild(undoBtn);
        } catch (err) {
          btn.disabled = false;
          btn.textContent = `修改失败，点此重试（${(err as Error).message}）`;
        }
      };
      actions.appendChild(btn);
    } else if (d.items.length > 1 && d.items.length <= 10) {
      const editor = el("div", "diff-rows");
      const inputs: { row: number; orig: number; input: HTMLInputElement }[] = [];
      for (const it of d.items) {
        const rowEl = el("div", "diff-row");
        const lbl = el("span", "diff-row-lbl", `序号${it.seq} ${it.name || `第${it.row}行`}`);
        lbl.title = `${d.orderSheet}!${d.costCol}${it.row} ${it.name}`;
        const inp = el("input", "diff-num") as HTMLInputElement;
        inp.type = "number";
        inp.step = "0.01";
        inp.value = String(it.cost);
        inputs.push({ row: it.row, orig: it.cost, input: inp });
        rowEl.append(lbl, inp);
        editor.appendChild(rowEl);
      }
      card.appendChild(editor);

      const sumEl = el("span", "diff-sum");
      const updateSum = () => {
        const s = inputs.reduce((acc, x) => acc + (Number(x.input.value) || 0), 0);
        const delta = s - d.invoicedAmount;
        sumEl.textContent =
          Math.abs(delta) < 0.005
            ? `合计 ${fmt(s)}（与发票一致 ✓）`
            : `合计 ${fmt(s)}（距发票差 ${delta > 0 ? "+" : ""}${fmt(delta)}）`;
        sumEl.className = Math.abs(delta) < 0.005 ? "diff-sum ok" : "diff-sum";
      };
      inputs.forEach((x) => x.input.addEventListener("input", updateSum));
      updateSum();

      const saveBtn = el("button", "primary-btn small", "保存修改（免AI费用）");
      saveBtn.onclick = async () => {
        const updates = inputs
          .map((x) => ({ row: x.row, cost: Number(x.input.value), orig: x.orig }))
          .filter((x) => isFinite(x.cost) && Math.abs(x.cost - x.orig) >= 0.0001);
        if (!updates.length) {
          sumEl.textContent = "未修改任何金额";
          return;
        }
        saveBtn.disabled = true;
        saveBtn.textContent = "保存中…";
        try {
          const { status, undo } = await updateOrderRows(
            d,
            updates.map((u) => ({ row: u.row, cost: u.cost }))
          );
          inputs.forEach((x) => (x.input.disabled = true));
          actions.innerHTML = "";
          actions.appendChild(
            el("span", "diff-ok", `✓ 已更新 ${updates.length} 行${status ? `（${status}）` : ""}`)
          );
          const undoBtn = el("button", "small-btn", "↶ 撤销此修改");
          undoBtn.onclick = async () => {
            undoBtn.disabled = true;
            try {
              await undo();
              inputs.forEach((x) => {
                x.input.disabled = false;
                x.input.value = String(x.orig);
              });
              updateSum();
              actions.innerHTML = "";
              saveBtn.disabled = false;
              saveBtn.textContent = "保存修改（免AI费用）";
              actions.append(sumEl, saveBtn);
            } catch (err) {
              undoBtn.disabled = false;
              undoBtn.textContent = `撤销失败：${(err as Error).message}`;
            }
          };
          actions.appendChild(undoBtn);
        } catch (err) {
          saveBtn.disabled = false;
          saveBtn.textContent = `保存失败，重试（${(err as Error).message}）`;
        }
      };
      actions.append(sumEl, saveBtn);
    } else {
      actions.appendChild(
        el("span", "hint", `该订单包含 ${d.orderSheetRows.length} 个产品行（可能是分批开票），请点右上角地址手动核对。`)
      );
    }
    card.appendChild(actions);
    block.appendChild(card);
  }
  scrollToBottom();
}

// USD per 1M tokens [input, output incl. thinking]
const PRICES: Record<string, [number, number]> = {
  "gemini-3.8-flash": [0.75, 3.75],
  "gemini-4-hydrogen": [1.5, 7.5],
  "gemini-4-neon": [3, 15],
};

function renderTurnFooter(block: HTMLElement, undoOps: UndoOp[], usage: UsageMetadata, calls = 1) {
  const foot = el("div", "turn-footer");
  if (undoOps.length) {
    const btn = el("button", "small-btn", `↶ 撤销本轮修改（${undoOps.length} 项）`);
    let remaining = [...undoOps].reverse();
    let errBox: HTMLElement | null = null;
    btn.onclick = async () => {
      if (busy) {
        statusText.textContent = "请等当前任务结束后再撤销";
        return;
      }
      btn.disabled = true;
      btn.textContent = "撤销中…";
      const failed: UndoOp[] = [];
      const errors: string[] = [];
      for (const op of remaining) {
        try {
          await op();
        } catch (e) {
          failed.push(op);
          errors.push((e as Error).message || String(e));
        }
      }
      errBox?.remove();
      errBox = null;
      if (failed.length) {
        remaining = failed;
        btn.disabled = false;
        btn.textContent = `↶ 重试撤销（${failed.length} 项失败）`;
        errBox = el("div", "error-box", `撤销失败：\n${errors.join("\n")}`);
        errBox.style.whiteSpace = "pre-wrap";
        foot.after(errBox);
      } else {
        btn.textContent = "✓ 已撤销";
        pushUser([{ text: "[系统提示] 用户已撤销了你上一轮对工作簿所做的全部修改。" }]);
      }
      refreshSelectionPill();
    };
    foot.appendChild(btn);
  }
  const total = (usage.promptTokenCount ?? 0) + (usage.candidatesTokenCount ?? 0) + (usage.thoughtsTokenCount ?? 0);
  if (total) {
    const price = PRICES[settings.model];
    const cost = price
      ? ((usage.promptTokenCount ?? 0) * price[0] +
          ((usage.candidatesTokenCount ?? 0) + (usage.thoughtsTokenCount ?? 0)) * price[1]) /
        1e6
      : null;
    foot.appendChild(
      el(
        "span",
        "usage",
        `${settings.model} · ${calls} 次调用 · 输入 ${usage.promptTokenCount} · 输出 ${usage.candidatesTokenCount}${
          usage.thoughtsTokenCount ? ` · 思考 ${usage.thoughtsTokenCount}` : ""
        } tokens${cost !== null ? ` · 约 $${cost.toFixed(3)}` : ""}`
      )
    );
  }
  if (foot.childNodes.length) block.appendChild(foot);
  scrollToBottom();
}

/* ------------------------------ wiring ------------------------------ */

function send() {
  let text = inputEl.value.trim();
  if (busy || (!text && !attachments.length)) return;
  const atts = attachments;
  attachments = [];
  renderAttachStrip();
  inputEl.value = "";
  autosize();
  // Attachments with no instruction (or the invoice button) use the cheap one-call invoice path.
  if (atts.length && (!text || text === INVOICE_PROMPT || forceInvoice)) runInvoiceTurn(text || INVOICE_PROMPT, atts);
  else runTurn(text, atts);
  forceInvoice = false;
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

let sendAfterPick = false;
let forceInvoice = false;

function bindAttachments() {
  const fileInput = $<HTMLInputElement>("fileInput");
  $("attachBtn").onclick = () => fileInput.click();
  fileInput.onchange = async () => {
    if (fileInput.files?.length) await addFiles(fileInput.files);
    fileInput.value = "";
    if (sendAfterPick && attachments.length) {
      forceInvoice = true;
      send();
    }
    sendAfterPick = false;
  };
  $("invoiceChip").onclick = () => {
    sendAfterPick = true;
    fileInput.click();
  };
  inputEl.addEventListener("paste", (e) => {
    const files = Array.from(e.clipboardData?.files ?? []);
    if (files.length) {
      e.preventDefault();
      addFiles(files);
    }
  });
  const composer = document.querySelector<HTMLElement>(".composer")!;
  document.addEventListener("dragover", (e) => {
    e.preventDefault();
    composer.classList.add("dragging");
  });
  document.addEventListener("dragleave", (e) => {
    if (!e.relatedTarget) composer.classList.remove("dragging");
  });
  document.addEventListener("drop", (e) => {
    e.preventDefault();
    composer.classList.remove("dragging");
    if (e.dataTransfer?.files.length) addFiles(e.dataTransfer.files);
  });
}

function init() {
  bindSettings();
  bindAttachments();
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
  document.querySelectorAll<HTMLButtonElement>(".chip:not(.invoice)").forEach((c) => {
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
