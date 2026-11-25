// Low-cost invoice pipeline: ONE structured-output Gemini call extracts all invoices,
// then plain Office.js code writes them to the sheet (no agent loop, no repeated image uploads).
import { Content, Part, UsageMetadata, streamGenerate } from "./gemini";
import { Attachment } from "./attachments";
import { ToolContext, colLetter, executeTool, pushSnapshot } from "./tools";

export const INVOICE_SHEET = "收票登记";

export const HEADERS = [
  "序号", "发票类型", "发票代码", "发票号码", "开票日期", "购买方名称", "购买方税号",
  "销售方名称", "销售方税号", "项目名称", "金额(不含税)", "税率", "税额", "价税合计",
  "校验", "来源文件", "备注",
] as const;

export interface InvoiceRecord {
  file_index: number;
  invoice_type: string | null;
  invoice_code: string | null;
  invoice_number: string | null;
  issue_date: string | null;
  buyer_name: string | null;
  buyer_tax_id: string | null;
  seller_name: string | null;
  seller_tax_id: string | null;
  items: string | null;
  amount: number | null;
  tax_rate: string | null;
  tax: number | null;
  total: number | null;
  note: string | null;
}

const S = { type: "STRING", nullable: true };
const N = { type: "NUMBER", nullable: true };

const SCHEMA = {
  type: "OBJECT",
  properties: {
    invoices: {
      type: "ARRAY",
      items: {
        type: "OBJECT",
        properties: {
          file_index: { type: "INTEGER", description: "来自第几个附件（从1开始）" },
          invoice_type: { ...S, description: "如 电子发票（增值税专用发票）/ 电子发票（普通发票）" },
          invoice_code: { ...S, description: "发票代码，全电发票没有则为 null" },
          invoice_number: { ...S, description: "发票号码，原样保留所有数字" },
          issue_date: { ...S, description: "开票日期 yyyy-mm-dd" },
          buyer_name: S,
          buyer_tax_id: { ...S, description: "购买方统一社会信用代码/纳税人识别号" },
          seller_name: S,
          seller_tax_id: S,
          items: { ...S, description: "项目名称，多项用；连接" },
          amount: { ...N, description: "合计金额（不含税）" },
          tax_rate: { ...S, description: "税率，如 13% / 6% / 免税；多种税率用 / 连接" },
          tax: { ...N, description: "合计税额" },
          total: { ...N, description: "价税合计（小写）" },
          note: { ...S, description: "看不清或不确定的字段说明；正常则为 null" },
        },
        required: ["file_index", "invoice_number", "total"],
      },
    },
  },
  required: ["invoices"],
};

const PROMPT = `识别以下附件中的所有发票（一个文件可能包含多张），逐字段提取。
规则：看不清或不存在的字段返回 null，并在 note 里说明，绝不猜测号码或金额；金额为纯数字；日期为 yyyy-mm-dd。`;

/** Gemini 3+/4 accept thinkingLevel "low"; older models get no thinking config. */
function cheapThinking(model: string, userLevel: string): string {
  if (/gemini-(3|4)/.test(model)) return "low";
  return userLevel === "" ? "" : userLevel;
}

export async function extractInvoices(
  apiKey: string,
  model: string,
  userThinking: string,
  atts: Attachment[],
  extra: string,
  signal?: AbortSignal
): Promise<{ invoices: InvoiceRecord[]; usage?: UsageMetadata }> {
  const parts: Part[] = [{ text: PROMPT + (extra ? `\n用户补充要求：${extra}` : "") }];
  atts.forEach((a, i) => {
    parts.push({ text: `附件${i + 1}：${a.name}` });
    parts.push({ inlineData: { mimeType: a.mimeType, data: a.data } });
  });
  const contents: Content[] = [{ role: "user", parts }];
  let text = "";
  let usage: UsageMetadata | undefined;
  for await (const c of streamGenerate({
    apiKey,
    model,
    thinkingLevel: cheapThinking(model, userThinking),
    responseSchema: SCHEMA,
    contents,
    signal,
  })) {
    for (const p of c.parts) if (p.text && !p.thought) text += p.text;
    if (c.usage) usage = c.usage;
  }
  const json = JSON.parse(text.replace(/^```(?:json)?\s*|\s*```$/g, ""));
  return { invoices: (json.invoices ?? []) as InvoiceRecord[], usage };
}

function parseRate(r: string | null): number | string {
  if (!r) return "";
  const m = r.match(/^\s*(\d+(?:\.\d+)?)\s*%\s*$/);
  return m ? Number(m[1]) / 100 : r;
}

const txt = (v: string | null) => (v ? `'${v}` : "");
const num = (v: number | null) => (typeof v === "number" && isFinite(v) ? v : "");

export interface WriteResult {
  sheet: string;
  address: string | null;
  written: InvoiceRecord[];
  duplicates: InvoiceRecord[];
  created: boolean;
}

export async function writeInvoices(
  invoices: InvoiceRecord[],
  atts: Attachment[],
  tctx: ToolContext
): Promise<WriteResult> {
  // 1. Inspect target sheet.
  const info = await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItemOrNullObject(INVOICE_SHEET);
    await ctx.sync();
    if (ws.isNullObject) return null;
    const used = ws.getUsedRangeOrNullObject(true);
    used.load("rowCount,columnCount,rowIndex,columnIndex,values");
    await ctx.sync();
    if (used.isNullObject) return { headers: [] as string[], lastRow: -1, values: [] as unknown[][], col0: 0 };
    return {
      headers: (used.values[0] ?? []).map((v) => String(v).trim()),
      lastRow: used.rowIndex + used.rowCount - 1,
      values: used.values,
      col0: used.columnIndex,
    };
  });

  let created = false;
  if (!info) {
    await executeTool("add_sheet", { name: INVOICE_SHEET }, tctx);
    created = true;
  }
  const isFresh = !info || info.headers.every((h) => !h);
  const headers: string[] = isFresh ? [...HEADERS] : info!.headers;
  const col0 = isFresh ? 0 : info!.col0;
  const col = (h: string) => headers.indexOf(h);

  // 2. De-duplicate by invoice number against existing rows and within this batch.
  const numCol = col("发票号码");
  const existing = new Set<string>();
  if (!isFresh && numCol >= 0)
    info!.values.slice(1).forEach((r) => {
      const v = String(r[numCol] ?? "").trim();
      if (v) existing.add(v);
    });
  const written: InvoiceRecord[] = [];
  const duplicates: InvoiceRecord[] = [];
  for (const inv of invoices) {
    const n = (inv.invoice_number ?? "").trim();
    if (n && existing.has(n)) duplicates.push(inv);
    else {
      if (n) existing.add(n);
      written.push(inv);
    }
  }

  const firstDataRow = isFresh ? 1 : info!.lastRow + 1; // 0-based
  const serialStart = isFresh ? 1 : Math.max(1, info!.values.length);

  // 3. Build rows following the sheet's own header order.
  const L = (h: string) => (col(h) >= 0 ? colLetter(col0 + col(h)) : null);
  const rows = written.map((inv, i) => {
    const excelRow = firstDataRow + i + 1;
    const amt = L("金额(不含税)"), tax = L("税额"), tot = L("价税合计");
    const check =
      amt && tax && tot
        ? `=IF(OR(${amt}${excelRow}="",${tot}${excelRow}=""),"",IF(ABS(N(${amt}${excelRow})+N(${tax}${excelRow})-${tot}${excelRow})<0.01,"✓","需核对"))`
        : "";
    const map: Record<string, unknown> = {
      序号: serialStart + i,
      发票类型: inv.invoice_type ?? "",
      发票代码: txt(inv.invoice_code),
      发票号码: txt(inv.invoice_number),
      开票日期: inv.issue_date ?? "",
      购买方名称: inv.buyer_name ?? "",
      购买方税号: txt(inv.buyer_tax_id),
      销售方名称: inv.seller_name ?? "",
      销售方税号: txt(inv.seller_tax_id),
      项目名称: inv.items ?? "",
      "金额(不含税)": num(inv.amount),
      税率: parseRate(inv.tax_rate),
      税额: num(inv.tax),
      价税合计: num(inv.total),
      校验: check,
      来源文件: atts[inv.file_index - 1]?.name ?? "",
      备注: inv.note ?? "",
    };
    return headers.map((h) => (h in map ? map[h] : ""));
  });

  if (!rows.length && !isFresh) return { sheet: INVOICE_SHEET, address: null, written, duplicates, created };

  // 4. Write (header too if fresh) with undo snapshot.
  const block = isFresh ? [headers as unknown[], ...rows] : rows;
  const startRow = isFresh ? 0 : firstDataRow;
  const address = `${colLetter(col0)}${startRow + 1}:${colLetter(col0 + headers.length - 1)}${startRow + block.length}`;
  await pushSnapshot(tctx, INVOICE_SHEET, address);

  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(INVOICE_SHEET);
    const r = ws.getRange(address);
    r.formulas = block.map((row) => row.map((v) => (v === null || v === undefined ? "" : v))) as any;
    const dataRows = rows.length;
    const dataStart = isFresh ? 2 : firstDataRow + 1;
    const colRange = (h: string, fmt: string) => {
      const l = L(h);
      if (l && dataRows) ws.getRange(`${l}${dataStart}:${l}${dataStart + dataRows - 1}`).numberFormat = [[fmt]];
    };
    ["金额(不含税)", "税额", "价税合计"].forEach((h) => colRange(h, "#,##0.00"));
    colRange("税率", "0%");
    colRange("开票日期", "yyyy-mm-dd");
    if (isFresh) {
      const head = ws.getRange(`${colLetter(col0)}1:${colLetter(col0 + headers.length - 1)}1`);
      head.format.font.bold = true;
      head.format.fill.color = "#DDEBF7";
      ws.freezePanes.freezeRows(1);
    }
    r.format.autofitColumns();
    ws.activate();
    r.select();
    await ctx.sync();
  });

  return { sheet: INVOICE_SHEET, address, written, duplicates, created };
}

const fmtMoney = (v: number | null) => (typeof v === "number" ? v.toLocaleString("zh-CN", { minimumFractionDigits: 2 }) : "—");

export function summarize(res: WriteResult): string {
  const lines: string[] = [];
  if (res.written.length) {
    lines.push(`已录入 **${res.written.length}** 张发票到「${res.sheet}」${res.address ? `（${res.address}）` : ""}：`, "");
    lines.push("| 发票号码 | 销售方 | 价税合计 | 备注 |", "|---|---|---:|---|");
    for (const i of res.written)
      lines.push(`| ${i.invoice_number ?? "—"} | ${i.seller_name ?? "—"} | ${fmtMoney(i.total)} | ${i.note ?? ""} |`);
    const check = res.written.filter(
      (i) => i.note || typeof i.amount !== "number" || typeof i.total !== "number" ||
        (typeof i.tax === "number" && Math.abs((i.amount ?? 0) + i.tax - (i.total ?? 0)) >= 0.01)
    );
    if (check.length) lines.push("", `⚠️ ${check.length} 张需要人工核对（见“备注”和“校验”列）。`);
  } else lines.push("没有新的发票需要录入。");
  if (res.duplicates.length)
    lines.push("", `已跳过 ${res.duplicates.length} 张重复发票：${res.duplicates.map((d) => d.invoice_number).join("、")}`);
  return lines.join("\n");
}
