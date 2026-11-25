// Low-cost invoice pipeline:
// 1. Inspects 「收票登记」 (and pending rows of 「采购订单」 if present) locally via Office.js.
// 2. Makes ONE structured-output Gemini call (low thinking) to extract invoice fields and match pending orders.
// 3. Writes directly into Excel with undo snapshots — in project-template mode it only touches input columns
//    (A, B, D, G, H, M, N, O) starting at the first blank data row, preserving all pre-filled formulas.
import { Content, Part, UsageMetadata, streamGenerate } from "./gemini";
import { Attachment } from "./attachments";
import { ToolContext, colLetter, executeTool, pushSnapshot } from "./tools";

export const INVOICE_SHEET = "收票登记";
const ORDER_SHEET = "采购订单";

export const HEADERS = [
  "序号", "发票类型", "发票代码", "发票号码", "开票日期", "购买方名称", "购买方税号",
  "销售方名称", "销售方税号", "项目名称", "金额(不含税)", "税率", "税额", "价税合计",
  "校验", "来源文件", "备注",
] as const;

export interface MatchedOrder {
  order_no: string;
  allocated_amount: number | null;
  platform: string | null;
  ticket_type: string | null;
  match_note: string | null;
}

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
  matched_orders?: MatchedOrder[] | null;
}

export interface PendingOrder {
  seq: string;
  item: string;
  platform: string;
  orderNo: string;
  qty: string;
  supplier: string;
  ticketType: string;
  cost: string;
  remaining: string;
  status: string;
}

interface TemplateLayout {
  mode: "project";
  headerRow: number; // 0-based
  blankRows: number[]; // 0-based empty data rows in ascending order
  maxRow: number; // 0-based inclusive
  existingNumbers: Set<string>;
  cols: {
    number: number;
    date: number;
    order: number;
    total: number;
    alloc: number;
    note: number;
    platform: number;
    ticketType: number;
    check: number;
  };
}

interface GenericLayout {
  mode: "generic";
  exists: boolean;
  isFresh: boolean;
  headers: string[];
  col0: number;
  lastRow: number;
  dataCount: number;
  existingNumbers: Set<string>;
}

export type SheetLayout = TemplateLayout | GenericLayout;

const S = { type: "STRING", nullable: true };
const N = { type: "NUMBER", nullable: true };

function buildSchema(withOrders: boolean) {
  const props: Record<string, unknown> = {
    file_index: { type: "INTEGER", description: "来自第几个附件（从1开始）" },
    invoice_type: { ...S, description: "如 电子发票（增值税专用发票）/ 电子发票（普通发票）" },
    invoice_code: { ...S, description: "发票代码，全电发票没有则为 null" },
    invoice_number: { ...S, description: "发票号码，原样保留所有数字" },
    issue_date: { ...S, description: "开票日期 yyyy-mm-dd" },
    buyer_name: S,
    buyer_tax_id: { ...S, description: "购买方统一社会信用代码/纳税人识别号" },
    seller_name: S,
    seller_tax_id: S,
    items: { ...S, description: "项目名称、规格及数量，多项用；连接" },
    amount: { ...N, description: "合计金额（不含税）" },
    tax_rate: { ...S, description: "税率，如 13% / 6% / 1% / 免税；多种税率用 / 连接" },
    tax: { ...N, description: "合计税额" },
    total: { ...N, description: "价税合计（小写）" },
    note: { ...S, description: "看不清或不确定的字段说明；正常则为 null" },
  };
  if (withOrders) {
    props.matched_orders = {
      type: "ARRAY",
      nullable: true,
      description:
        "从“待收票采购订单列表”中匹配到的订单。通常 1 张票对应 1 个订单（allocated_amount 填 null）；如果 1 张票跨多个订单则按订单拆成多项，每项填 allocated_amount 且合计等于价税合计；找不到确定匹配时返回空数组 []。",
      items: {
        type: "OBJECT",
        properties: {
          order_no: { type: "STRING", description: "必须原样复制候选列表中的平台订单号" },
          allocated_amount: { ...N, description: "仅当 1 张票拆给多个订单时填写本订单分配额，单订单票填 null" },
          platform: { ...S, description: "仅当同一订单号在列表中对应多个平台时填写，否则填 null" },
          ticket_type: { ...S, description: "仅当同一订单号在列表中对应多种票种时填写，否则填 null" },
          match_note: { ...S, description: "简述匹配依据（如匹配采购订单序号XX及金额差异），无差异可简写" },
        },
        required: ["order_no"],
      },
    };
  }
  return {
    type: "OBJECT",
    properties: {
      invoices: {
        type: "ARRAY",
        items: { type: "OBJECT", properties: props, required: ["file_index", "invoice_number", "total"] },
      },
    },
    required: ["invoices"],
  };
}

/** Finds header row (scanning top 10 rows) and detects whether 「收票登记」 uses the project template. */
export async function inspectInvoiceSheet(): Promise<SheetLayout> {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItemOrNullObject(INVOICE_SHEET);
    await ctx.sync();
    if (ws.isNullObject) {
      return {
        mode: "generic",
        exists: false,
        isFresh: true,
        headers: [...HEADERS],
        col0: 0,
        lastRow: -1,
        dataCount: 0,
        existingNumbers: new Set(),
      };
    }
    const used = ws.getUsedRangeOrNullObject(true);
    used.load("rowCount,columnCount,rowIndex,columnIndex");
    await ctx.sync();
    if (used.isNullObject) {
      return {
        mode: "generic",
        exists: true,
        isFresh: true,
        headers: [...HEADERS],
        col0: 0,
        lastRow: -1,
        dataCount: 0,
        existingNumbers: new Set(),
      };
    }

    // Read top 10 rows to detect the header row.
    const probeRows = Math.min(10, used.rowCount);
    const probeCols = Math.min(26, used.columnCount);
    const probe = used.getCell(0, 0).getResizedRange(probeRows - 1, probeCols - 1);
    probe.load("values");
    await ctx.sync();

    let projectHeaderOffset = -1;
    let headerTexts: string[] = [];
    for (let r = 0; r < probe.values.length; r++) {
      const row = probe.values[r].map((v) => String(v ?? "").trim());
      if (row.some((h) => h.includes("发票号码")) && row.some((h) => h.includes("平台订单号")) && row.some((h) => h.includes("整张票金额"))) {
        projectHeaderOffset = r;
        headerTexts = row;
        break;
      }
    }

    if (projectHeaderOffset >= 0) {
      const findCol = (kw: string) => {
        const i = headerTexts.findIndex((h) => h.includes(kw));
        return i >= 0 ? used.columnIndex + i : -1;
      };
      const cols = {
        number: findCol("发票号码"),
        date: findCol("开票日期"),
        order: findCol("平台订单号"),
        total: findCol("整张票金额"),
        alloc: findCol("本订单分配额"),
        note: findCol("备注"),
        platform: findCol("匹配平台"),
        ticketType: findCol("匹配票种"),
        check: findCol("发票分配校验"),
      };
      const headerRow = used.rowIndex + projectHeaderOffset;
      const dataStart = headerRow + 1;
      const maxRow = used.rowIndex + used.rowCount - 1;
      const numRows = Math.max(0, maxRow - dataStart + 1);
      const existingNumbers = new Set<string>();
      const blankRows: number[] = [];

      if (numRows > 0) {
        // Check input columns (A 票号, D 订单号, G 整张票金额, H 分配额) so a row with any data is never overwritten.
        const checkCols = [cols.number, cols.order, cols.total, cols.alloc].filter((c) => c >= 0);
        const ranges = checkCols.map((c) => {
          const rg = ws.getRangeByIndexes(dataStart, c, numRows, 1);
          rg.load("values");
          return rg;
        });
        await ctx.sync();
        for (let i = 0; i < numRows; i++) {
          const a = String(ranges[0].values[i][0] ?? "").trim();
          if (a) existingNumbers.add(a);
          const isEmpty = ranges.every((rg) => String(rg.values[i][0] ?? "").trim() === "");
          if (isEmpty) blankRows.push(dataStart + i);
        }
      }
      return { mode: "project", headerRow, blankRows, maxRow, existingNumbers, cols };
    }

    // Generic mode fallback
    used.load("values");
    await ctx.sync();
    const headers = (used.values[0] ?? []).map((v) => String(v).trim());
    const isFresh = headers.every((h) => !h);
    const existingNumbers = new Set<string>();
    const numCol = headers.indexOf("发票号码");
    if (!isFresh && numCol >= 0) {
      used.values.slice(1).forEach((r) => {
        const v = String(r[numCol] ?? "").trim();
        if (v) existingNumbers.add(v);
      });
    }
    return {
      mode: "generic",
      exists: true,
      isFresh,
      headers: isFresh ? [...HEADERS] : headers,
      col0: isFresh ? 0 : used.columnIndex,
      lastRow: used.rowIndex + used.rowCount - 1,
      dataCount: Math.max(0, used.values.length - 1),
      existingNumbers,
    };
  });
}

/** Reads only pending (un-invoiced / not yet reconciled) rows from 「采购订单」. */
export async function readPendingOrders(): Promise<PendingOrder[]> {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItemOrNullObject(ORDER_SHEET);
    await ctx.sync();
    if (ws.isNullObject) return [];
    const used = ws.getUsedRangeOrNullObject(true);
    used.load("rowCount,columnCount,rowIndex,columnIndex");
    await ctx.sync();
    if (used.isNullObject) return [];

    // Read columns A..W (up to 23 cols) across used rows; filter locally in JS.
    const cols = Math.min(23, used.columnCount);
    const r = used.getCell(0, 0).getResizedRange(used.rowCount - 1, cols - 1);
    r.load("values");
    await ctx.sync();

    let hRow = -1;
    let headers: string[] = [];
    for (let i = 0; i < Math.min(10, r.values.length); i++) {
      const row = r.values[i].map((v) => String(v ?? "").trim());
      if (row.some((h) => h.includes("平台订单号")) && row.some((h) => h.includes("含税成本"))) {
        hRow = i;
        headers = row;
        break;
      }
    }
    if (hRow < 0) return [];
    const idx = (kw: string) => headers.findIndex((h) => h.includes(kw));
    const cSeq = idx("序号");
    const cItem = idx("品名");
    const cPlat = idx("采购平台");
    const cOrder = idx("平台订单号");
    const cQty = idx("采购数量");
    const cSup = idx("商家");
    const cType = idx("成本类别");
    const cCost = idx("含税成本");
    const cRem = idx("成本减已收票");
    const cStatus = idx("核对状态");
    const cNeed = idx("需收票");
    if (cOrder < 0) return [];

    const out: PendingOrder[] = [];
    for (let i = hRow + 1; i < r.values.length; i++) {
      const row = r.values[i];
      const seq = cSeq >= 0 ? String(row[cSeq] ?? "").trim() : "";
      const orderNo = String(row[cOrder] ?? "").trim();
      if (!seq && !orderNo) continue;
      if (cNeed >= 0 && String(row[cNeed] ?? "").trim() === "否") continue;
      const status = cStatus >= 0 ? String(row[cStatus] ?? "").trim() : "";
      if (status === "已对平") continue;
      out.push({
        seq,
        item: cItem >= 0 ? String(row[cItem] ?? "").trim() : "",
        platform: cPlat >= 0 ? String(row[cPlat] ?? "").trim() : "",
        orderNo,
        qty: cQty >= 0 ? String(row[cQty] ?? "").trim() : "",
        supplier: cSup >= 0 ? String(row[cSup] ?? "").trim() : "",
        ticketType: cType >= 0 ? String(row[cType] ?? "").trim() : "",
        cost: cCost >= 0 ? String(row[cCost] ?? "").trim() : "",
        remaining: cRem >= 0 ? String(row[cRem] ?? "").trim() : "",
        status,
      });
    }
    return out;
  });
}

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
  pendingOrders: PendingOrder[],
  signal?: AbortSignal
): Promise<{ invoices: InvoiceRecord[]; usage?: UsageMetadata }> {
  let prompt = `识别以下附件中的所有发票（一个文件可能包含多张），逐字段提取。
规则：看不清或不存在的字段返回 null，并在 note 里说明，绝不猜测号码或金额；金额为纯数字；日期为 yyyy-mm-dd。`;

  if (pendingOrders.length) {
    const lines = pendingOrders.map(
      (o) =>
        `序号${o.seq} | ${o.item} | 数量:${o.qty} | 平台:${o.platform} | 订单号:${o.orderNo} | 票种:${o.ticketType} | 含税成本:${o.cost} | 未收票:${o.remaining} | 状态:${o.status}`
    );
    prompt += `\n\n以下是当前「采购订单」中尚未收齐发票的候选订单，请根据品名、规格、数量、金额、税率/票种为每张发票匹配对应的平台订单号（放入 matched_orders；若金额略有差异也可暂匹配并在 match_note 说明差额；若完全无匹配则 matched_orders 返空数组）：\n${lines.join("\n")}`;
  }
  if (extra) prompt += `\n\n用户补充要求：${extra}`;

  const parts: Part[] = [{ text: prompt }];
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
    responseSchema: buildSchema(pendingOrders.length > 0),
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

/** Converts yyyy-mm-dd to Excel date serial number (1900 date system). */
function dateToExcelSerial(iso: string | null): number | string {
  if (!iso) return "";
  const m = iso.trim().match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (!m) return iso;
  const utc = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return Math.round((utc - Date.UTC(1899, 11, 30)) / 86400000);
}

const txt = (v: string | null | undefined) => (v ? `'${v}` : "");
const num = (v: number | null | undefined) => (typeof v === "number" && isFinite(v) ? v : "");

function buildProjectNote(inv: InvoiceRecord, matchNote: string | null | undefined, fileName: string): string {
  const parts: string[] = [];
  const head = [inv.invoice_type || "发票", inv.invoice_number || ""].filter(Boolean).join("");
  if (head) parts.push(head);
  if (inv.issue_date) parts.push(`开票日期${inv.issue_date}`);
  if (inv.seller_name) parts.push(`开票方${inv.seller_name}`);
  const detail: string[] = [];
  if (inv.items) detail.push(inv.items);
  if (typeof inv.amount === "number") detail.push(`金额${inv.amount.toFixed(2)}元`);
  if (typeof inv.tax === "number") detail.push(`税额${inv.tax.toFixed(2)}元${inv.tax_rate ? `（${inv.tax_rate}）` : ""}`);
  if (typeof inv.total === "number") detail.push(`价税合计${inv.total.toFixed(2)}元`);
  if (detail.length) parts.push(detail.join("，"));
  if (matchNote) parts.push(matchNote);
  if (inv.note) parts.push(`核对提示：${inv.note}`);
  if (!matchNote && fileName) parts.push(`文件：${fileName}`);
  return parts.join("；");
}

export interface WrittenRowSummary {
  rowNum: number; // 1-based Excel row
  invoiceNumber: string;
  seller: string;
  orderNo: string;
  total: number | null;
  checkStatus: string;
  note: string;
}

export interface WriteResult {
  sheet: string;
  address: string | null;
  rows: WrittenRowSummary[];
  duplicates: InvoiceRecord[];
  created: boolean;
}

async function writeProjectInvoices(
  layout: TemplateLayout,
  invoices: InvoiceRecord[],
  atts: Attachment[],
  tctx: ToolContext
): Promise<WriteResult> {
  const duplicates: InvoiceRecord[] = [];
  const toWrite: {
    inv: InvoiceRecord;
    orderNo: string;
    total: number | "";
    alloc: number | "";
    platform: string;
    ticketType: string;
    note: string;
  }[] = [];

  for (const inv of invoices) {
    const n = (inv.invoice_number ?? "").trim();
    if (n && layout.existingNumbers.has(n)) {
      duplicates.push(inv);
      continue;
    }
    if (n) layout.existingNumbers.add(n);
    const fileName = atts[inv.file_index - 1]?.name ?? "";
    const matches = (inv.matched_orders ?? []).filter((m) => m.order_no?.trim());
    if (matches.length <= 1) {
      const m = matches[0];
      toWrite.push({
        inv,
        orderNo: m?.order_no?.trim() ?? "",
        total: num(inv.total),
        alloc: "",
        platform: m?.platform?.trim() ?? "",
        ticketType: m?.ticket_type?.trim() ?? "",
        note: buildProjectNote(inv, m?.match_note, fileName),
      });
    } else {
      matches.forEach((m, idx) => {
        toWrite.push({
          inv,
          orderNo: m.order_no.trim(),
          total: idx === 0 ? num(inv.total) : "",
          alloc: num(m.allocated_amount),
          platform: m.platform?.trim() ?? "",
          ticketType: m.ticket_type?.trim() ?? "",
          note: buildProjectNote(inv, m.match_note, fileName),
        });
      });
    }
  }

  if (!toWrite.length) return { sheet: INVOICE_SHEET, address: null, rows: [], duplicates, created: false };

  // Assign each row to the nearest blank row (falling back to maxRow + 1 only if template rows are all full).
  let overflow = 0;
  const targetRows = toWrite.map((_, i) =>
    i < layout.blankRows.length ? layout.blankRows[i] : layout.maxRow + 1 + overflow++
  );
  const { cols } = layout;

  // Group consecutive target rows into contiguous chunks (usually just 1 chunk).
  const chunks: { startRow: number; indices: number[] }[] = [];
  targetRows.forEach((rowIdx, i) => {
    const prev = chunks[chunks.length - 1];
    if (prev && prev.startRow + prev.indices.length === rowIdx) prev.indices.push(i);
    else chunks.push({ startRow: rowIdx, indices: [i] });
  });

  const buildColWrites = (indices: number[]) => {
    const sub = indices.map((i) => toWrite[i]);
    const list: { col: number; values: (string | number)[][] }[] = [
      { col: cols.number, values: sub.map((r) => [txt(r.inv.invoice_number)]) },
      { col: cols.date, values: sub.map((r) => [dateToExcelSerial(r.inv.issue_date)]) },
      { col: cols.order, values: sub.map((r) => [txt(r.orderNo)]) },
      { col: cols.total, values: sub.map((r) => [r.total]) },
      { col: cols.alloc, values: sub.map((r) => [r.alloc]) },
      { col: cols.note, values: sub.map((r) => [r.note]) },
    ];
    if (cols.platform >= 0 && sub.some((r) => r.platform))
      list.push({ col: cols.platform, values: sub.map((r) => [r.platform]) });
    if (cols.ticketType >= 0 && sub.some((r) => r.ticketType))
      list.push({ col: cols.ticketType, values: sub.map((r) => [r.ticketType]) });
    return list;
  };

  for (const ch of chunks) {
    for (const cw of buildColWrites(ch.indices)) {
      if (cw.col < 0) continue;
      const L = colLetter(cw.col);
      const addr = `${L}${ch.startRow + 1}:${L}${ch.startRow + ch.indices.length}`;
      await pushSnapshot(tctx, INVOICE_SHEET, addr);
    }
  }

  const checkStatuses = await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(INVOICE_SHEET);
    for (const ch of chunks) {
      for (const cw of buildColWrites(ch.indices)) {
        if (cw.col < 0) continue;
        const rg = ws.getRangeByIndexes(ch.startRow, cw.col, ch.indices.length, 1);
        rg.formulas = cw.values as any;
      }
    }
    const lastCh = chunks[chunks.length - 1];
    const sel = ws.getRangeByIndexes(
      lastCh.startRow,
      cols.number,
      lastCh.indices.length,
      Math.max(1, cols.note - cols.number + 1)
    );
    ws.activate();
    sel.select();
    await ctx.sync();

    if (cols.check >= 0) {
      const ckRanges = chunks.map((ch) => {
        const ck = ws.getRangeByIndexes(ch.startRow, cols.check, ch.indices.length, 1);
        ck.load("values");
        return ck;
      });
      await ctx.sync();
      return ckRanges.flatMap((ck) => ck.values.map((r) => String(r[0] ?? "")));
    }
    return toWrite.map(() => "");
  });

  const rows: WrittenRowSummary[] = toWrite.map((r, i) => ({
    rowNum: targetRows[i] + 1,
    invoiceNumber: r.inv.invoice_number ?? "—",
    seller: r.inv.seller_name ?? "—",
    orderNo: r.orderNo || "（未匹配）",
    total: typeof r.total === "number" ? r.total : typeof r.alloc === "number" ? r.alloc : r.inv.total,
    checkStatus: checkStatuses[i] || "",
    note: r.inv.note ?? "",
  }));
  const address =
    chunks.length === 1
      ? `第 ${targetRows[0] + 1}${targetRows.length > 1 ? `～${targetRows[targetRows.length - 1] + 1}` : ""} 行`
      : `第 ${targetRows.map((r) => r + 1).join("、")} 行`;
  return { sheet: INVOICE_SHEET, address, rows, duplicates, created: false };
}

async function writeGenericInvoices(
  layout: GenericLayout,
  invoices: InvoiceRecord[],
  atts: Attachment[],
  tctx: ToolContext
): Promise<WriteResult> {
  let created = false;
  if (!layout.exists) {
    await executeTool("add_sheet", { name: INVOICE_SHEET }, tctx);
    created = true;
  }
  const { isFresh, headers, col0 } = layout;
  const col = (h: string) => headers.indexOf(h);

  const written: InvoiceRecord[] = [];
  const duplicates: InvoiceRecord[] = [];
  for (const inv of invoices) {
    const n = (inv.invoice_number ?? "").trim();
    if (n && layout.existingNumbers.has(n)) duplicates.push(inv);
    else {
      if (n) layout.existingNumbers.add(n);
      written.push(inv);
    }
  }

  const firstDataRow = isFresh ? 1 : layout.lastRow + 1;
  const serialStart = isFresh ? 1 : Math.max(1, layout.dataCount + 1);
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
      开票日期: dateToExcelSerial(inv.issue_date),
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

  if (!rows.length && !isFresh) return { sheet: INVOICE_SHEET, address: null, rows: [], duplicates, created };

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
      const c = L(h);
      if (c && dataRows) ws.getRange(`${c}${dataStart}:${c}${dataStart + dataRows - 1}`).numberFormat = [[fmt]];
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

  const summaries: WrittenRowSummary[] = written.map((inv, i) => ({
    rowNum: firstDataRow + i + 1,
    invoiceNumber: inv.invoice_number ?? "—",
    seller: inv.seller_name ?? "—",
    orderNo: "",
    total: inv.total,
    checkStatus: "",
    note: inv.note ?? "",
  }));
  return { sheet: INVOICE_SHEET, address, rows: summaries, duplicates, created };
}

export async function writeInvoices(
  layout: SheetLayout,
  invoices: InvoiceRecord[],
  atts: Attachment[],
  tctx: ToolContext
): Promise<WriteResult> {
  if (layout.mode === "project") return writeProjectInvoices(layout, invoices, atts, tctx);
  return writeGenericInvoices(layout, invoices, atts, tctx);
}

const fmtMoney = (v: number | null) =>
  typeof v === "number" ? v.toLocaleString("zh-CN", { minimumFractionDigits: 2 }) : "—";

export function summarize(res: WriteResult): string {
  const lines: string[] = [];
  if (res.rows.length) {
    lines.push(`已录入 **${res.rows.length}** 行到「${res.sheet}」${res.address ? `（${res.address}）` : ""}：`, "");
    const hasOrder = res.rows.some((r) => r.orderNo);
    if (hasOrder) {
      lines.push("| 行 | 发票号码 | 销售方 | 匹配订单号 | 金额 | 校验 |", "|---:|---|---|---|---:|---|");
      for (const r of res.rows)
        lines.push(`| ${r.rowNum} | ${r.invoiceNumber} | ${r.seller} | ${r.orderNo} | ${fmtMoney(r.total)} | ${r.checkStatus || r.note} |`);
    } else {
      lines.push("| 发票号码 | 销售方 | 价税合计 | 备注 |", "|---|---|---:|---|");
      for (const r of res.rows)
        lines.push(`| ${r.invoiceNumber} | ${r.seller} | ${fmtMoney(r.total)} | ${r.note} |`);
    }
    const warn = res.rows.filter((r) => r.note || (r.checkStatus && r.checkStatus !== "分配已核对"));
    if (warn.length) lines.push("", `⚠️ ${warn.length} 行需要人工核对（见“校验”和“备注”列）。`);
  } else {
    lines.push("没有新的发票需要录入。");
  }
  if (res.duplicates.length)
    lines.push("", `已跳过 ${res.duplicates.length} 张重复发票：${res.duplicates.map((d) => d.invoice_number).join("、")}`);
  return lines.join("\n");
}
