// Excel tools exposed to Gemini via function calling, plus undo snapshots.
import type { FunctionDeclaration } from "./gemini";

export type UndoOp = () => Promise<void>;

export interface ToolContext {
  undo: UndoOp[];
  createdSheets?: Set<string>;
}

export interface ChangedRange {
  sheet: string;
  address: string;
}

export interface ToolResult {
  response: Record<string, unknown>;
  changed?: ChangedRange[];
}

const MAX_READ_CELLS = 4000;
const MAX_WRITE_CELLS = 20000;
const MAX_SNAPSHOT_CELLS = 100000;
const MAX_PROPS_SNAPSHOT_CELLS = 10000;

/* ----------------------------- helpers ----------------------------- */

function splitAddress(sheet: string | undefined, address: string): { sheet?: string; address: string } {
  const m = address.trim().match(/^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/);
  if (m) return { sheet: m[1] ? m[1].replace(/''/g, "'") : m[2], address: m[3] };
  return { sheet: sheet || undefined, address: address.trim() };
}

function getSheet(ctx: Excel.RequestContext, name?: string): Excel.Worksheet {
  return name ? ctx.workbook.worksheets.getItem(name) : ctx.workbook.worksheets.getActiveWorksheet();
}

function stripSheet(address: string): string {
  const i = address.lastIndexOf("!");
  return i >= 0 ? address.slice(i + 1) : address;
}

export function colLetter(index: number): string {
  let s = "";
  let n = index + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function cellName(row: number, col: number): string {
  return `${colLetter(col)}${row + 1}`;
}

function str(v: unknown): string | undefined {
  return typeof v === "string" && v.trim() ? v.trim() : undefined;
}

function trimValue(v: unknown): unknown {
  if (typeof v === "string" && v.length > 300) return v.slice(0, 300) + "…";
  return v;
}

const PROPS_LOAD: Excel.CellPropertiesLoadOptions = {
  format: {
    fill: { color: true, pattern: true },
    font: { bold: true, italic: true, color: true, size: true, name: true, underline: true, strikethrough: true },
    horizontalAlignment: true,
    verticalAlignment: true,
    wrapText: true,
    borders: { color: true, style: true, weight: true },
  },
};

function stripNulls(o: unknown): unknown {
  if (Array.isArray(o)) return o.map(stripNulls);
  if (o && typeof o === "object") {
    const out: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(o)) {
      if (v === null || v === undefined) continue;
      const s = stripNulls(v);
      if (s && typeof s === "object" && !Array.isArray(s) && !Object.keys(s).length) continue;
      out[k] = s;
    }
    return out;
  }
  return o;
}

function sanitizeProps(cell: Excel.CellProperties): Excel.SettableCellProperties {
  const c = JSON.parse(JSON.stringify(cell)) as { format?: Record<string, any> };
  const f = c.format;
  if (f) {
    const fill = f.fill as { pattern?: string; color?: string } | undefined;
    if (fill) {
      // fill.pattern needs a newer API set; keep only a real solid color.
      if (fill.pattern === "None" || !fill.color || /^#?FFFFFF$/i.test(fill.color)) delete f.fill;
      else f.fill = { color: fill.color };
    }
    if (f.borders) {
      for (const k of Object.keys(f.borders)) {
        const b = f.borders[k];
        if (!b || b.style === "None") delete f.borders[k];
        else f.borders[k] = { style: b.style, color: b.color, weight: b.weight };
      }
    }
  }
  return stripNulls(c) as Excel.SettableCellProperties;
}

const ALL_EDGES = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"] as const;

/** Captures formulas, number formats and (for small ranges) cell formatting; returns an undo op. */
async function snapshotRange(sheetName: string | undefined, address: string): Promise<UndoOp | null> {
  return Excel.run(async (ctx) => {
    const ws = getSheet(ctx, sheetName);
    ws.load("name");
    const r = ws.getRange(address);
    r.load("cellCount");
    await ctx.sync();
    if (r.cellCount > MAX_SNAPSHOT_CELLS) return null;
    r.load(["formulas", "numberFormat"]);
    let propsResult: OfficeExtension.ClientResult<Excel.CellProperties[][]> | null = null;
    if (r.cellCount <= MAX_PROPS_SNAPSHOT_CELLS) propsResult = r.getCellProperties(PROPS_LOAD);
    try {
      await ctx.sync();
    } catch {
      // getCellProperties unsupported on this Excel build: retry without formatting.
      propsResult = null;
      r.load(["formulas", "numberFormat"]);
      await ctx.sync();
    }

    const sheet = ws.name;
    const formulas = r.formulas;
    const numberFormat = r.numberFormat;
    const props = propsResult ? propsResult.value.map((row) => row.map(sanitizeProps)) : null;

    return async () => {
      const errors: string[] = [];
      const step = async (label: string, fn: (rg: Excel.Range) => void) => {
        try {
          await Excel.run(async (ctx2) => {
            fn(ctx2.workbook.worksheets.getItem(sheet).getRange(address));
            await ctx2.sync();
          });
        } catch (e) {
          errors.push(`${label}: ${(e as Error).message}`);
        }
      };
      await step("取消合并", (rg) => rg.unmerge());
      await step("还原内容", (rg) => (rg.formulas = formulas));
      await step("还原数字格式", (rg) => (rg.numberFormat = numberFormat));
      if (props) {
        await step("还原格式", (rg) => {
          rg.format.fill.clear();
          ALL_EDGES.forEach((e) => (rg.format.borders.getItem(e).style = "None"));
          rg.format.font.bold = false;
          rg.format.font.italic = false;
          rg.setCellProperties(props);
        });
      }
      // Content restore is what matters; only fail if that failed.
      const fatal = errors.filter((e) => e.startsWith("还原内容"));
      if (fatal.length) throw new Error(`${sheet}!${address} ${fatal.join("; ")}`);
      if (errors.length) console.warn("Undo formatting partially failed", errors);
    };
  });
}

export async function pushSnapshot(tctx: ToolContext, sheet: string | undefined, address: string) {
  // Sheets created in this turn are simply deleted on undo; no need to snapshot them.
  if (sheet && tctx.createdSheets?.has(sheet)) return;
  const op = await snapshotRange(sheet, address);
  if (op) tctx.undo.push(op);
}

/* --------------------------- declarations -------------------------- */

const S = (description: string) => ({ type: "STRING", description });
const B = (description: string) => ({ type: "BOOLEAN", description });
const N = (description: string) => ({ type: "NUMBER", description });
const SHEET = S("工作表名称。省略则使用当前活动工作表。");

export const TOOL_DECLARATIONS: FunctionDeclaration[] = [
  {
    name: "get_workbook_info",
    description:
      "获取工作簿结构概览：所有工作表（名称、已用区域、行列数、前几行预览、表格、图表数量）、活动工作表、当前选区、定义的名称。开始任何任务前先调用它了解结构。",
    parameters: { type: "OBJECT", properties: {} },
  },
  {
    name: "read_range",
    description:
      "读取单元格区域的计算值（values 二维数组）。默认不返回公式源码以节省 token；仅当需要查看或修改公式时才传 include_formulas=true。大区域会被截断，可分段读取。",
    parameters: {
      type: "OBJECT",
      properties: {
        sheet: SHEET,
        address: S("A1 格式地址，例如 'A1:F50'，也可以是 'Sheet2!B2:D10'。"),
        include_formulas: B("是否同时返回公式源码，默认 false。仅排查/编辑公式时设为 true。"),
        include_number_format: B("是否同时返回数字格式，默认 false。"),
      },
      required: ["address"],
    },
  },
  {
    name: "get_selection",
    description: "获取用户当前选中的区域地址及其内容。",
    parameters: { type: "OBJECT", properties: {} },
  },
  {
    name: "write_range",
    description:
      "从 start_cell 开始写入一个二维数组（按行）。以 '=' 开头的字符串作为公式写入；数字直接写成 '123.45' 这样的字符串即可。公式必须使用英文函数名和英文逗号分隔参数（例如 =SUM(A1:A10)、=IF(A1>0,\"是\",\"否\")）。尽量使用公式引用而不是硬编码计算结果。返回写入地址、公式计算结果和错误单元格。",
    parameters: {
      type: "OBJECT",
      properties: {
        sheet: SHEET,
        start_cell: S("左上角起始单元格，例如 'B2'。"),
        values: {
          type: "ARRAY",
          description: "二维数组，外层为行，内层为列。",
          items: { type: "ARRAY", items: { type: "STRING" } },
        },
      },
      required: ["start_cell", "values"],
    },
  },
  {
    name: "clear_range",
    description: "清除区域的内容、格式或全部。",
    parameters: {
      type: "OBJECT",
      properties: {
        sheet: SHEET,
        address: S("A1 格式地址。"),
        what: { type: "STRING", enum: ["contents", "formats", "all"], description: "清除内容，默认 contents。" },
      },
      required: ["address"],
    },
  },
  {
    name: "format_range",
    description: "设置区域格式。仅传入需要修改的属性。",
    parameters: {
      type: "OBJECT",
      properties: {
        sheet: SHEET,
        address: S("A1 格式地址。"),
        bold: B("粗体"),
        italic: B("斜体"),
        font_color: S("字体颜色，如 '#FF0000'"),
        font_size: N("字号"),
        fill_color: S("填充色，如 '#FFF2CC'；传 'none' 清除填充"),
        number_format: S("数字格式，如 '#,##0.00'、'0.0%'、'yyyy-mm-dd'"),
        horizontal_alignment: { type: "STRING", enum: ["Left", "Center", "Right", "General"] },
        wrap_text: B("自动换行"),
        borders: { type: "STRING", enum: ["all", "outline", "none"], description: "边框" },
        merge: B("true 合并单元格，false 取消合并"),
        autofit_columns: B("自动调整列宽"),
        column_width: N("列宽（磅）"),
      },
      required: ["address"],
    },
  },
  {
    name: "insert_cells",
    description: "在指定地址插入空白单元格/整行/整列（例如 '5:7' 插入3行，'C:C' 插入一列）。",
    parameters: {
      type: "OBJECT",
      properties: {
        sheet: SHEET,
        address: S("A1 格式地址"),
        shift: { type: "STRING", enum: ["down", "right"], description: "现有单元格移动方向" },
      },
      required: ["address", "shift"],
    },
  },
  {
    name: "delete_cells",
    description: "删除单元格/整行/整列。",
    parameters: {
      type: "OBJECT",
      properties: {
        sheet: SHEET,
        address: S("A1 格式地址"),
        shift: { type: "STRING", enum: ["up", "left"], description: "剩余单元格移动方向" },
      },
      required: ["address", "shift"],
    },
  },
  {
    name: "add_sheet",
    description: "新建工作表。",
    parameters: {
      type: "OBJECT",
      properties: { name: S("新工作表名称"), activate: B("是否切换到该表，默认 true") },
      required: ["name"],
    },
  },
  {
    name: "rename_sheet",
    description: "重命名工作表。",
    parameters: {
      type: "OBJECT",
      properties: { sheet: S("原名称"), new_name: S("新名称") },
      required: ["sheet", "new_name"],
    },
  },
  {
    name: "create_table",
    description: "将区域转换为 Excel 表格（ListObject）。",
    parameters: {
      type: "OBJECT",
      properties: {
        sheet: SHEET,
        address: S("包含表头的数据区域"),
        has_headers: B("首行是否为表头，默认 true"),
        name: S("表格名称（可选，不能有空格）"),
        style: S("表格样式，如 'TableStyleMedium2'"),
      },
      required: ["address"],
    },
  },
  {
    name: "create_chart",
    description: "基于数据区域创建图表。",
    parameters: {
      type: "OBJECT",
      properties: {
        sheet: SHEET,
        data_address: S("数据区域（含表头）"),
        chart_type: {
          type: "STRING",
          description:
            "图表类型，如 ColumnClustered, ColumnStacked, BarClustered, Line, LineMarkers, Pie, Doughnut, Area, XYScatter",
        },
        title: S("图表标题"),
        anchor_cell: S("图表左上角放置的单元格，如 'H2'"),
        series_by: { type: "STRING", enum: ["auto", "columns", "rows"] },
      },
      required: ["data_address", "chart_type"],
    },
  },
  {
    name: "sort_range",
    description: "按某列排序区域。",
    parameters: {
      type: "OBJECT",
      properties: {
        sheet: SHEET,
        address: S("要排序的区域"),
        column_index: N("排序依据列，相对于区域首列的 0 基索引"),
        ascending: B("升序，默认 true"),
        has_headers: B("首行是否为表头，默认 true"),
      },
      required: ["address", "column_index"],
    },
  },
  {
    name: "find_in_workbook",
    description: "在所有工作表中搜索文本，返回匹配单元格地址。",
    parameters: {
      type: "OBJECT",
      properties: {
        text: S("要搜索的文本"),
        match_case: B("区分大小写"),
        complete_match: B("整个单元格完全匹配"),
      },
      required: ["text"],
    },
  },
  {
    name: "select_range",
    description: "在 Excel 中选中并跳转到某个区域，方便用户查看。",
    parameters: {
      type: "OBJECT",
      properties: { sheet: SHEET, address: S("A1 格式地址") },
      required: ["address"],
    },
  },
];

export const WRITE_TOOLS = new Set([
  "write_range",
  "clear_range",
  "format_range",
  "insert_cells",
  "delete_cells",
  "add_sheet",
  "rename_sheet",
  "create_table",
  "create_chart",
  "sort_range",
]);

export const TOOL_LABELS: Record<string, string> = {
  get_workbook_info: "查看工作簿结构",
  read_range: "读取区域",
  get_selection: "读取选区",
  write_range: "写入单元格",
  clear_range: "清除区域",
  format_range: "设置格式",
  insert_cells: "插入单元格",
  delete_cells: "删除单元格",
  add_sheet: "新建工作表",
  rename_sheet: "重命名工作表",
  create_table: "创建表格",
  create_chart: "创建图表",
  sort_range: "排序",
  find_in_workbook: "搜索",
  select_range: "跳转到区域",
};

/* ---------------------------- executors ---------------------------- */

type Args = Record<string, unknown>;

async function readRangeData(
  sheetName: string | undefined,
  address: string,
  includeFormat: boolean,
  includeFormulas: boolean
) {
  return Excel.run(async (ctx) => {
    const ws = getSheet(ctx, sheetName);
    let r = ws.getRange(address);
    r.load(["rowCount", "columnCount", "address"]);
    await ctx.sync();
    const fullAddress = r.address;
    const totalRows = r.rowCount;
    const cols = r.columnCount;
    let truncated: string | undefined;
    if (totalRows * cols > MAX_READ_CELLS) {
      const rowsAllowed = Math.max(1, Math.floor(MAX_READ_CELLS / cols));
      const colsAllowed = Math.min(cols, MAX_READ_CELLS);
      r = r.getCell(0, 0).getResizedRange(rowsAllowed - 1, colsAllowed - 1);
      truncated = `区域共 ${totalRows} 行 × ${cols} 列，仅返回前 ${rowsAllowed} 行${colsAllowed < cols ? `、前 ${colsAllowed} 列` : ""}。请分段读取剩余部分。`;
    }
    const props: string[] = ["values", "address", "rowIndex", "columnIndex"];
    if (includeFormulas) props.push("formulas");
    if (includeFormat) props.push("numberFormat");
    r.load(props);
    await ctx.sync();

    const out: Record<string, unknown> = {
      address: fullAddress,
      returned_address: r.address,
      values: r.values.map((row) => row.map(trimValue)),
    };
    if (includeFormulas) {
      const formulas: Record<string, string> = {};
      r.formulas.forEach((row, i) =>
        row.forEach((f, j) => {
          if (typeof f === "string" && f.startsWith("=")) formulas[cellName(r.rowIndex + i, r.columnIndex + j)] = f;
        })
      );
      if (Object.keys(formulas).length) out.formulas = formulas;
    }
    if (includeFormat) out.number_format = r.numberFormat;
    if (truncated) out.note = truncated;
    return out;
  });
}

async function getWorkbookInfo() {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name,items/visibility,items/position");
    const active = sheets.getActiveWorksheet();
    active.load("name");
    const sel = ctx.workbook.getSelectedRange();
    sel.load("address");
    const names = ctx.workbook.names;
    names.load("items/name,items/formula");
    await ctx.sync();

    const used = sheets.items.map((ws) => {
      const u = ws.getUsedRangeOrNullObject(true);
      u.load("address,rowCount,columnCount");
      ws.tables.load("items/name");
      ws.charts.load("items/name");
      return u;
    });
    await ctx.sync();

    const previews = used.map((u) => {
      if (u.isNullObject) return null;
      const p = u.getCell(0, 0).getResizedRange(Math.min(3, u.rowCount) - 1, Math.min(20, u.columnCount) - 1);
      p.load("values");
      return p;
    });
    await ctx.sync();

    return {
      active_sheet: active.name,
      selection: sel.address,
      sheets: sheets.items.map((ws, i) => {
        const u = used[i];
        return {
          name: ws.name,
          visibility: ws.visibility,
          used_range: u.isNullObject ? null : stripSheet(u.address),
          rows: u.isNullObject ? 0 : u.rowCount,
          columns: u.isNullObject ? 0 : u.columnCount,
          preview_first_rows: previews[i]?.values.map((row) => row.map(trimValue)) ?? [],
          tables: ws.tables.items.map((t) => t.name),
          charts: ws.charts.items.length,
        };
      }),
      defined_names: names.items.slice(0, 50).map((n) => ({ name: n.name, refers_to: n.formula })),
    };
  });
}

async function writeRange(a: Args, tctx: ToolContext): Promise<ToolResult> {
  const { sheet, address: start } = splitAddress(str(a.sheet), String(a.start_cell ?? "A1"));
  const raw = (a.values as unknown[][]) ?? [];
  if (!Array.isArray(raw) || !raw.length) throw new Error("values 不能为空");
  const cols = Math.max(...raw.map((r) => (Array.isArray(r) ? r.length : 1)));
  const matrix = raw.map((r) => {
    const row = (Array.isArray(r) ? r : [r]).map((v) => (v === null || v === undefined ? "" : String(v)));
    while (row.length < cols) row.push("");
    return row;
  });
  if (matrix.length * cols > MAX_WRITE_CELLS) throw new Error(`一次最多写入 ${MAX_WRITE_CELLS} 个单元格，请分批写入`);

  // Resolve target address first, then snapshot it for undo.
  const target = await Excel.run(async (ctx) => {
    const ws = getSheet(ctx, sheet);
    ws.load("name");
    const r = ws.getRange(start).getCell(0, 0).getResizedRange(matrix.length - 1, cols - 1);
    r.load("address");
    await ctx.sync();
    return { sheet: ws.name, address: stripSheet(r.address) };
  });
  await pushSnapshot(tctx, target.sheet, target.address);

  const result = await Excel.run(async (ctx) => {
    const r = ctx.workbook.worksheets.getItem(target.sheet).getRange(target.address);
    r.formulas = matrix;
    await ctx.sync();
    r.load(["values", "valueTypes", "rowIndex", "columnIndex"]);
    await ctx.sync();
    const errors: Record<string, unknown>[] = [];
    const computed: Record<string, unknown> = {};
    let n = 0;
    r.values.forEach((row, i) =>
      row.forEach((v, j) => {
        const name = cellName(r.rowIndex + i, r.columnIndex + j);
        const f = matrix[i][j];
        if (r.valueTypes[i][j] === "Error" && errors.length < 30) errors.push({ cell: name, value: v, formula: f });
        else if (f.startsWith("=") && n < 60) {
          computed[name] = trimValue(v);
          n++;
        }
      })
    );
    return { errors, computed };
  });

  return {
    response: {
      written: `${target.sheet}!${target.address}`,
      formula_results_sample: result.computed,
      ...(result.errors.length ? { errors: result.errors, hint: "存在错误值，请检查并修正公式。" } : {}),
    },
    changed: [target],
  };
}

async function resolveTarget(a: Args): Promise<ChangedRange> {
  const { sheet, address } = splitAddress(str(a.sheet), String(a.address ?? a.data_address ?? ""));
  if (!address) throw new Error("缺少 address");
  return Excel.run(async (ctx) => {
    const ws = getSheet(ctx, sheet);
    ws.load("name");
    const r = ws.getRange(address);
    r.load("address");
    await ctx.sync();
    return { sheet: ws.name, address: stripSheet(r.address) };
  });
}

const BORDER_EDGES_OUTLINE = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"] as const;
const BORDER_EDGES_ALL = [...BORDER_EDGES_OUTLINE, "InsideHorizontal", "InsideVertical"] as const;

async function formatRange(a: Args, tctx: ToolContext): Promise<ToolResult> {
  const t = await resolveTarget(a);
  await pushSnapshot(tctx, t.sheet, t.address);
  await Excel.run(async (ctx) => {
    const r = ctx.workbook.worksheets.getItem(t.sheet).getRange(t.address);
    const f = r.format;
    if (typeof a.bold === "boolean") f.font.bold = a.bold;
    if (typeof a.italic === "boolean") f.font.italic = a.italic;
    if (str(a.font_color)) f.font.color = str(a.font_color)!;
    if (typeof a.font_size === "number") f.font.size = a.font_size;
    if (str(a.fill_color)) {
      if (str(a.fill_color)!.toLowerCase() === "none") f.fill.clear();
      else f.fill.color = str(a.fill_color)!;
    }
    if (str(a.number_format)) r.numberFormat = [[str(a.number_format)!]];
    if (str(a.horizontal_alignment)) f.horizontalAlignment = str(a.horizontal_alignment) as Excel.HorizontalAlignment;
    if (typeof a.wrap_text === "boolean") f.wrapText = a.wrap_text;
    if (str(a.borders)) {
      const mode = str(a.borders);
      if (mode === "none") {
        BORDER_EDGES_ALL.forEach((e) => (f.borders.getItem(e).style = "None"));
      } else {
        (mode === "all" ? BORDER_EDGES_ALL : BORDER_EDGES_OUTLINE).forEach((e) => {
          const b = f.borders.getItem(e);
          b.style = "Continuous";
          b.color = "#000000";
        });
      }
    }
    if (a.merge === true) r.merge(false);
    if (a.merge === false) r.unmerge();
    if (typeof a.column_width === "number") f.columnWidth = a.column_width;
    if (a.autofit_columns) f.autofitColumns();
    await ctx.sync();
  });
  if (a.merge === true) {
    tctx.undo.push(async () =>
      Excel.run(async (ctx) => {
        ctx.workbook.worksheets.getItem(t.sheet).getRange(t.address).unmerge();
        await ctx.sync();
      })
    );
  }
  return { response: { formatted: `${t.sheet}!${t.address}` }, changed: [t] };
}

async function clearRange(a: Args, tctx: ToolContext): Promise<ToolResult> {
  const t = await resolveTarget(a);
  await pushSnapshot(tctx, t.sheet, t.address);
  const what = str(a.what) ?? "contents";
  const applyTo = what === "all" ? "All" : what === "formats" ? "Formats" : "Contents";
  await Excel.run(async (ctx) => {
    ctx.workbook.worksheets.getItem(t.sheet).getRange(t.address).clear(applyTo as Excel.ClearApplyTo);
    await ctx.sync();
  });
  return { response: { cleared: `${t.sheet}!${t.address}`, what }, changed: [t] };
}

async function insertCells(a: Args, tctx: ToolContext): Promise<ToolResult> {
  const t = await resolveTarget(a);
  const down = str(a.shift) !== "right";
  await Excel.run(async (ctx) => {
    ctx.workbook.worksheets.getItem(t.sheet).getRange(t.address).insert(down ? "Down" : "Right");
    await ctx.sync();
  });
  tctx.undo.push(async () =>
    Excel.run(async (ctx) => {
      ctx.workbook.worksheets.getItem(t.sheet).getRange(t.address).delete(down ? "Up" : "Left");
      await ctx.sync();
    })
  );
  return { response: { inserted: `${t.sheet}!${t.address}`, shift: down ? "down" : "right" }, changed: [t] };
}

async function deleteCells(a: Args, tctx: ToolContext): Promise<ToolResult> {
  const t = await resolveTarget(a);
  const up = str(a.shift) !== "left";
  const restore = await snapshotRange(t.sheet, t.address);
  await Excel.run(async (ctx) => {
    ctx.workbook.worksheets.getItem(t.sheet).getRange(t.address).delete(up ? "Up" : "Left");
    await ctx.sync();
  });
  tctx.undo.push(async () => {
    await Excel.run(async (ctx) => {
      ctx.workbook.worksheets.getItem(t.sheet).getRange(t.address).insert(up ? "Down" : "Right");
      await ctx.sync();
    });
    if (restore) await restore();
  });
  return { response: { deleted: `${t.sheet}!${t.address}` } };
}

async function addSheet(a: Args, tctx: ToolContext): Promise<ToolResult> {
  const name = str(a.name);
  if (!name) throw new Error("缺少 name");
  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.add(name);
    if (a.activate !== false) ws.activate();
    await ctx.sync();
  });
  tctx.createdSheets ??= new Set();
  tctx.createdSheets.add(name);
  tctx.undo.push(async () =>
    Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(name);
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name,items/visibility");
      await ctx.sync();
      const other = sheets.items.find((s) => s.name !== name && s.visibility === "Visible");
      if (other) other.activate();
      ws.delete();
      await ctx.sync();
    })
  );
  return { response: { added_sheet: name } };
}

async function renameSheet(a: Args, tctx: ToolContext): Promise<ToolResult> {
  const oldName = str(a.sheet)!;
  const newName = str(a.new_name)!;
  await Excel.run(async (ctx) => {
    ctx.workbook.worksheets.getItem(oldName).name = newName;
    await ctx.sync();
  });
  tctx.undo.push(async () =>
    Excel.run(async (ctx) => {
      ctx.workbook.worksheets.getItem(newName).name = oldName;
      await ctx.sync();
    })
  );
  return { response: { renamed: { from: oldName, to: newName } } };
}

async function createTable(a: Args, tctx: ToolContext): Promise<ToolResult> {
  const t = await resolveTarget(a);
  await pushSnapshot(tctx, t.sheet, t.address);
  const tableName = await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(t.sheet);
    const table = ws.tables.add(t.address, a.has_headers !== false);
    if (str(a.name)) table.name = str(a.name)!;
    if (str(a.style)) table.style = str(a.style)!;
    table.load("name");
    await ctx.sync();
    return table.name;
  });
  // Undo order is reversed: convert table back to a range first, then the snapshot restores formatting.
  tctx.undo.push(async () =>
    Excel.run(async (ctx) => {
      ctx.workbook.tables.getItem(tableName).convertToRange();
      await ctx.sync();
    })
  );
  return { response: { table: tableName, range: `${t.sheet}!${t.address}` }, changed: [t] };
}

async function createChart(a: Args, tctx: ToolContext): Promise<ToolResult> {
  const t = await resolveTarget(a);
  const chartName = `GeminiChart_${Date.now()}`;
  const seriesBy = str(a.series_by) === "rows" ? "Rows" : str(a.series_by) === "columns" ? "Columns" : "Auto";
  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(t.sheet);
    const chart = ws.charts.add(
      (str(a.chart_type) ?? "ColumnClustered") as Excel.ChartType,
      ws.getRange(t.address),
      seriesBy as Excel.ChartSeriesBy
    );
    chart.name = chartName;
    if (str(a.title)) chart.title.text = str(a.title)!;
    if (str(a.anchor_cell)) chart.setPosition(str(a.anchor_cell)!);
    await ctx.sync();
  });
  tctx.undo.push(async () =>
    Excel.run(async (ctx) => {
      ctx.workbook.worksheets.getItem(t.sheet).charts.getItem(chartName).delete();
      await ctx.sync();
    })
  );
  return { response: { chart: chartName, data: `${t.sheet}!${t.address}` } };
}

async function sortRange(a: Args, tctx: ToolContext): Promise<ToolResult> {
  const t = await resolveTarget(a);
  await pushSnapshot(tctx, t.sheet, t.address);
  await Excel.run(async (ctx) => {
    ctx.workbook.worksheets
      .getItem(t.sheet)
      .getRange(t.address)
      .sort.apply([{ key: Number(a.column_index ?? 0), ascending: a.ascending !== false }], false, a.has_headers !== false);
    await ctx.sync();
  });
  return { response: { sorted: `${t.sheet}!${t.address}` }, changed: [t] };
}

async function findInWorkbook(a: Args): Promise<ToolResult> {
  const text = String(a.text ?? "");
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    const found = sheets.items.map((ws) => {
      const f = ws.findAllOrNullObject(text, {
        completeMatch: a.complete_match === true,
        matchCase: a.match_case === true,
      });
      f.load("address,cellCount");
      return f;
    });
    await ctx.sync();
    const matches: { sheet: string; cells: string; count: number }[] = [];
    found.forEach((f, i) => {
      if (!f.isNullObject) {
        const addr = f.address.split(",").map(stripSheet).join(",");
        matches.push({
          sheet: sheets.items[i].name,
          count: f.cellCount,
          cells: addr.length > 1500 ? addr.slice(0, 1500) + "…" : addr,
        });
      }
    });
    return { response: { text, matches } };
  });
}

async function selectRange(a: Args): Promise<ToolResult> {
  const t = await resolveTarget(a);
  await selectAddress(t.sheet, t.address);
  return { response: { selected: `${t.sheet}!${t.address}` } };
}

export async function selectAddress(sheet: string, address: string) {
  await Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheet);
    ws.activate();
    ws.getRange(address).select();
    await ctx.sync();
  });
}

export async function getSelectionInfo(maxCells = 200) {
  return Excel.run(async (ctx) => {
    const r = ctx.workbook.getSelectedRange();
    r.load("address,cellCount");
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    ws.load("name");
    await ctx.sync();
    const out: Record<string, unknown> = { active_sheet: ws.name, selection: r.address, cell_count: r.cellCount };
    if (r.cellCount <= maxCells) {
      r.load("values,formulas");
      await ctx.sync();
      out.values = r.values.map((row) => row.map(trimValue));
      if (r.formulas.some((row) => row.some((f) => typeof f === "string" && f.startsWith("="))))
        out.formulas = r.formulas;
    }
    return out;
  });
}

export async function executeTool(name: string, args: Args, tctx: ToolContext): Promise<ToolResult> {
  switch (name) {
    case "get_workbook_info":
      return { response: await getWorkbookInfo() };
    case "read_range": {
      const { sheet, address } = splitAddress(str(args.sheet), String(args.address ?? ""));
      return {
        response: await readRangeData(sheet, address, args.include_number_format === true, args.include_formulas === true),
      };
    }
    case "get_selection":
      return { response: await getSelectionInfo(MAX_READ_CELLS) };
    case "write_range":
      return writeRange(args, tctx);
    case "clear_range":
      return clearRange(args, tctx);
    case "format_range":
      return formatRange(args, tctx);
    case "insert_cells":
      return insertCells(args, tctx);
    case "delete_cells":
      return deleteCells(args, tctx);
    case "add_sheet":
      return addSheet(args, tctx);
    case "rename_sheet":
      return renameSheet(args, tctx);
    case "create_table":
      return createTable(args, tctx);
    case "create_chart":
      return createChart(args, tctx);
    case "sort_range":
      return sortRange(args, tctx);
    case "find_in_workbook":
      return findInWorkbook(args);
    case "select_range":
      return selectRange(args);
    default:
      throw new Error(`未知工具: ${name}`);
  }
}
