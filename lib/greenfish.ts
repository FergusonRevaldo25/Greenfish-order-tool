import { Row, Destination, destinationFor } from "./address";
import { openBook, readCells, lastGridRow, rotatedStyles, colNum } from "./xl";

export type { Row, Destination };
export type Item = { title: string; qty: number };
export type Order = { order: string; dest: Destination; items: Item[] };
export type Template = {
  buf: ArrayBuffer;
  /** column letter -> header text, from row 3 of Sheet1 */
  headers: Record<string, string>;
  /** normalised Shopify title -> column header (or IGNORE), from the template's own `mapping` sheet */
  mapping: Map<string, string>;
  /** first and last order row on the sheet (4 to 32 on the standard Greenfish sheet) */
  firstRow: number;
  lastRow: number;
};
export type Placed = { order: Order; cells: Record<string, number>; unmapped: Item[] };

export const TEMPLATE_FIRST_ROW = 4;
export const rowsPerPage = (t: Template) => t.lastRow - t.firstRow + 1;

export const norm = (s: string) => s.toLowerCase().replace(/\s+/g, " ").replace(/\s*\|\s*/g, " | ").trim();
/** "#GF60207", "GF60207" and " gf60207 " are the same order. */
export const orderKey = (s: string) => String(s).replace(/^#/, "").trim().toUpperCase();

const FIXED_HEADERS = ["DELIVERED", "ORDER#", "AREA"];
export const isProductHeader = (h: string) => !FIXED_HEADERS.includes(h.trim().toUpperCase());

/** Raised for a problem the person can fix (wrong file, damaged template); the message is shown as is. */
export class FriendlyError extends Error {}

/* ---------- reading the Shopify export (file parsing lives in lib/exportFile.ts) ---------- */

/** Shopify exports one row per line item; only the first row of an order carries the shipping details. */
export function groupOrders(rows: Row[]): Order[] {
  const orders = new Map<string, Order>();
  for (const r of rows) {
    const id = (r["Name"] || "").trim();
    if (!id) continue;
    let o = orders.get(id);
    if (!o) orders.set(id, (o = { order: id, dest: { kind: "unknown", label: "" }, items: [] }));
    if (o.dest.kind === "unknown") o.dest = destinationFor(r);
    const title = (r["Lineitem name"] || "").trim();
    const qty = parseFloat(r["Lineitem quantity"] || "0");
    if (title && qty) o.items.push({ title, qty });
  }
  return [...orders.values()];
}

/* ---------- reading a template, or a sheet that was already filled in ---------- */
export type Loaded = { template: Template; cells: Map<string, string> };

export async function loadSheetFile(buf: ArrayBuffer): Promise<Loaded> {
  let book;
  try { book = await openBook(buf); }
  catch { throw new FriendlyError("That file couldn't be opened. Please use an Excel (.xlsx) file."); }
  if (!book.sheets.length) throw new FriendlyError("That Excel file has no sheets in it.");
  const s1 = book.sheets[0];
  const cells = await readCells(book, s1.path);

  const headers: Record<string, string> = {};
  for (const [ref, v] of cells) {
    const m = /^([A-Z]+)3$/.exec(ref);
    if (m) headers[m[1]] = v.trim();
  }
  if (!/order/i.test(headers["B"] || "") || !/area/i.test(headers["C"] || ""))
    throw new FriendlyError("This doesn't look like a Greenfish daily sheet: row 3 of the first sheet should have ORDER# in column B and AREA in column C.");
  if (!Object.values(headers).some(isProductHeader))
    throw new FriendlyError("I couldn't find any product column headings in row 3 of the first sheet.");

  const xml = await book.zip.file(s1.path)!.async("string");
  const lastRow = lastGridRow(xml, TEMPLATE_FIRST_ROW);
  if (lastRow < TEMPLATE_FIRST_ROW)
    throw new FriendlyError("The order rows (from row 4 down) are missing in the first sheet. Keep the empty formatted rows in the template.");

  const mapping = new Map<string, string>();
  const ms = book.sheets.find(s => s.name.toLowerCase() === "mapping");
  if (ms) {
    const mc = await readCells(book, ms.path);
    for (const [ref, v] of mc) {
      const m = /^A(\d+)$/.exec(ref);
      if (!m || +m[1] < 2) continue;
      const target = mc.get(`B${m[1]}`);
      if (target) mapping.set(norm(v), target.trim());
    }
  }
  return { template: { buf, headers, mapping, firstRow: TEMPLATE_FIRST_ROW, lastRow }, cells };
}

export async function readTemplate(buf: ArrayBuffer): Promise<Template> {
  return (await loadSheetFile(buf)).template;
}

/** A short description of a template for the screen, with anything that will cause trouble later. */
export function describeTemplate(t: Template) {
  const cols = new Set(Object.values(t.headers).filter(isProductHeader).map(h => h.toUpperCase()));
  const warnings: string[] = [];
  if (!t.mapping.size) warnings.push("There's no 'mapping' sheet (or it is empty), so every product will have to be matched by hand.");
  let broken = 0;
  for (const target of t.mapping.values()) if (target.toUpperCase() !== "IGNORE" && !cols.has(target.toUpperCase())) broken++;
  if (broken) warnings.push(`${broken} ${broken === 1 ? "product in the mapping sheet points" : "products in the mapping sheet point"} to a column that isn't on the sheet, so they will be flagged for matching.`);
  return { columns: cols.size, mapped: t.mapping.size, rows: rowsPerPage(t), warnings };
}

export const headersKey = (t: Template) =>
  Object.entries(t.headers).sort((a, b) => colNum(a[0]) - colNum(b[0])).map(([c, h]) => `${c}:${h.toUpperCase()}`).join("|");

/** A sheet someone has already filled in (earlier today, say): which orders it holds and where the next one goes. */
export type Existing = {
  name: string;
  template: Template;
  orders: string[];
  nextRow: number; // first empty order row; lastRow + 1 when the sheet is full
  pageNo: number | null;
  dateText: string;
};

const excelDate = (v: string) => {
  if (!/^\d+(\.\d+)?$/.test(v) || +v < 30000) return v;
  const d = new Date(Math.round((+v - 25569) * 86400000));
  const p = (n: number) => String(n).padStart(2, "0");
  return `${p(d.getUTCDate())}/${p(d.getUTCMonth() + 1)}/${d.getUTCFullYear()}`;
};

export async function readExisting(name: string, buf: ArrayBuffer): Promise<Existing> {
  const { template, cells } = await loadSheetFile(buf);
  const orders: string[] = [];
  let nextRow = template.firstRow;
  for (let r = template.firstRow; r <= template.lastRow; r++) {
    const v = cells.get(`B${r}`);
    if (v) { orders.push(orderKey(v)); nextRow = r + 1; }
  }
  const page = /PAGE:?\s*(\d+)/i.exec(cells.get("G1") || "");
  return { name, template, orders, nextRow, pageNo: page ? +page[1] : null, dateText: excelDate(cells.get("C1") || "") };
}

/* ---------- title -> column ---------- */
export function lookup(title: string, t: Template, extra: Record<string, string> = {}): string | undefined {
  const n = norm(title);
  if (extra[n]) return extra[n];
  if (t.mapping.has(n)) return t.mapping.get(n);
  // Shopify appends " - Variant" to the product title: fall back to the longest mapped title that is a prefix.
  let best: string | undefined, bestLen = 0;
  for (const [k, v] of t.mapping) if (n.startsWith(k + " - ") && k.length > bestLen) { best = v; bestLen = k.length; }
  return best;
}

/** The sheet has two NOR SALMON columns: the first sits in the fresh block, the second in the frozen block. */
function columnFor(header: string, title: string, t: Template): string | undefined {
  const cols = Object.entries(t.headers)
    .filter(([, h]) => h.toUpperCase() === header.toUpperCase())
    .sort((a, b) => colNum(a[0]) - colNum(b[0]))
    .map(([c]) => c);
  if (!cols.length) return undefined;
  if (cols.length === 1) return cols[0];
  return /fresh/i.test(title) ? cols[0] : cols[cols.length - 1];
}

export function place(orders: Order[], t: Template, extra: Record<string, string> = {}, skipEmpty = true): Placed[] {
  const out: Placed[] = [];
  for (const order of orders) {
    const cells: Record<string, number> = {};
    const unmapped: Item[] = [];
    for (const it of order.items) {
      const h = lookup(it.title, t, extra);
      if (h === undefined) { unmapped.push(it); continue; }
      if (h.toUpperCase() === "IGNORE") continue;
      const col = columnFor(h, it.title, t);
      if (!col) { unmapped.push(it); continue; }
      cells[col] = (cells[col] || 0) + it.qty; // sizes of the same product add up
    }
    if (skipEmpty && !Object.keys(cells).length && !unmapped.length) continue;
    out.push({ order, cells, unmapped });
  }
  return out;
}

/* ---------- planning: which order goes on which row of which sheet ---------- */
export type SheetPlan = {
  /** "existing": a sheet that was uploaded and gets more rows. "new": a fresh page. */
  source: "existing" | "new";
  template: Template;
  startRow: number;
  items: Placed[];
  pageNo: number;
  pageCount: number;
  existingName?: string;
  dateText: string; // date already printed on the sheet, if any
};
export type Plan = { placed: Placed[]; alreadyThere: number; sheets: SheetPlan[]; template: Template };

export const paginate = <T,>(a: T[], size: number) => {
  const pages: T[][] = [];
  for (let i = 0; i < a.length; i += size) pages.push(a.slice(i, i + size));
  return pages;
};

/**
 * Orders already on an uploaded sheet are skipped. New ones fill the empty rows of the uploaded sheets first
 * (in page order), then go onto fresh pages made from the same sheet.
 */
export function buildPlan(orders: Order[], base: Template, existing: Existing[], extra: Record<string, string>): Plan {
  const sorted = [...existing].sort((a, b) => (a.pageNo ?? 1e9) - (b.pageNo ?? 1e9));
  const taken = new Set(sorted.flatMap(e => e.orders));
  const fresh = orders.filter(o => !taken.has(orderKey(o.order)));
  const template = sorted[0]?.template ?? base;
  const placed = place(fresh, template, extra);

  const queue = [...placed];
  const sheets: Omit<SheetPlan, "pageNo" | "pageCount">[] = [];
  for (const e of sorted) {
    const room = Math.max(0, e.template.lastRow - e.nextRow + 1);
    sheets.push({ source: "existing", template: e.template, startRow: e.nextRow, items: queue.splice(0, room), existingName: e.name, dateText: e.dateText });
  }
  for (const chunk of paginate(queue, rowsPerPage(template)))
    sheets.push({ source: "new", template, startRow: template.firstRow, items: chunk, dateText: sorted[0]?.dateText ?? "" });

  return {
    placed,
    alreadyThere: orders.length - fresh.length,
    template,
    sheets: sheets.map((s, i) => ({ ...s, pageNo: i + 1, pageCount: sheets.length })),
  };
}

/* ---------- writing into the template without touching anything else ---------- */
const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const cellRe = (ref: string) => new RegExp(`<c r="${ref}"((?:\\s+[A-Za-z0-9:_]+="[^"]*")*)\\s*(?:/>|>[\\s\\S]*?</c>)`);
const hasCell = (xml: string, ref: string) => cellRe(ref).test(xml);

export function setCell(xml: string, ref: string, value: string | number): string {
  const m = cellRe(ref).exec(xml);
  if (!m) throw new FriendlyError(`The sheet has no cell ${ref} to write into. Was the sheet changed (rows deleted)?`);
  const attrs = m[1].replace(/\s+t="[^"]*"/, "");
  const next = typeof value === "number"
    ? `<c r="${ref}"${attrs}><v>${value}</v></c>`
    : `<c r="${ref}"${attrs} t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`;
  return xml.replace(cellRe(ref), () => next);
}

/** A clickable cell: Excel's HYPERLINK formula, so no extra relationship parts are added to the file. */
export function setLink(xml: string, ref: string, text: string, url: string): string {
  const m = cellRe(ref).exec(xml);
  if (!m) throw new FriendlyError(`The sheet has no cell ${ref} to write into. Was the sheet changed (rows deleted)?`);
  const attrs = m[1].replace(/\s+t="[^"]*"/, "");
  const q = (v: string) => esc(v.replace(/"/g, '""'));
  const next = `<c r="${ref}"${attrs} t="str"><f>HYPERLINK("${q(url)}","${q(text)}")</f><v>${esc(text)}</v></c>`;
  return xml.replace(cellRe(ref), () => next);
}

/** Empties every order row but keeps the formatting, so an uploaded sheet can become a blank next page. */
export function clearRows(xml: string, first: number, last: number): string {
  return xml.replace(/<c r="([A-Z]+)(\d+)"((?:\s+[A-Za-z0-9:_]+="[^"]*")*)\s*(?:\/>|>[\s\S]*?<\/c>)/g, (all, col, row, attrs) => {
    if (+row < first || +row > last) return all;
    const s = /\ss="\d+"/.exec(attrs)?.[0] ?? "";
    return `<c r="${col}${row}"${s}/>`;
  });
}

/**
 * The standard template has one stray row (11) whose cells are rotated sideways in a heavy font, which made the
 * address on that row unreadable. Any order-row cell with rotated text takes the look of the same column in the
 * first order row; the row height follows.
 */
export function tidyRows(xml: string, rotated: Set<number>, first: number, last: number): string {
  if (!rotated.size) return xml;
  const style = new Map<string, string>();
  for (const m of xml.matchAll(new RegExp(`<c r="([A-Z]+)${first}"((?:\\s+[A-Za-z0-9:_]+="[^"]*")*)`, "g")))
    style.set(m[1], /\ss="(\d+)"/.exec(m[2])?.[1] ?? "");
  const rowAttrs = new RegExp(`<row r="${first}"([^>]*)>`).exec(xml)?.[1] ?? "";
  const refHt = /\sht="([\d.]+)"/.exec(rowAttrs)?.[1];
  const fixedRows = new Set<number>();
  let out = xml.replace(/<c r="([A-Z]+)(\d+)"((?:\s+[A-Za-z0-9:_]+="[^"]*")*)/g, (all, col, row, attrs) => {
    if (+row <= first || +row > last) return all;
    const cur = /\ss="(\d+)"/.exec(attrs)?.[1];
    const want = style.get(col);
    if (cur === undefined || want === undefined || !rotated.has(+cur) || rotated.has(+want)) return all;
    fixedRows.add(+row);
    return `<c r="${col}${row}"${attrs.replace(/\ss="\d+"/, ` s="${want}"`)}`;
  });
  if (refHt) for (const r of fixedRows)
    out = out.replace(new RegExp(`(<row r="${r}"[^>]*?)\\sht="[\\d.]+"`), `$1 ht="${refHt}"`);
  return out;
}

/** Order numbers like GF60207 were cut off in the 8.5-wide column B, and addresses need room too. Only ever widens. */
export function widenColumns(xml: string, min: Record<number, number>): string {
  return xml.replace(/<col\b[^>]*\/>/g, tag => {
    const lo = /\smin="(\d+)"/.exec(tag)?.[1], hi = /\smax="(\d+)"/.exec(tag)?.[1], w = /\swidth="([\d.]+)"/.exec(tag)?.[1];
    if (!lo || lo !== hi || !w || min[+lo] === undefined || +w >= min[+lo]) return tag;
    return tag.replace(/\swidth="[\d.]+"/, ` width="${min[+lo]}"`);
  });
}
export const MIN_COLUMN_WIDTHS = { 2: 11, 3: 26 };

export async function fillSheet(s: SheetPlan, date: string): Promise<Uint8Array> {
  const t = s.template;
  const book = await openBook(t.buf);
  const zip = book.zip;
  const path = book.sheets[0].path;
  let xml = await zip.file(path)!.async("string");

  // A fresh page made from an uploaded sheet must start empty.
  if (s.source === "new") xml = clearRows(xml, t.firstRow, t.lastRow);
  const styles = zip.file("xl/styles.xml");
  if (styles) xml = tidyRows(xml, rotatedStyles(await styles.async("string")), t.firstRow, t.lastRow);
  xml = widenColumns(xml, MIN_COLUMN_WIDTHS);

  if (!s.dateText && hasCell(xml, "C1")) xml = setCell(xml, "C1", date);
  if (hasCell(xml, "G1")) xml = setCell(xml, "G1", `PAGE:  ${s.pageNo}  OF  ${s.pageCount}`);
  s.items.forEach((p, i) => {
    const r = s.startRow + i;
    const n = p.order.order.replace(/^#/, "");
    xml = setCell(xml, `B${r}`, /^\d+$/.test(n) ? Number(n) : n);
    const d = p.order.dest;
    if (d.kind === "delivery") xml = setLink(xml, `C${r}`, d.label, d.url);
    else if (d.kind === "collection") xml = setCell(xml, `C${r}`, d.label);
    for (const [col, q] of Object.entries(p.cells)) xml = setCell(xml, `${col}${r}`, q);
  });
  zip.file(path, xml);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}
