import JSZip from "jszip";
import * as XLSX from "xlsx";
import Papa from "papaparse";

export type Row = Record<string, string>;
export type Item = { title: string; qty: number };
export type Order = {
  order: string;
  area: string;
  items: Item[];
  collection: boolean;
  collectionReason: "" | "method" | "no-address";
  method: string;
  street: string;
  suburb: string;
  city: string;
  province: string;
  zip: string;
  customer: string;
  mapsUrl: string;
};
export type AreaMode = "zip" | "suburb" | "city" | "street";
export type Template = {
  buf: ArrayBuffer;
  /** column letter -> header text, from row 3 of Sheet1 */
  headers: Record<string, string>;
  /** normalised Shopify title -> column header (or IGNORE), from the template's own `mapping` sheet */
  mapping: Map<string, string>;
  /** normalised product family (title before the first "|") -> header, only where every mapped variant agrees */
  families: Map<string, string>;
};
/** One product line of an order, and where it landed. */
export type Line = { title: string; qty: number; entree: boolean; header?: string; ignored?: boolean; /** product with no column of its own: goes in a spare column */ spare?: string; how: "template" | "family" | "saved" | "none" };
/** `spare` lines go in the empty product columns H to N; `unlisted` is products found on no list at all. */
export type Placed = { order: Order; cells: Record<string, number | string>; lines: Line[]; spare: Line[]; unlisted: Line[] };
/** One upload. Once added, a batch is never changed by later uploads. */
export type Batch = { id: string; at: string; fileName: string; orders: Order[]; skipped: number };

export const FIRST_ROW = 4;
export const LAST_ROW = 32; // 29 order rows per page
export const ROWS_PER_PAGE = LAST_ROW - FIRST_ROW + 1;
/** The blank, rotated header columns in the template where products without a column of their own are added. */
export const SPARE_COLUMNS = ["H", "I", "J", "K", "L", "M", "N"];
/** If a page needs more spare columns than that, the rest is written as text here so no order is ever dropped. */
export const OTHER_COLUMN = "AS";
/** Not seafood or food: gift cards, merchandise, fees and tools are never added to the sheet. */
const NON_PRODUCT = /gift card|hoodie|greenfish cap|service fee|turtle programme|test product|courier|braai tong|oyster knife|chopsticks|sushi mat|sauce dish/i;

/** "Yellowtail Fillets | Fresh Fish Box | Wild caught" -> "FRESH YELLOWTAIL FILLETS" */
export function spareLabel(title: string): string {
  const [first, ...rest] = title.split("|").map(x => x.trim());
  let base = first.replace(/\([^)]*\)/g, "").replace(/entr[eé]e/gi, "").replace(/\s+/g, " ").trim();
  const state = /fresh/i.test(base) ? "" : /fresh/i.test(rest.join(" ")) ? "FRESH " : /frozen/i.test(base) ? "" : /frozen/i.test(rest.join(" ")) ? "FROZEN " : "";
  const label = (state + base).toUpperCase();
  return label.length > 30 ? label.slice(0, 29).trim() + "." : label;
}

/** lower-case, no accents, straight quotes, single spaces, tidy pipes */
export const norm = (s: string) =>
  s.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[\u2018\u2019\u00b4`]/g, "'")
    .toLowerCase().replace(/\s+/g, " ").replace(/\s*\|\s*/g, " | ").trim();
const isEntree = (s: string) => /entr[eé]e/i.test(s.normalize("NFC"));
const familyOf = (s: string) => norm(s).split(" | ")[0].replace(/\bentree\b/g, "").replace(/\s+/g, " ").trim();

/* ---------- reading the Shopify export ---------- */
export async function readExport(file: File): Promise<Row[]> {
  if (/\.csv$/i.test(file.name)) return Papa.parse<Row>(await file.text(), { header: true, skipEmptyLines: true }).data;
  const wb = XLSX.read(await file.arrayBuffer());
  return XLSX.utils.sheet_to_json<Row>(wb.Sheets[wb.SheetNames[0]], { defval: "", raw: false });
}

const COLLECT = /collect|pick.?up|in.?store/i;

/** Shopify exports one row per line item; only the first row of an order carries the shipping details. */
export function groupOrders(rows: Row[]): Order[] {
  const orders = new Map<string, Order>();
  const g = (r: Row, k: string) => (r[k] || "").trim();
  for (const r of rows) {
    const id = g(r, "Name");
    if (!id) continue;
    let o = orders.get(id);
    if (!o) {
      o = { order: id, area: "", items: [], collection: false, collectionReason: "", method: "", street: "", suburb: "", city: "", province: "", zip: "", customer: "", mapsUrl: "" };
      orders.set(id, o);
    }
    o.method ||= g(r, "Shipping Method");
    o.street ||= g(r, "Shipping Address1");
    o.suburb ||= g(r, "Shipping Address2");
    o.city ||= g(r, "Shipping City");
    o.province ||= g(r, "Shipping Province Name") || g(r, "Shipping Province");
    o.zip ||= g(r, "Shipping Zip").replace(/^'/, "");
    o.customer ||= g(r, "Shipping Name") || g(r, "Billing Name");
    const title = g(r, "Lineitem name");
    const qty = parseFloat(r["Lineitem quantity"] || "0");
    if (title && qty) o.items.push({ title, qty });
  }
  for (const o of orders.values()) {
    // A collection is any order whose shipping method says so, or that has no street address at all.
    o.collectionReason = COLLECT.test(o.method) ? "method" : !o.street ? "no-address" : "";
    o.collection = o.collectionReason !== "";
    const parts = [o.street, o.suburb, o.city, o.province, o.zip, "South Africa"]
      .filter(Boolean).filter((x, i, a) => i === 0 || x.toLowerCase() !== a[i - 1].toLowerCase());
    // The full address only feeds the map link. It is never shown or written to the sheet.
    o.mapsUrl = o.collection ? "" : "https://www.google.com/maps/search/?api=1&query=" + encodeURIComponent(parts.join(", "));
  }
  return [...orders.values()];
}

export function areaText(o: Order, mode: AreaMode): string {
  if (o.collection) return "COLLECTION";
  if (mode === "street") return o.street || o.suburb || o.city;
  if (mode === "city") return o.city;
  if (mode === "suburb") return o.suburb || o.city;
  return o.zip || o.suburb || o.city; // postal code
}

/** Adds a new upload to the day. Orders already on the sheet are left exactly as they were. */
export function mergeBatch(batches: Batch[], incoming: Order[], fileName: string, at = new Date().toISOString()) {
  const have = new Set(batches.flatMap(b => b.orders.map(o => o.order)));
  const fresh = incoming.filter(o => !have.has(o.order));
  const skipped = incoming.length - fresh.length;
  const next = fresh.length ? [...batches, { id: at + fileName, at, fileName, orders: fresh, skipped }] : batches;
  return { batches: next, added: fresh.length, skipped };
}

/* ---------- reading the template ---------- */
export function readTemplate(buf: ArrayBuffer): Template {
  const wb = XLSX.read(buf);
  const s1 = wb.Sheets[wb.SheetNames[0]];
  const headers: Record<string, string> = {};
  for (const key of Object.keys(s1)) {
    const m = /^([A-Z]+)3$/.exec(key);
    if (m && s1[key].v != null && String(s1[key].v).trim()) headers[m[1]] = String(s1[key].v).trim();
  }
  const mapping = new Map<string, string>();
  const ms = wb.Sheets["mapping"];
  if (ms) for (const r of XLSX.utils.sheet_to_json<string[]>(ms, { header: 1, defval: "" }).slice(1))
    if (r[0] && r[1]) mapping.set(norm(String(r[0])), String(r[1]).trim());

  const seen = new Map<string, Set<string>>();
  for (const [title, h] of mapping) {
    const f = familyOf(title);
    if (!seen.has(f)) seen.set(f, new Set());
    seen.get(f)!.add(h);
  }
  const families = new Map<string, string>();
  for (const [f, hs] of seen) if (hs.size === 1) families.set(f, [...hs][0]);
  return { buf, headers, mapping, families };
}

/* ---------- title -> column ---------- */
function resolve(title: string, t: Template, saved: Record<string, string>): { header?: string; how: Line["how"] } {
  const n = norm(title);
  const exact = t.mapping.get(n);
  if (exact) return { header: exact, how: "template" };
  // Shopify appends " - Variant" to the product title: use the longest mapped title that is a prefix.
  let best: string | undefined, len = 0;
  for (const [k, v] of t.mapping) if (n.startsWith(k + " - ") && k.length > len) { best = v; len = k.length; }
  if (best) return { header: best, how: "template" };
  // A new size/variant of a product the template already knows (all its variants sit in one column).
  const fam = t.families.get(familyOf(title));
  if (fam) return { header: fam, how: "family" };
  // Choices the user made in the app only fill gaps; they never override the template.
  if (saved[n]) return { header: saved[n], how: "saved" };
  return { how: "none" };
}

/** The sheet has two NOR SALMON columns: the first sits in the fresh block, the second in the frozen block. */
function columnFor(header: string, title: string, t: Template): string | undefined {
  const cols = Object.entries(t.headers).filter(([, h]) => h.toUpperCase() === header.toUpperCase()).map(([c]) => c);
  if (!cols.length) return undefined;
  if (cols.length === 1) return cols[0];
  return /fresh/i.test(title) ? cols[0] : cols[cols.length - 1];
}

/** Whole numbers print as 2, entrée orders print as 1e, a mix prints as 2+1e. */
export function show(q: number, e: number): number | string {
  if (q && e) return `${q}+${e}e`;
  return e ? `${e}e` : q;
}

export function place(orders: Order[], t: Template, saved: Record<string, string> = {}, areaMode: AreaMode = "zip"): Placed[] {
  return orders.map(order => {
    const sums: Record<string, { q: number; e: number }> = {};
    const lines: Line[] = [];
    for (const it of order.items) {
      let r = resolve(it.title, t, saved);
      // "IGNORE" only means the template has no column for it, so a choice made in the app may fill it in.
      const choice = saved[norm(it.title)];
      if (r.header?.toUpperCase() === "IGNORE" && choice && choice.toUpperCase() !== "IGNORE") r = { header: choice, how: "saved" };
      const line: Line = { title: it.title, qty: it.qty, entree: isEntree(it.title), how: r.how, header: r.header };
      const col = r.header && r.header.toUpperCase() !== "IGNORE" ? columnFor(r.header, it.title, t) : undefined;
      if (col) {
        const s = (sums[col] ||= { q: 0, e: 0 });
        if (line.entree) s.e += it.qty; else s.q += it.qty; // sizes of one product add up
      } else if (NON_PRODUCT.test(it.title)) {
        line.ignored = true; line.header = "IGNORE";
      } else {
        line.spare = spareLabel(it.title); // no column of its own: add it in a spare column
        if (!r.header) line.how = "none"; else if (r.header.toUpperCase() !== "IGNORE") line.how = "none";
      }
      lines.push(line);
    }
    const cells: Record<string, number | string> = {};
    for (const [c, s] of Object.entries(sums)) cells[c] = show(s.q, s.e);
    return {
      order: { ...order, area: areaText(order, areaMode) }, cells, lines,
      spare: lines.filter(l => l.spare), unlisted: lines.filter(l => l.how === "none"),
    };
  });
}

/** Which spare column each extra product gets on one page: first come, first served. */
export function assignSpare(page: Placed[]) {
  const map: Record<string, string> = {};
  const overflow: string[] = [];
  for (const p of page) for (const l of p.spare) {
    const label = l.spare!;
    if (map[label] || overflow.includes(label)) continue;
    const free = SPARE_COLUMNS[Object.keys(map).length];
    if (free) map[label] = free; else overflow.push(label);
  }
  return { map, overflow };
}

/* ---------- writing into the template without touching anything else ---------- */
const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

export function setCell(xml: string, ref: string, value: string | number): string {
  const re = new RegExp(`<c r="${ref}"((?:\\s+[A-Za-z:]+="[^"]*")*)\\s*(?:/>|>[\\s\\S]*?</c>)`);
  const m = re.exec(xml);
  if (!m) throw new Error(`Template cell ${ref} not found`);
  const attrs = m[1].replace(/\s+t="[^"]*"/, "");
  const next = typeof value === "number"
    ? `<c r="${ref}"${attrs}><v>${value}</v></c>`
    : `<c r="${ref}"${attrs} t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`;
  return xml.replace(re, () => next);
}

export async function fillTemplate(t: Template, page: Placed[], opts: { date: string; pageNo: number; pageCount: number }): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(t.buf, { createFolders: false });
  const wbXml = await zip.file("xl/workbook.xml")!.async("string");
  const rid = /<sheet [^>]*r:id="([^"]+)"/.exec(wbXml)![1];
  const rels = await zip.file("xl/_rels/workbook.xml.rels")!.async("string");
  const target = new RegExp(`<Relationship [^>]*Id="${rid}"[^>]*>`).exec(rels)![0];
  const path = "xl/" + /Target="([^"]+)"/.exec(target)![1].replace(/^\/?(xl\/)?/, "");
  let xml = await zip.file(path)!.async("string");

  const spare = assignSpare(page);
  xml = setCell(xml, "C1", opts.date);
  xml = setCell(xml, "G1", `PAGE:  ${opts.pageNo}  OF  ${opts.pageCount}`);
  page.forEach((p, i) => {
    const r = FIRST_ROW + i;
    const n = p.order.order.replace(/^#/, "");
    xml = setCell(xml, `B${r}`, /^\d+$/.test(n) ? Number(n) : n);
    if (p.order.area) xml = setCell(xml, `C${r}`, p.order.area);
    for (const [col, v] of Object.entries(p.cells)) xml = setCell(xml, `${col}${r}`, v);
    const per: Record<string, { q: number; e: number }> = {};
    for (const l of p.spare) { const x = (per[l.spare!] ||= { q: 0, e: 0 }); if (l.entree) x.e += l.qty; else x.q += l.qty; }
    const over: string[] = [];
    for (const [label, x] of Object.entries(per)) {
      if (spare.map[label]) xml = setCell(xml, `${spare.map[label]}${r}`, show(x.q, x.e));
      else over.push(`${label} ${x.e && !x.q ? x.e + "e" : "x" + (x.q + x.e)}`);
    }
    if (over.length) xml = setCell(xml, `${OTHER_COLUMN}${r}`, "OTHER: " + over.join("; "));
  });
  for (const [label, letter] of Object.entries(spare.map)) xml = setCell(xml, `${letter}3`, label);
  zip.file(path, xml);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

const colNum = (l: string) => l.split("").reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);

/** A compact view of one sheet: only the products that were actually ordered become columns. */
export function compactTable(t: Template, page: Placed[]) {
  const spare = assignSpare(page);
  const cols = new Map<string, string>(); // column letter -> product name
  for (const p of page) for (const c of Object.keys(p.cells)) cols.set(c, t.headers[c] || c);
  for (const [label, letter] of Object.entries(spare.map)) cols.set(letter, label);
  const letters = [...cols.keys()].sort((a, b) => colNum(a) - colNum(b));
  const rows = page.map(p => {
    const per: Record<string, { q: number; e: number }> = {};
    for (const l of p.spare) { const x = (per[l.spare!] ||= { q: 0, e: 0 }); if (l.entree) x.e += l.qty; else x.q += l.qty; }
    const values: Record<string, number | string> = { ...p.cells };
    const over: string[] = [];
    for (const [label, x] of Object.entries(per)) {
      if (spare.map[label]) values[spare.map[label]] = show(x.q, x.e); else over.push(`${label} ${show(x.q, x.e)}`);
    }
    return { placed: p, values, over };
  });
  return { letters, names: cols, rows, hasOver: rows.some(r => r.over.length > 0) };
}

export const paginate = <T,>(a: T[]) => {
  const pages: T[][] = [];
  for (let i = 0; i < a.length; i += ROWS_PER_PAGE) pages.push(a.slice(i, i + ROWS_PER_PAGE));
  return pages.length ? pages : [[]];
};
