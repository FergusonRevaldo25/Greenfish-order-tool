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
export type Line = { title: string; qty: number; entree: boolean; header?: string; ignored?: boolean; how: "template" | "family" | "saved" | "none" };
/** `other` is what gets written as OTHER text on the sheet; `unlisted` is products found on no list at all. */
export type Placed = { order: Order; cells: Record<string, number | string>; lines: Line[]; other: Line[]; unlisted: Line[] };
/** One upload. Once added, a batch is never changed by later uploads. */
export type Batch = { id: string; at: string; fileName: string; orders: Order[]; skipped: number };

export const FIRST_ROW = 4;
export const LAST_ROW = 32; // 29 order rows per page
export const ROWS_PER_PAGE = LAST_ROW - FIRST_ROW + 1;
/** Products that are on no list are written here as text so no order is ever dropped. */
export const OTHER_COLUMN = "AS";

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
function show(q: number, e: number): number | string {
  if (q && e) return `${q}+${e}e`;
  return e ? `${e}e` : q;
}

export function place(orders: Order[], t: Template, saved: Record<string, string> = {}, areaMode: AreaMode = "zip", includeIgnored = false): Placed[] {
  return orders.map(order => {
    const sums: Record<string, { q: number; e: number }> = {};
    const lines: Line[] = [];
    for (const it of order.items) {
      const r = resolve(it.title, t, saved);
      const line: Line = { title: it.title, qty: it.qty, entree: isEntree(it.title), how: r.how };
      const col = r.header && r.header.toUpperCase() !== "IGNORE" ? columnFor(r.header, it.title, t) : undefined;
      if (r.header?.toUpperCase() === "IGNORE") { line.ignored = true; line.header = "IGNORE"; }
      else if (col) {
        line.header = r.header;
        const s = (sums[col] ||= { q: 0, e: 0 });
        if (line.entree) s.e += it.qty; else s.q += it.qty; // sizes of one product add up
      } else line.how = "none";
      lines.push(line);
    }
    const cells: Record<string, number | string> = {};
    for (const [c, s] of Object.entries(sums)) cells[c] = show(s.q, s.e);
    const unlisted = lines.filter(l => l.how === "none");
    return { order: { ...order, area: areaText(order, areaMode) }, cells, lines, unlisted, other: includeIgnored ? lines.filter(l => l.how === "none" || l.ignored) : unlisted };
  });
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

  xml = setCell(xml, "C1", opts.date);
  xml = setCell(xml, "G1", `PAGE:  ${opts.pageNo}  OF  ${opts.pageCount}`);
  page.forEach((p, i) => {
    const r = FIRST_ROW + i;
    const n = p.order.order.replace(/^#/, "");
    xml = setCell(xml, `B${r}`, /^\d+$/.test(n) ? Number(n) : n);
    if (p.order.area) xml = setCell(xml, `C${r}`, p.order.area);
    for (const [col, v] of Object.entries(p.cells)) xml = setCell(xml, `${col}${r}`, v);
    if (p.other.length)
      xml = setCell(xml, `${OTHER_COLUMN}${r}`, "OTHER: " + p.other.map(l => `${l.title.split(" | ")[0]} ${l.entree ? l.qty + "e" : "x" + l.qty}`).join("; "));
  });
  zip.file(path, xml);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

export const paginate = <T,>(a: T[]) => {
  const pages: T[][] = [];
  for (let i = 0; i < a.length; i += ROWS_PER_PAGE) pages.push(a.slice(i, i + ROWS_PER_PAGE));
  return pages.length ? pages : [[]];
};
