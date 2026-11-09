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
export type AreaMode = "suburb" | "city";
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
export type Placed = { order: Order; cells: Record<string, number | string>; lines: Line[]; other: Line[]; otherText: string; otherCut: number };

export const FIRST_ROW = 4;
export const LAST_ROW = 32; // 29 order rows per page
export const ROWS_PER_PAGE = LAST_ROW - FIRST_ROW + 1;
/** Products that are on no list are written here as text so no order is ever dropped. */
export const OTHER_COLUMN = "AS";

/** lower-case, no accents, straight quotes, single spaces, tidy pipes */
export const norm = (s: string) =>
  s.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[\u2018\u2019\u00b4`]/g, "'")
    .toLowerCase().replace(/\s+/g, " ").replace(/\s*\|\s*/g, " | ").trim();
const isEntree = (s: string) => norm(s).includes("entree");
const familyOf = (s: string) => norm(s).split(" | ")[0].split(" - ")[0].replace(/\bentree\b/g, "").replace(/\s+/g, " ").trim();

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

/* Words that mean "this is part of the address, not an area". */
const UNIT = /\b(flat|apt|apartment|unit|suite|floor|block|bldg|building|complex|estate|house|room|po box|private bag|c\/o|cnr|corner)\b/i;
const COUNTRY = /^(south africa|za|rsa)$/i;

/** Reduce a free-text address field to just the place name: drops street numbers, units, postcodes, city/province/country. */
function cleanArea(text: string, o: Order): string {
  const same = (a: string, b: string) => !!b && norm(a) === norm(b);
  const keep = text.split(/[,;\n\/]+/).map(x => x.replace(/\b\d{4}\b/g, "").replace(/\s+/g, " ").trim()).filter(x =>
    x && !/\d/.test(x) && !UNIT.test(x) && !COUNTRY.test(x) && !same(x, o.province) && !same(x, o.city));
  return keep.length ? keep[keep.length - 1] : "";
}

export function areaText(o: Order, mode: AreaMode): string {
  if (o.collection) return "COLLECTION";
  const city = cleanArea(o.city, { ...o, city: "" }) || o.city;
  if (mode === "city") return city;
  // Suburb: address line 2 -> else a place name after the comma in the street line -> else the city. Never the street itself.
  return cleanArea(o.suburb, o) || cleanArea(o.street.replace(/^[^,]*,?/, ""), o) || city;
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

/** Products on no list go in the spare columns at the right as short text, so no order is ever dropped. */
export const OTHER_MAX_CHARS = 52; // what fits across AS:BH at 8pt
export function otherSummary(other: Line[]): { text: string; cut: number } {
  const parts = other.map(l => {
    const name = l.title.split(" | ")[0].trim();
    return `${name.length > 24 ? name.slice(0, 23).trim() + "…" : name} ${l.entree ? l.qty + "e" : "x" + l.qty}`;
  });
  let text = "OTHER: ", cut = 0;
  for (let i = 0; i < parts.length; i++) {
    const next = text + (i ? "; " : "") + parts[i];
    if (next.length > OTHER_MAX_CHARS - (i < parts.length - 1 ? 9 : 0)) { cut = parts.length - i; break; }
    text = next;
  }
  if (cut) text += ` +${cut} more`;
  return { text: parts.length ? text : "", cut };
}

export function place(orders: Order[], t: Template, saved: Record<string, string> = {}, areaMode: AreaMode = "suburb", areaEdits: Record<string, string> = {}): Placed[] {
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
    const other = lines.filter(l => l.how === "none");
    const { text: otherText, cut: otherCut } = otherSummary(other);
    return { order: { ...order, area: areaEdits[order.order] ?? areaText(order, areaMode) }, cells, lines, other, otherText, otherCut };
  });
}

/* ---------- writing into the template without touching anything else ---------- */
const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const attrsOf = (tag: string) => Object.fromEntries([...tag.matchAll(/([A-Za-z:]+)="([^"]*)"/g)].map(m => [m[1], m[2]]));
const attrStr = (a: Record<string, string>) => Object.entries(a).map(([k, v]) => ` ${k}="${v}"`).join("");

/**
 * The sheet's cells are 14pt in columns about 2.8 characters wide, so text such as "1e" or "2+1e" would spill into the
 * neighbouring product column. For cells we write TEXT into, we add a copy of that cell's own style that centres the text and
 * shrinks it to fit. Existing styles are never edited, new ones are only appended, so everything else looks exactly the same.
 */
export class StyleBook {
  changed = false;
  private cache = new Map<string, number>();
  constructor(public xml: string) {}

  variant(baseId: number, o: { fontSize?: number; center?: boolean; shrink?: boolean; upright?: boolean }): number {
    const key = `${baseId}|${o.fontSize ?? ""}|${o.center ?? ""}|${o.shrink ?? ""}|${o.upright ?? ""}`;
    const hit = this.cache.get(key);
    if (hit !== undefined) return hit;

    const xfRe = /(<cellXfs count=")(\d+)(">)([\s\S]*?)(<\/cellXfs>)/;
    const xm = xfRe.exec(this.xml)!;
    const xfs = xm[4].match(/<xf [^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)!;
    const base = xfs[baseId];
    // "upright" only matters for the few template cells that were left rotated 90° / wrapped
    const crooked = !!o.upright && /textRotation|wrapText/.test(base);
    if (!o.center && !o.shrink && !o.fontSize && !crooked) { this.cache.set(key, baseId); return baseId; }
    const head = /<xf ([^>]*?)\/?>/.exec(base)![1];
    const attrs = attrsOf(head);
    const inner = base.endsWith("/>") ? "" : base.replace(/^<xf [^>]*?>/, "").replace(/<\/xf>$/, "");
    const al = attrsOf(/<alignment [^>]*?\/>/.exec(inner)?.[0] ?? "");
    const rest = inner.replace(/<alignment [^>]*?\/>/, "");
    if (crooked) { delete al.textRotation; delete al.wrapText; }
    if (o.center) al.horizontal = "center";
    if (o.shrink) al.shrinkToFit = "1";
    attrs.applyAlignment = "1";

    if (o.fontSize) {
      const fRe = /(<fonts count=")(\d+)(")([^>]*>)([\s\S]*?)(<\/fonts>)/;
      const fm = fRe.exec(this.xml)!;
      const fonts = fm[5].match(/<font>[\s\S]*?<\/font>|<font\/>/g)!;
      const clone = fonts[Number(attrs.fontId)].replace(/<sz val="[^"]*"\/>/, `<sz val="${o.fontSize}"/>`);
      this.xml = this.xml.replace(fRe, (_m, a, n, b, c, d, e) => `${a}${Number(n) + 1}${b}${c}${d}${clone}${e}`);
      attrs.fontId = String(fonts.length);
      attrs.applyFont = "1";
    }
    const xf = `<xf${attrStr(attrs)}><alignment${attrStr(al)}/>${rest}</xf>`;
    this.xml = this.xml.replace(xfRe, (_m, a, n, b, c, d) => `${a}${Number(n) + 1}${b}${c}${xf}${d}`);
    this.changed = true;
    const id = xfs.length;
    this.cache.set(key, id);
    return id;
  }
}

export function setCell(xml: string, ref: string, value: string | number, restyle?: (styleId: number) => number): string {
  const re = new RegExp(`<c r="${ref}"((?:\\s+[A-Za-z:]+="[^"]*")*)\\s*(?:/>|>[\\s\\S]*?</c>)`);
  const m = re.exec(xml);
  if (!m) throw new Error(`Template cell ${ref} not found`);
  let attrs = m[1].replace(/\s+t="[^"]*"/, "");
  if (restyle) {
    const sid = Number(/\ss="(\d+)"/.exec(attrs)?.[1] ?? 0);
    attrs = attrs.replace(/\s+s="[^"]*"/, "") + ` s="${restyle(sid)}"`;
  }
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
  const book = new StyleBook(await zip.file("xl/styles.xml")!.async("string"));

  // text that must stay inside its own cell: centred, shrunk to fit
  const fit = (id: number) => book.variant(id, { center: true, shrink: true, upright: true });
  // numbers keep their look; the only change is straightening the few cells the template left rotated (row 11)
  const upright = (id: number) => book.variant(id, { upright: true });
  // the spare "OTHER" area spans many empty columns, so it only needs a smaller font
  const small = (id: number) => book.variant(id, { fontSize: 8, upright: true });

  xml = setCell(xml, "C1", opts.date);
  xml = setCell(xml, "G1", `PAGE:  ${opts.pageNo}  OF  ${opts.pageCount}`);
  page.forEach((p, i) => {
    const r = FIRST_ROW + i;
    const n = p.order.order.replace(/^#/, "");
    xml = setCell(xml, `B${r}`, /^\d+$/.test(n) ? Number(n) : n, upright);
    if (p.order.area) xml = setCell(xml, `C${r}`, p.order.area, id => book.variant(id, { shrink: true, upright: true }));
    for (const [col, v] of Object.entries(p.cells)) xml = setCell(xml, `${col}${r}`, v, typeof v === "string" ? fit : upright);
    if (p.otherText) xml = setCell(xml, `${OTHER_COLUMN}${r}`, p.otherText, small);
  });
  zip.file(path, xml);
  if (book.changed) zip.file("xl/styles.xml", book.xml);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

export const paginate = <T,>(a: T[]) => {
  const pages: T[][] = [];
  for (let i = 0; i < a.length; i += ROWS_PER_PAGE) pages.push(a.slice(i, i + ROWS_PER_PAGE));
  return pages.length ? pages : [[]];
};
