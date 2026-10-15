import JSZip from "jszip";

/** Just enough of the .xlsx format to read cells and edit one sheet in place, using only JSZip. */
export type Book = { zip: JSZip; sheets: { name: string; path: string }[]; shared: string[] };

const unesc = (s: string) =>
  s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n))
    .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16)))
    .replace(/&amp;/g, "&");

const attr = (tag: string, name: string) => {
  const m = new RegExp(`\\s${name}="([^"]*)"`).exec(tag);
  return m ? unesc(m[1]) : undefined;
};

const textOf = (xml: string) => {
  let out = "";
  for (const m of xml.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)) out += unesc(m[1]);
  return out;
};

export async function openBook(buf: ArrayBuffer | Uint8Array): Promise<Book> {
  const zip = await JSZip.loadAsync(buf, { createFolders: false });
  const wb = zip.file("xl/workbook.xml");
  const rels = zip.file("xl/_rels/workbook.xml.rels");
  if (!wb || !rels) throw new Error("This doesn't look like an Excel (.xlsx) file.");
  const wbXml = await wb.async("string");
  const relXml = await rels.async("string");
  const target = new Map<string, string>();
  for (const m of relXml.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = attr(m[0], "Id"), t = attr(m[0], "Target");
    if (id && t) target.set(id, t);
  }
  const sheets: Book["sheets"] = [];
  for (const m of wbXml.matchAll(/<sheet\b[^>]*>/g)) {
    const name = attr(m[0], "name"), rid = attr(m[0], "r:id"), t = rid ? target.get(rid) : undefined;
    if (name && t) sheets.push({ name, path: t.startsWith("/") ? t.slice(1) : "xl/" + t.replace(/^\.\//, "") });
  }
  const shared: string[] = [];
  const ss = zip.file("xl/sharedStrings.xml");
  if (ss) for (const m of (await ss.async("string")).matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)) shared.push(textOf(m[1]));
  return { zip, sheets, shared };
}

/** Cell reference -> displayed text, for every non-empty cell of one sheet part. */
export async function readCells(book: Book, path: string): Promise<Map<string, string>> {
  const xml = await book.zip.file(path)!.async("string");
  const out = new Map<string, string>();
  for (const m of xml.matchAll(/<c\b((?:\s+[A-Za-z0-9:_]+="[^"]*")*)\s*(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const ref = attr(m[1], "r");
    const inner = m[2];
    if (!ref || !inner) continue;
    const t = attr(m[1], "t");
    let v = "";
    if (t === "inlineStr") v = textOf(inner);
    else {
      const raw = /<v>([\s\S]*?)<\/v>/.exec(inner);
      if (!raw) continue;
      v = t === "s" ? book.shared[+raw[1]] ?? "" : unesc(raw[1]);
    }
    if (v.trim() !== "") out.set(ref, v);
  }
  return out;
}

export const colNum = (letters: string) => [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);

/** Rows in `[first, ..]` that exist in the sheet's XML for column B, up to the first gap. Gives the sheet's last order row. */
export function lastGridRow(xml: string, first: number): number {
  let r = first;
  while (new RegExp(`<c r="B${r}"[\\s>/]`).test(xml)) r++;
  return r - 1;
}

/** Styles (cellXfs indexes) that rotate their text. */
export function rotatedStyles(stylesXml: string): Set<number> {
  const block = /<cellXfs\b[\s\S]*?<\/cellXfs>/.exec(stylesXml)?.[0] ?? "";
  const out = new Set<number>();
  let i = 0;
  for (const m of block.matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)) {
    if (/textRotation="(?!0")/.test(m[0])) out.add(i);
    i++;
  }
  return out;
}
