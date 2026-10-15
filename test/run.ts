import fs from "fs";
import JSZip from "jszip";
import { openBook, readCells } from "../lib/xl";
import {
  readTemplate, describeTemplate, readExisting, groupOrders, buildPlan, fillSheet, place, orderKey, FriendlyError, headersKey, Template, Order,
} from "../lib/greenfish";

let fails = 0;
const ok = (cond: boolean, msg: string) => { console.log(cond ? "PASS" : "FAIL", msg); if (!cond) fails++; };
const ab = (b: Buffer | Uint8Array) => b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer;

// minimal RFC4180 parser so the test needs no extra packages
function parseCsv(text: string) {
  const rows: string[][] = []; let row: string[] = [], f = "", q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) { if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c; }
    else if (c === '"') q = true;
    else if (c === ",") { row.push(f); f = ""; }
    else if (c === "\n") { row.push(f.replace(/\r$/, "")); rows.push(row); row = []; f = ""; }
    else f += c;
  }
  if (f || row.length) { row.push(f); rows.push(row); }
  const [h, ...d] = rows;
  return d.filter(r => r.length > 1).map(r => Object.fromEntries(h.map((k, i) => [k, r[i] ?? ""])));
}

async function cellsOf(bytes: Uint8Array) {
  const b = await openBook(ab(Buffer.from(bytes)));
  return { b, cells: await readCells(b, b.sheets[0].path), xml: await b.zip.file(b.sheets[0].path)!.async("string"), styles: await b.zip.file("xl/styles.xml")!.async("string") };
}

(async () => {
  const tbuf = ab(fs.readFileSync("public/template.xlsx"));
  const t = await readTemplate(tbuf);
  const info = describeTemplate(t);
  ok(t.firstRow === 4 && t.lastRow === 32, `template rows 4..32 (got ${t.firstRow}..${t.lastRow})`);
  ok(info.columns >= 30 && info.mapped > 200, `template has ${info.columns} product columns, ${info.mapped} mapped, warnings: ${JSON.stringify(info.warnings)}`);

  const rows = parseCsv(fs.readFileSync("/mnt/user-data/uploads/orders_export__26_.csv", "utf8"));
  const orders = groupOrders(rows as any);
  console.log("orders in export:", orders.length);

  // ---- 1. a plain run on the real export
  const plan = buildPlan(orders, t, [], {});
  const um = new Map<string, number>(); plan.placed.forEach(p => p.unmapped.forEach(u => um.set(u.title, (um.get(u.title) || 0) + u.qty)));
  console.log("unmapped in real export:", [...um.keys()]);
  console.log("placed", plan.placed.length, "sheets", plan.sheets.map(s => s.items.length));
  ok(plan.sheets.length === Math.ceil(plan.placed.length / 29), "pages = ceil(orders/29)");
  // pretend every unmapped product has been matched, so we can fill the sheet
  const extra: Record<string, string> = {}; for (const k of um.keys()) extra[k.toLowerCase().replace(/\s+/g, " ").replace(/\s*\|\s*/g, " | ").trim()] = "IGNORE";
  const full = buildPlan(orders, t, [], extra);
  const out1 = await fillSheet(full.sheets[0], "15/10/2026");
  const r1 = await cellsOf(out1);
  ok(r1.cells.get("C1") === "15/10/2026" && /PAGE:\s+1\s+OF\s+2/.test(r1.cells.get("G1")!), `C1/G1 written: ${r1.cells.get("C1")} | ${r1.cells.get("G1")}`);
  ok(r1.cells.get("B4") === full.sheets[0].items[0].order.order.replace(/^#/, ""), "B4 holds first order number " + r1.cells.get("B4"));
  // row 11 fix: no rotated styles left in order rows' A..C cells
  const row11 = /<row r="11"[^>]*>[\s\S]*?<\/row>/.exec(r1.xml)![0];
  ok(/<c r="C11"[^>]*s="3"/.test(row11) && /<row r="11"[^>]*ht="19.8"/.test(row11), "row 11 address cell uses the normal style and row height");
  ok(!/<c r="(A|C)11"[^>]*s="(5|6)"/.test(row11), "row 11 no longer rotated (A11, C11)");
  const cols = /<cols>[\s\S]*?<\/cols>/.exec(r1.xml)![0];
  ok(/min="2" max="2" width="11"/.test(cols) && /min="3" max="3" width="26"/.test(cols), "column B widened to 11, C to 26");
  ok(/<c r="C5"[^>]*><f>HYPERLINK/.test(r1.xml) || /<c r="C4"[^>]*><f>HYPERLINK|>COLLECTION</.test(r1.xml), "area cells written (links / COLLECTION)");

  // ---- 2. morning sheet, then later orders added to it
  const morningN = 20;
  const morning = buildPlan(orders.slice(0, morningN), t, [], extra);
  const morningFile = await fillSheet(morning.sheets[0], "15/10/2026");
  fs.writeFileSync("/tmp/morning.xlsx", morningFile);
  const ex = await readExisting("morning.xlsx", ab(Buffer.from(morningFile)));
  ok(ex.orders.length === morning.placed.length && ex.nextRow === 4 + morning.placed.length, `morning sheet read back: ${ex.orders.length} orders, next row ${ex.nextRow}, date ${ex.dateText}, page ${ex.pageNo}`);
  ok(headersKey(ex.template) === headersKey(t), "same columns as template");

  const later = buildPlan(orders, t, [ex], extra);
  ok(later.alreadyThere === morning.placed.length, `orders already on the sheet are skipped: ${later.alreadyThere} of the ${morningN} early ones were on it`);
  const total = later.placed.length + morning.placed.length;
  console.log("sheets after adding:", later.sheets.map(s => `${s.source} p${s.pageNo}/${s.pageCount} start ${s.startRow} +${s.items.length}`));
  ok(later.sheets[0].source === "existing" && later.sheets[0].startRow === ex.nextRow, "existing sheet continues at the first empty row");
  const files = await Promise.all(later.sheets.map(s => fillSheet(s, "99/99/9999")));
  const seen: string[] = [];
  for (let i = 0; i < files.length; i++) {
    const c = await cellsOf(files[i]);
    for (let r = 4; r <= 32; r++) { const v = c.cells.get(`B${r}`); if (v) seen.push(orderKey(v)); }
    ok(/PAGE:\s+\d+\s+OF\s+2/.test(c.cells.get("G1")!), `page ${i + 1} G1 = ${c.cells.get("G1")}`);
    ok(c.cells.get("C1") === "15/10/2026", `page ${i + 1} keeps the sheet's date`);
  }
  ok(new Set(seen).size === seen.length, "no order appears twice");
  ok(seen.length === total, `all ${total} orders present exactly once (found ${seen.length})`);
  const expected = [...orders.slice(0, morningN), ...later.placed.map(p => p.order)].map(o => orderKey(o.order));
  // orders are in export order on the sheets (morning ones first)
  ok(seen.slice(0, ex.orders.length).join() === ex.orders.join(), "morning rows untouched and still first");
  // the second page must hold only new orders and be blank below them
  const p2 = await cellsOf(files[1]);
  const p2orders = []; for (let r = 4; r <= 32; r++) if (p2.cells.get(`B${r}`)) p2orders.push(r);
  ok(p2orders.length === later.sheets[1].items.length && p2orders[0] === 4, "page 2 starts at row 4 with only the new orders");
  // manual edit survives: tick DELIVERED on the morning sheet, add orders again
  const edited = await (async () => {
    const z = await JSZip.loadAsync(morningFile); const p = (await openBook(ab(Buffer.from(morningFile)))).sheets[0].path;
    let x = await z.file(p)!.async("string");
    x = x.replace(/<c r="A4"((?:\s+[A-Za-z0-9:_]+="[^"]*")*)\s*\/>/, '<c r="A4"$1 t="inlineStr"><is><t>Y</t></is></c>');
    z.file(p, x); return z.generateAsync({ type: "uint8array" });
  })();
  const exEdited = await readExisting("edited.xlsx", ab(Buffer.from(edited)));
  const f2 = await fillSheet(buildPlan(orders, t, [exEdited], extra).sheets[0], "x");
  ok((await cellsOf(f2)).cells.get("A4") === "Y", "a tick typed in DELIVERED on the old sheet is kept");

  // ---- 3. the sheet after Excel saved it (shared strings instead of inline strings)
  const z = await JSZip.loadAsync(morningFile);
  const book = await openBook(ab(Buffer.from(morningFile)));
  let sx = await z.file(book.sheets[0].path)!.async("string");
  let ss = await z.file("xl/sharedStrings.xml")!.async("string");
  let idx = (ss.match(/<si>/g) || []).length; let added = "";
  sx = sx.replace(/<c r="(B\d+)"((?:\s+[A-Za-z0-9:_]+="[^"]*")*) t="inlineStr"><is><t[^>]*>([^<]*)<\/t><\/is><\/c>/g, (_m, ref, a, txt) => { added += `<si><t>${txt}</t></si>`; return `<c r="${ref}"${a} t="s"><v>${idx++}</v></c>`; });
  ss = ss.replace("</sst>", added + "</sst>");
  z.file(book.sheets[0].path, sx); z.file("xl/sharedStrings.xml", ss);
  const resaved = await z.generateAsync({ type: "uint8array" });
  const exRe = await readExisting("resaved.xlsx", ab(Buffer.from(resaved)));
  ok(exRe.orders.join() === ex.orders.join(), "reads orders from a sheet re-saved with shared strings");

  // ---- 4. template checks
  try { await readTemplate(ab(Buffer.from("not an excel file"))); ok(false, "garbage rejected"); } catch (e) { ok(e instanceof FriendlyError, "garbage file gives a friendly error: " + (e as Error).message); }
  const zNo = await JSZip.loadAsync(tbuf); const bk = await openBook(tbuf);
  let sh = await zNo.file(bk.sheets[0].path)!.async("string");
  sh = sh.replace(/<c r="B3"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/, '<c r="B3" s="11"/>');
  zNo.file(bk.sheets[0].path, sh);
  try { await readTemplate(ab(Buffer.from(await zNo.generateAsync({ type: "uint8array" })))); ok(false, "wrong layout rejected"); } catch (e) { ok(e instanceof FriendlyError, "sheet without ORDER# gives a friendly error: " + (e as Error).message); }

  // ---- 5. a new template with an extra product column + mapping row
  const zc = await JSZip.loadAsync(tbuf);
  let s1 = await zc.file(bk.sheets[0].path)!.async("string");
  const mp = bk.sheets.find(s => s.name === "mapping")!.path;
  let s2 = await zc.file(mp)!.async("string");
  ok(/<c r="AS3"/.test(s1), "template has a spare formatted header cell AS3");
  s1 = s1.replace(/<c r="AS3"((?:\s+[A-Za-z0-9:_]+="[^"]*")*)\s*(?:\/>|>[\s\S]*?<\/c>)/, '<c r="AS3"$1 t="inlineStr"><is><t>BARRAMUNDI</t></is></c>');
  const lastMapRow = Math.max(...[...s2.matchAll(/<row r="(\d+)"/g)].map(m => +m[1]));
  s2 = s2.replace("</sheetData>", `<row r="${lastMapRow + 1}"><c r="A${lastMapRow + 1}" t="inlineStr"><is><t>Barramundi Fillet | Fresh | 1kg</t></is></c><c r="B${lastMapRow + 1}" t="inlineStr"><is><t>BARRAMUNDI</t></is></c></row></sheetData>`);
  zc.file(bk.sheets[0].path, s1); zc.file(mp, s2);
  const custom = await readTemplate(ab(Buffer.from(await zc.generateAsync({ type: "uint8array" }))));
  ok(custom.headers["AS"] === "BARRAMUNDI" && custom.mapping.has("barramundi fillet | fresh | 1kg"), "custom template: new column and mapping row are read");
  const newOrder: Order = { order: "#GF70001", dest: { kind: "collection", label: "COLLECTION" }, items: [{ title: "Barramundi Fillet | Fresh | 1kg", qty: 3 }] };
  const bad = place([newOrder], t, {});
  const good = place([newOrder], custom, {});
  ok(bad[0].unmapped.length === 1, "with the old template the new product is flagged as unmatched");
  ok(good[0].unmapped.length === 0 && good[0].cells["AS"] === 3, "with the new template it lands in the new column");
  const cf = await fillSheet(buildPlan([newOrder], custom, [], {}).sheets[0], "15/10/2026");
  ok((await cellsOf(cf)).cells.get("AS4") === "3", "new column value written to AS4");

  console.log(fails ? `\n${fails} FAILED` : "\nALL PASSED");
  process.exit(fails ? 1 : 0);
})();
