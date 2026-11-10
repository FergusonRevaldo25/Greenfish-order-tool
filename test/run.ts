import fs from "fs";
import JSZip from "jszip";
import { readTemplate, groupOrders, place, fillTemplate, paginate } from "../lib/greenfish";
const buf = fs.readFileSync("public/template.xlsx");
const t = readTemplate(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer);
const base = { "Shipping City": "", "Shipping Address1": "", "Shipping Address2": "", "Shipping Method": "" };
const addr = { "Shipping Zip": "8005", "Shipping Address1": "12 Main Rd", "Shipping Address2": "Sea Point", "Shipping City": "Cape Town", "Shipping Method": "Delivery" };
const rows = [
  { ...base, ...addr, Name: "#1001", "Lineitem name": "Prawn Lover | Frozen Box | Wild caught | 2kg", "Lineitem quantity": "2" },
  { ...base, Name: "#1001", "Lineitem name": "Prawn Lover Entrée | Frozen Box | Wild caught | 650g", "Lineitem quantity": "1" },
  { ...base, Name: "#1001", "Lineitem name": "Squid Lover’s Box Entrée | Calamari | Frozen Box | Wild caught | 900g", "Lineitem quantity": "1" },
  { ...base, Name: "#1002", "Shipping Method": "Collection - Claremont", "Lineitem name": "Prawn Lover | Frozen Box | Wild caught | 5kg", "Lineitem quantity": "1" },
  { ...base, ...addr, "Shipping Address2": "Hout Bay", Name: "#1003", "Lineitem name": "Totally New Product | 1kg", "Lineitem quantity": "3" },
];
const placed = place(groupOrders(rows as any), t, { "prawn lover | frozen box | wild caught | 5kg": "HAKE" });
for (const p of placed) console.log(p.order.order, p.order.area, JSON.stringify(p.cells), p.lines.map(l => l.how).join(","), p.other.length);
(async () => {
  const out = await fillTemplate(t, paginate(placed)[0], { date: "15/10/2026", pageNo: 1, pageCount: 1 });
  fs.writeFileSync("/tmp/out.xlsx", out);
  const a = await JSZip.loadAsync(buf), b = await JSZip.loadAsync(out);
  for (const n of Object.keys(a.files)) if (!a.files[n].dir) {
    const x = Buffer.from(await a.file(n)!.async("uint8array")), y = Buffer.from(await b.file(n)!.async("uint8array"));
    if (Buffer.compare(x, y)) console.log("changed part:", n);
  }
  const strip = (s: string) => s.replace(/<c r="(?:C1|G1|[A-Z]+(?:[4-9]|1\d|2\d|3[0-2]))"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g, m => m.replace(/<c r="([A-Z]+\d+)"[^>]*?s="(\d+)"[\s\S]*/, '<c r="$1" s="$2"/>'));
  console.log("rest of sheet identical:", strip(await a.file("xl/worksheets/sheet1.xml")!.async("string")) === strip(await b.file("xl/worksheets/sheet1.xml")!.async("string")));
})();

import { mergeBatch, Batch } from "../lib/greenfish";
const first = groupOrders(rows.slice(0, 4) as any), second = groupOrders(rows as any);
let day: Batch[] = [];
let r = mergeBatch(day, first, "11h.csv", "2026-11-10T11:00:00Z"); day = r.batches; console.log("11:00 added", r.added, "skipped", r.skipped);
const edited = second.map(o => o.order === "#1001" ? { ...o, items: [] } : o);   // re-export with a changed #1001
r = mergeBatch(day, edited, "14h.csv", "2026-11-10T14:00:00Z"); day = r.batches; console.log("14:00 added", r.added, "skipped", r.skipped);
console.log("order of orders:", day.flatMap(b => b.orders.map(o => o.order)).join(","), "| #1001 still has", day[0].orders[0].items.length, "items");
r = mergeBatch(day, edited, "15h.csv"); console.log("same file again added", r.added, "batches", r.batches.length);
