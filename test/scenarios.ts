import fs from "fs";
import JSZip from "jszip";
import {
  readTemplate,
  groupOrders,
  place,
  fillTemplate,
  paginate,
} from "../lib/greenfish";

const buf = fs.readFileSync("public/template.xlsx");
const ab = buf.buffer.slice(
  buf.byteOffset,
  buf.byteOffset + buf.byteLength,
) as ArrayBuffer;
const t = readTemplate(ab);
let fails = 0;
const ok = (c: boolean, msg: string) => {
  console.log(c ? "  PASS" : "  FAIL", msg);
  if (!c) fails++;
};

// one Shopify row helper (Shopify only fills the address on an order's first row)
const R = (
  name: string,
  title: string,
  qty: number,
  addr: Record<string, string> = {},
) => ({
  "Shipping Address1": "",
  "Shipping Address2": "",
  "Shipping City": "",
  "Shipping Province Name": "",
  "Shipping Zip": "",
  "Shipping Method": addr["Shipping Address1"] ? "Delivery" : "",
  ...addr,
  Name: name,
  "Lineitem name": title,
  "Lineitem quantity": String(qty),
});
const col = (header: string) =>
  Object.entries(t.headers)
    .filter(([, h]) => h === header)
    .map(([c]) => c);

const rows = [
  // 1 + 2: Prawn Lover plain and entrée in the same order, squid entrée, scallop entrée alone
  R("#2001", "Prawn Lover | Frozen Box | Wild caught | 2kg", 2, {
    "Shipping Address1": "1 A St",
    "Shipping City": "Cape Town",
    "Shipping Address2": "Sea Point",
  }),
  R("#2001", "Prawn Lover Entrée | Frozen Box | Wild caught | 650g", 1),
  R(
    "#2001",
    "Squid Lover’s Box Entrée | Calamari | Frozen Box | Wild caught | 900g",
    1,
  ),
  R("#2002", "Large Scallops Entrée | Frozen Box | Wild caught | 300g", 3, {
    "Shipping Address1": "2 B St",
    "Shipping City": "Cape Town",
    "Shipping Address2": "Gardens",
  }),
  R(
    "#2003",
    "Squid Lover's Box | Calamari | Frozen Box | Wild caught | 3x 900g",
    1,
    {
      "Shipping Address1": "3 C St",
      "Shipping City": "Cape Town",
      "Shipping Address2": "Gardens",
    },
  ),
  R("#2004", "Prawn Lover | Frozen Box | Wild caught | 2kg", 1, {
    "Shipping Address1": "4 D St",
    "Shipping City": "Cape Town",
    "Shipping Address2": "Gardens",
  }),
  R("#2005", "Prawn Lover Entrée - 650g", 1, {
    "Shipping Address1": "5 E St",
    "Shipping City": "Cape Town",
    "Shipping Address2": "Gardens",
  }), // variant suffix form
  // 3: messy addresses, only the area may come out
  R("#2010", "Hake Fillets | Frozen Box | Premium Fillet Range | +-900g", 1, {
    "Shipping Address1": "12 Kloof Street",
    "Shipping Address2": "Flat 4, Ocean View, Sea Point",
    "Shipping City": "Cape Town",
    "Shipping Province Name": "Western Cape",
    "Shipping Zip": "8005",
  }),
  R("#2011", "Hake Fillets | Frozen Box | Premium Fillet Range | +-900g", 1, {
    "Shipping Address1": "5 Main Rd",
    "Shipping Address2": "Apt 3B",
    "Shipping City": "Claremont",
  }),
  R("#2012", "Hake Fillets | Frozen Box | Premium Fillet Range | +-900g", 1, {
    "Shipping Address1": "9 Vineyard Rd, Constantia",
    "Shipping Address2": "",
    "Shipping City": "Cape Town",
    "Shipping Zip": "7806",
  }),
  R("#2013", "Hake Fillets | Frozen Box | Premium Fillet Range | +-900g", 1, {
    "Shipping Address1": "7 Bay Rd",
    "Shipping Address2": "Green Point, Cape Town 8005",
    "Shipping City": "Cape Town",
  }),
  // 4: products that are on no list, long names, several on one order
  R("#2020", "Brand New Product | Frozen Box", 2, {
    "Shipping Address1": "1 Z St",
    "Shipping Address2": "Hout Bay",
    "Shipping City": "Cape Town",
  }),
  R(
    "#2021",
    "Kingklip (1.0kg) & Hake (1.5kg) Fillets Special Edition | Fresh",
    1,
    {
      "Shipping Address1": "1 Y St",
      "Shipping Address2": "Hout Bay",
      "Shipping City": "Cape Town",
    },
  ),
  R("#2021", "Another Unlisted Seafood Platter | 2kg", 2),
  R("#2021", "Third Unlisted Thing | 1kg", 1),
  R("#2021", "Fourth Unlisted Thing | 1kg", 1),
  R("#2022", "Hake Fillets | Frozen Box | Premium Fillet Range | +-900g", 1, {
    "Shipping Address1": "1 X St",
    "Shipping Address2": "Hout Bay",
    "Shipping City": "Cape Town",
  }),
  R("#2022", "Spring rolls | Vegetable | Frozen box | 10x 45g", 1),
];
const placed = place(groupOrders(rows as any), t);
const by = Object.fromEntries(placed.map((p) => [p.order.order, p]));
const [PL] = col("PRAWN LOVER"),
  [SQ] = col("SQUID"),
  [SC] = col("SCALLOP");

console.log("\n#1 Prawn Lover goes to the PRAWN LOVER column only");
ok(
  by["#2001"].cells[PL] === "2+1e",
  `plain + entrée in one order -> ${JSON.stringify(by["#2001"].cells[PL])} in ${PL}`,
);
ok(
  by["#2004"].cells[PL] === 1 && Object.keys(by["#2004"].cells).length === 1,
  "plain Prawn Lover x1 -> 1, nowhere else",
);
ok(
  by["#2005"].cells[PL] === "1e",
  `"Prawn Lover Entrée - 650g" (variant suffix) -> ${JSON.stringify(by["#2005"].cells)}`,
);
console.log("\n#2 Entrée prints as 1e, plain as 1");
ok(
  by["#2001"].cells[SQ] === "1e",
  `Squid Entrée -> ${JSON.stringify(by["#2001"].cells[SQ])}`,
);
ok(
  by["#2003"].cells[SQ] === 1,
  `Squid (not entrée) -> ${JSON.stringify(by["#2003"].cells[SQ])}`,
);
ok(
  by["#2002"].cells[SC] === "3e",
  `Scallop Entrée x3 -> ${JSON.stringify(by["#2002"].cells[SC])}`,
);
console.log("\n#3 Area only, never the address");
for (const [id, want] of [
  ["#2010", "Sea Point"],
  ["#2011", "Claremont"],
  ["#2012", "Constantia"],
  ["#2013", "Green Point"],
  ["#2001", "Sea Point"],
])
  ok(
    by[id].order.area === want,
    `${id} area = ${JSON.stringify(by[id].order.area)} (want ${want})`,
  );
console.log("\n#4 Unlisted products never dropped");
ok(
  by["#2020"].otherText === "OTHER: Brand New Product x2",
  JSON.stringify(by["#2020"].otherText),
);
ok(
  by["#2021"].other.length === 4,
  "4 unlisted products on #2021 are all captured",
);
ok(
  by["#2021"].otherText.length <= 52 && by["#2021"].otherCut > 0,
  `long list is shortened to fit, with a count: ${JSON.stringify(by["#2021"].otherText)}`,
);
ok(
  by["#2022"].otherText === "" && by["#2022"].cells[col("HAKE")[0]] === 1,
  "listed + ignored products unaffected",
);

(async () => {
  const out = await fillTemplate(t, paginate(placed)[0], {
    date: "15/10/2026",
    pageNo: 1,
    pageCount: 1,
  });
  fs.writeFileSync("scen-output.xlsx", out);
  const a = await JSZip.loadAsync(buf),
    b = await JSZip.loadAsync(out);
  console.log("\nTemplate integrity");
  for (const n of Object.keys(a.files))
    if (!a.files[n].dir) {
      const x = Buffer.from(await a.file(n)!.async("uint8array")),
        y = Buffer.from(await b.file(n)!.async("uint8array"));
      if (Buffer.compare(x, y)) console.log("  changed part:", n);
    }
  const sa = await a.file("xl/styles.xml")!.async("string"),
    sb = await b.file("xl/styles.xml")!.async("string");
  const xfs = (s: string) =>
    /<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/
      .exec(s)![1]
      .match(/<xf [^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)!;
  const fonts = (s: string) =>
    /<fonts[^>]*>([\s\S]*?)<\/fonts>/
      .exec(s)![1]
      .match(/<font>[\s\S]*?<\/font>|<font\/>/g)!;
  ok(
    xfs(sb).slice(0, xfs(sa).length).join() === xfs(sa).join(),
    `all ${xfs(sa).length} original cell styles untouched (${xfs(sb).length - xfs(sa).length} appended)`,
  );
  ok(
    fonts(sb).slice(0, fonts(sa).length).join() === fonts(sa).join(),
    "all original fonts untouched",
  );
  ok(
    sb
      .replace(/<cellXfs[\s\S]*?<\/cellXfs>/, "")
      .replace(/<fonts[\s\S]*?<\/fonts>/, "") ===
      sa
        .replace(/<cellXfs[\s\S]*?<\/cellXfs>/, "")
        .replace(/<fonts[\s\S]*?<\/fonts>/, ""),
    "rest of styles.xml byte-identical",
  );
  const strip = (s: string) =>
    s.replace(
      /<c r="(?:C1|G1|[A-Z]+(?:[4-9]|1\d|2\d|3[0-2]))"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g,
      (m) => m.replace(/<c r="([A-Z]+\d+)"[\s\S]*/, '<c r="$1"/>'),
    );
  ok(
    strip(await a.file("xl/worksheets/sheet1.xml")!.async("string")) ===
      strip(await b.file("xl/worksheets/sheet1.xml")!.async("string")),
    "everything in the sheet outside the written cells is identical",
  );
  // row 11 of the template has leftover rotated/wrapped cells: anything we write there must be upright
  const sh = await b.file("xl/worksheets/sheet1.xml")!.async("string");
  const xl = xfs(sb);
  const bad = [
    ...sh.matchAll(
      /<c r="([A-Z]+)(\d+)"[^>]*?s="(\d+)"[^>]*?>\s*(?:<v>|<is>)/g,
    ),
  ]
    .filter(
      (m) =>
        Number(m[2]) >= 4 && /textRotation|wrapText/.test(xl[Number(m[3])]),
    )
    .map((m) => m[1] + m[2]);
  ok(
    bad.length === 0,
    `no written cell is rotated or wrapped ${bad.length ? bad.join() : ""}`,
  );
  console.log(fails ? `\n${fails} FAILED` : "\nALL PASSED");
  process.exit(fails ? 1 : 0);
})();
