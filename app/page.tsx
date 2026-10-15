"use client";
import { useEffect, useMemo, useRef, useState } from "react";
import JSZip from "jszip";
import {
  Template, Existing, Order, FriendlyError, groupOrders, readTemplate, readExisting, describeTemplate,
  buildPlan, fillSheet, headersKey, isProductHeader, norm,
} from "@/lib/greenfish";
import { readExport } from "@/lib/exportFile";
import { loadSavedTemplate, saveTemplate, clearSavedTemplate } from "@/lib/store";

const KEY = "greenfish-extra-mapping";
const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const today = () => new Date().toISOString().slice(0, 10);
const msg = (e: unknown, fallback: string) => (e instanceof FriendlyError ? e.message : fallback);

function save(blob: Blob, name: string) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 10000);
}

function Drop({ title, hint, accept, multiple, compact, onFiles }: {
  title: string; hint: string; accept: string; multiple?: boolean; compact?: boolean; onFiles: (f: File[]) => void;
}) {
  const [over, setOver] = useState(false);
  return (
    <label
      className={`drop${over ? " over" : ""}${compact ? " compact" : ""}`}
      onDragOver={e => { e.preventDefault(); setOver(true); }}
      onDragLeave={() => setOver(false)}
      onDrop={e => { e.preventDefault(); setOver(false); onFiles([...e.dataTransfer.files]); }}
    >
      <input type="file" accept={accept} multiple={multiple} onChange={e => { onFiles([...(e.target.files || [])]); e.target.value = ""; }} />
      <strong>{title}</strong>
      <span>{hint}</span>
    </label>
  );
}

export default function Home() {
  const [builtin, setBuiltin] = useState<Template | null>(null);
  const [tpl, setTpl] = useState<Template | null>(null);
  const [tplName, setTplName] = useState("");
  const [tplCustom, setTplCustom] = useState(false);
  const [tplNote, setTplNote] = useState("");
  const [tplError, setTplError] = useState("");
  const [existing, setExisting] = useState<Existing[]>([]);
  const [existingError, setExistingError] = useState("");
  const [orders, setOrders] = useState<Order[]>([]);
  const [extra, setExtra] = useState<Record<string, string>>({});
  const [date, setDate] = useState(today());
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const tplInput = useRef<HTMLInputElement>(null);

  useEffect(() => {
    (async () => {
      try {
        const b = await (await fetch("/template.xlsx")).arrayBuffer();
        const built = await readTemplate(b);
        setBuiltin(built); setTpl(built); setTplName("Built-in Greenfish template");
        const saved = await loadSavedTemplate();
        if (saved) {
          try { setTpl(await readTemplate(saved.buf)); setTplName(saved.name); setTplCustom(true); }
          catch { await clearSavedTemplate(); }
        }
      } catch { setError("Could not load the sheet template."); }
    })();
    try { setExtra(JSON.parse(localStorage.getItem(KEY) || "{}")); } catch {}
  }, []);

  const plan = useMemo(() => (tpl ? buildPlan(orders, tpl, existing, extra) : null), [orders, tpl, existing, extra]);
  const placed = plan?.placed ?? [];
  const sheets = plan?.sheets ?? [];
  const active = plan?.template ?? tpl;
  const tplInfo = useMemo(() => (tpl ? describeTemplate(tpl) : null), [tpl]);
  const unmapped = useMemo(() => {
    const m = new Map<string, number>();
    placed.forEach(p => p.unmapped.forEach(u => m.set(u.title, (m.get(u.title) || 0) + u.qty)));
    return [...m.keys()];
  }, [placed]);
  const options = useMemo(() => active ? ["IGNORE", ...Array.from(new Set(Object.values(active.headers).filter(isProductHeader)))] : [], [active]);

  const deliveries = placed.filter(p => p.order.dest.kind === "delivery").length;
  const collections = placed.filter(p => p.order.dest.kind === "collection").length;
  const noAddress = placed.filter(p => p.order.dest.kind === "unknown").length;
  const saved = Object.entries(extra);
  const continuing = existing.length > 0;
  const sheetDate = existing.find(e => e.dateText)?.dateText || "";
  const differs = !!(tpl && continuing && headersKey(existing[0].template) !== headersKey(tpl));
  const toAdd = sheets.reduce((n, s) => n + s.items.length, 0);

  async function load(files: File[]) {
    const f = files[0];
    if (!f) return;
    setError("");
    try { setOrders(groupOrders(await readExport(f))); setFileName(f.name); }
    catch { setError("That file could not be read. Please use the Shopify orders export (CSV or Excel)."); }
  }

  async function addExisting(files: File[]) {
    setExistingError("");
    const next = [...existing];
    const problems: string[] = [];
    for (const f of files) {
      try {
        if (!/\.xlsx$/i.test(f.name)) throw new FriendlyError("Please use the .xlsx sheet you downloaded from this tool (or saved from Excel).");
        const e = await readExisting(f.name, await f.arrayBuffer());
        if (next.some(x => x.name === e.name && x.orders.join() === e.orders.join())) continue;
        if (next.length && headersKey(next[0].template) !== headersKey(e.template))
          throw new FriendlyError("It has different product columns from the sheet you already added, so they can't be combined.");
        next.push(e);
      } catch (err) { problems.push(`${f.name}: ${msg(err, "that file could not be read.")}`); }
    }
    setExisting(next);
    if (problems.length) setExistingError(problems.join(" "));
  }

  async function uploadTemplate(files: File[]) {
    const f = files[0];
    if (!f) return;
    setTplError(""); setTplNote("");
    try {
      if (!/\.xlsx$/i.test(f.name)) throw new FriendlyError("Please choose an Excel (.xlsx) file.");
      const buf = await f.arrayBuffer();
      const t = await readTemplate(buf);
      setTpl(t); setTplName(f.name); setTplCustom(true);
      const stored = await saveTemplate({ name: f.name, buf, savedAt: Date.now() });
      setTplNote(stored ? "New template loaded. It will be here next time too." : "New template loaded for this visit. Your browser wouldn't keep it, so upload it again next time.");
    } catch (e) { setTplError(msg(e, "That template could not be read.")); }
  }
  async function useBuiltin() {
    if (!builtin) return;
    await clearSavedTemplate();
    setTpl(builtin); setTplName("Built-in Greenfish template"); setTplCustom(false); setTplError(""); setTplNote("Back on the built-in template.");
  }
  function downloadTemplate() {
    if (tpl) save(new Blob([tpl.buf], { type: XLSX_TYPE }), tplCustom ? tplName : "Greenfish daily sales template.xlsx");
  }

  function store(next: Record<string, string>) { setExtra(next); localStorage.setItem(KEY, JSON.stringify(next)); }
  const assign = (title: string, col: string) => store({ ...extra, [norm(title)]: col });
  const forget = (key: string) => { const next = { ...extra }; delete next[key]; store(next); };

  async function download() {
    if (!plan) return;
    setBusy(true); setError("");
    try {
      const [d, m, y] = (sheetDate || "").split("/");
      const iso = /^\d{1,2}$/.test(d || "") && /^\d{4}$/.test(y || "") ? `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}` : date;
      const [py, pm, pd] = date.split("-");
      const files = await Promise.all(plan.sheets.map(s => fillSheet(s, `${pd}/${pm}/${py}`)));
      const names = plan.sheets.map(s => `Greenfish daily sales ${iso}${files.length > 1 ? ` (page ${s.pageNo})` : ""}.xlsx`);
      if (files.length === 1) save(new Blob([files[0] as BlobPart], { type: XLSX_TYPE }), names[0]);
      else {
        const zip = new JSZip();
        files.forEach((b, i) => zip.file(names[i], b));
        save(await zip.generateAsync({ type: "blob" }), `Greenfish daily sales ${iso}.zip`);
      }
    } catch (e) { setError(msg(e, "Something went wrong while filling the sheet.")); }
    setBusy(false);
  }

  const barText = !fileName ? "" :
    unmapped.length > 0 ? "Match the products above to enable the download." :
    !toAdd ? (continuing ? "Nothing new to add. Every order in this export is already on your sheet." : "No orders to put on the sheet.") :
    continuing ? `${toAdd} new ${toAdd === 1 ? "order" : "orders"} will be added${sheets.length > 1 ? ` across ${sheets.length} sheets` : ""}.` :
    `${toAdd} orders ready${sheets.length > 1 ? ` in ${sheets.length} sheets` : ""}.`;

  return (
    <>
      <header className="hero">
        <div className="wrap">
          <p className="brand">Greenfish</p>
          <h1>Daily sales sheet</h1>
          <p>Drop in your Shopify export and get your Greenfish sheet filled in, with a map link for every delivery.</p>
        </div>
      </header>

      <main className="main">
        <section className="card">
          <h2>Shopify export</h2>
          <p className="sub">CSV or Excel. Everything stays in your browser.</p>
          <Drop title={fileName ? "Choose a different file" : "Drop your export here"} hint="or click to browse" accept=".csv,.xlsx,.xls" onFiles={load} />
          <div className="row">
            {fileName && <span className="file-pill">{fileName}</span>}
            {sheetDate
              ? <span className="file-pill">Sheet date {sheetDate} (from your sheet)</span>
              : <label className="field">Sheet date
                  <input type="date" value={date} onChange={e => setDate(e.target.value)} />
                </label>}
          </div>
          {error && <p className="error" role="alert" style={{ marginTop: 14 }}>{error}</p>}
        </section>

        <section className="card">
          <h2>Adding to a sheet you already started?</h2>
          <p className="sub">
            Drop the sheet here (the one you have been working in, so your ticks stay) and then your newer export.
            Orders already on the sheet are skipped, and the new ones go into the next empty rows.
          </p>
          <Drop compact title="Drop your sheet so far" hint="or click to browse. You can add more than one page." accept=".xlsx" multiple onFiles={addExisting} />
          {existingError && <p className="error" role="alert" style={{ marginTop: 14 }}>{existingError}</p>}
          {continuing && (
            <ul className="orders" style={{ marginTop: 14 }}>
              {existing.map((e, i) => (
                <li className="order line" key={e.name + i}>
                  <span><b>{e.name}</b> <span className="muted">{e.orders.length} orders{e.nextRow > e.template.lastRow ? ", full" : `, ${e.template.lastRow - e.nextRow + 1} rows free`}</span></span>
                  <button className="btn ghost" onClick={() => setExisting(existing.filter((_, j) => j !== i))}>Remove</button>
                </li>
              ))}
            </ul>
          )}
          {differs && <p className="note" style={{ marginTop: 14 }}>This sheet's columns are different from your current template, so new orders are placed using the sheet's own columns. To use new columns, start a fresh sheet instead.</p>}
        </section>

        <section className="card">
          <h2>Sheet template</h2>
          <p className="sub">
            <b>{tplName || "Loading…"}</b>
            {tplInfo && <> · {tplInfo.columns} product columns · {tplInfo.mapped} products matched · {tplInfo.rows} orders per sheet</>}
          </p>
          <div className="row" style={{ marginTop: 0 }}>
            <button className="btn" onClick={() => tplInput.current?.click()}>Upload new template</button>
            <button className="btn ghost" onClick={downloadTemplate} disabled={!tpl}>Download current template</button>
            {tplCustom && <button className="btn ghost" onClick={useBuiltin}>Use built-in template</button>}
            <input ref={tplInput} className="vh" type="file" accept=".xlsx" tabIndex={-1} aria-hidden="true"
              onChange={e => { uploadTemplate([...(e.target.files || [])]); e.target.value = ""; }} />
          </div>
          {tplInfo?.warnings.map(w => <p className="note" key={w} style={{ marginTop: 14 }}>{w}</p>)}
          {tplNote && <p className="ok" role="status" style={{ marginTop: 14 }}>{tplNote}</p>}
          {tplError && <p className="error" role="alert" style={{ marginTop: 14 }}>{tplError}</p>}
          <p className="sub" style={{ margin: "14px 0 0" }}>
            Need a column for a new product? Download the template, add the product heading in row 3 of Sheet1 and its Shopify title in the
            <i> mapping</i> sheet, save it, then upload it here. It is kept for next time.
          </p>
        </section>

        {unmapped.length > 0 && (
          <section className="alert" role="alert">
            <h2>{unmapped.length} {unmapped.length === 1 ? "product isn't" : "products aren't"} on your mapping list</h2>
            <span>Choose a column for each one, or IGNORE to leave it off the sheet. The app remembers your choice. If the product needs a column of its own, add it to the template above.</span>
            {unmapped.map(t => (
              <div className="item" key={t}>
                <span>{t}</span>
                <select className="field" defaultValue="" aria-label={`Column for ${t}`} onChange={e => e.target.value && assign(t, e.target.value)}
                  style={{ padding: "8px 12px", border: "1px solid #f1d9a3", borderRadius: 10, background: "#fff" }}>
                  <option value="">Choose a column</option>
                  {options.map(o => <option key={o}>{o}</option>)}
                </select>
              </div>
            ))}
          </section>
        )}

        {fileName && (
          <>
            <section className="stats" aria-label="Summary">
              <div className="stat"><b>{placed.length}</b><span>{continuing ? "New orders" : "Orders"}</span></div>
              <div className="stat"><b>{deliveries}</b><span>Deliveries</span></div>
              <div className="stat"><b>{collections}</b><span>Collections</span></div>
              <div className="stat"><b>{sheets.length}</b><span>{sheets.length === 1 ? "Sheet" : "Sheets"}</span></div>
            </section>

            {continuing && plan && (
              <section className="card">
                <h2>What will happen</h2>
                <p className="sub">{plan.alreadyThere} {plan.alreadyThere === 1 ? "order from this export is" : "orders from this export are"} already on your sheet and will be skipped.</p>
                <ul className="orders">
                  {sheets.map(s => (
                    <li className="order line" key={s.pageNo}>
                      <span>
                        <b>Sheet {s.pageNo} of {s.pageCount}</b>{" "}
                        <span className="muted">
                          {s.source === "existing" ? s.existingName : "new sheet"}
                          {s.items.length ? ` · adds ${s.items.length} ${s.items.length === 1 ? "order" : "orders"} in rows ${s.startRow}–${s.startRow + s.items.length - 1}` : " · no change except the page number"}
                        </span>
                      </span>
                    </li>
                  ))}
                </ul>
              </section>
            )}

            {noAddress > 0 && (
              <section className="alert">
                <h2>{noAddress} {noAddress === 1 ? "order has" : "orders have"} no address</h2>
                <span>These are not marked as collections but the export has no shipping address for them. Their area cell will be left empty.</span>
              </section>
            )}

            <section className="card">
              <h2>{continuing ? "New orders" : "Orders"}</h2>
              <p className="sub">Each delivery links to Google Maps from the full address in the export.</p>
              <ul className="orders">
                {placed.map(p => {
                  const d = p.order.dest;
                  return (
                    <li className="order" key={p.order.order}>
                      <span className="num">{p.order.order}</span>
                      <div className="where">
                        {d.kind === "collection" && <span className="badge collection">Collection</span>}
                        {d.kind === "delivery" && <>
                          <span className="badge delivery">Delivery</span>
                          <span className="addr">{d.full}</span>
                          <a className="maplink" href={d.url} target="_blank" rel="noreferrer">View on map</a>
                        </>}
                        {d.kind === "unknown" && <span className="badge unknown">No address</span>}
                      </div>
                      <div className="what">
                        {Object.entries(p.cells).map(([c, q]) => <span className="chip" key={c}>{active?.headers[c]} × {q}</span>)}
                      </div>
                    </li>
                  );
                })}
              </ul>
            </section>
          </>
        )}

        {saved.length > 0 && (
          <section className="card">
            <h2>Saved choices</h2>
            <p className="sub">Products you matched to a column. Remove one to choose again.</p>
            <ul className="orders">
              {saved.map(([k, col]) => (
                <li className="order line" key={k}>
                  <span>{k} <span className="chip">{col}</span></span>
                  <button className="btn ghost" onClick={() => forget(k)}>Remove</button>
                </li>
              ))}
            </ul>
          </section>
        )}
      </main>

      {fileName && (
        <div className="bar">
          <div className="wrap">
            <p>{barText}</p>
            <button className="btn" onClick={download} disabled={busy || !toAdd || unmapped.length > 0}>
              {busy ? "Preparing…" : continuing ? "Download updated sheet" + (sheets.length > 1 ? "s" : "") : "Download filled sheet" + (sheets.length > 1 ? "s" : "")}
            </button>
          </div>
        </div>
      )}
    </>
  );
}
