"use client";
import { useEffect, useMemo, useState } from "react";
import { Template, Batch, AreaMode, readExport, groupOrders, readTemplate, place, paginate, fillTemplate, mergeBatch, norm } from "@/lib/greenfish";

const SAVED = "greenfish-extra-mapping";
const DAY = "greenfish-day-v1";
const today = () => new Date().toISOString().slice(0, 10);
const hhmm = (iso: string) => new Date(iso).toLocaleTimeString("en-ZA", { hour: "2-digit", minute: "2-digit" });

export default function Home() {
  const [tpl, setTpl] = useState<Template | null>(null);
  const [batches, setBatches] = useState<Batch[]>([]);
  const [date, setDate] = useState(today());
  const [ready, setReady] = useState(false);
  const [saved, setSaved] = useState<Record<string, string>>({});
  const [areaMode, setAreaMode] = useState<AreaMode>("zip");
  const [includeIgnored, setIncludeIgnored] = useState(false);
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [over, setOver] = useState(false);
  const [copied, setCopied] = useState("");

  useEffect(() => {
    fetch("/template.xlsx").then(r => r.arrayBuffer()).then(b => setTpl(readTemplate(b)))
      .catch(() => setError("The sheet template could not be loaded. Check that public/template.xlsx exists."));
    try { setSaved(JSON.parse(localStorage.getItem(SAVED) || "{}")); } catch {}
    try {
      const d = JSON.parse(localStorage.getItem(DAY) || "null");
      if (d) { setBatches(d.batches || []); setDate(d.date || today()); }
    } catch {}
    setReady(true);
  }, []);
  // The whole day is kept on this computer, so a refresh or a closed tab loses nothing.
  useEffect(() => { if (ready) localStorage.setItem(DAY, JSON.stringify({ date, batches })); }, [ready, date, batches]);

  const orders = useMemo(() => batches.flatMap(b => b.orders), [batches]);
  const batchOf = useMemo(() => { const m = new Map<string, Batch>(); batches.forEach(b => b.orders.forEach(o => m.set(o.order, b))); return m; }, [batches]);
  const placed = useMemo(() => (tpl ? place(orders, tpl, saved, areaMode, includeIgnored) : []), [orders, tpl, saved, areaMode, includeIgnored]);
  const pages = useMemo(() => paginate(placed), [placed]);
  const unlisted = useMemo(() => [...new Set(placed.flatMap(p => p.unlisted.map(l => l.title)))], [placed]);
  const options = useMemo(() => tpl ? ["IGNORE", ...Array.from(new Set(Object.values(tpl.headers))).filter(h => !["DELIVERED", "ORDER#", "AREA"].includes(h.toUpperCase()))] : [], [tpl]);
  const collections = placed.filter(p => p.order.collection).length;
  const staleDay = batches.length > 0 && date < today();

  async function load(f?: File) {
    if (!f) return;
    setError(""); setMessage("");
    try {
      const incoming = groupOrders(await readExport(f));
      if (!incoming.length) throw new Error("empty");
      const r = mergeBatch(batches, incoming, f.name);
      setBatches(r.batches);
      setMessage(r.added
        ? `${r.added} new order${r.added > 1 ? "s" : ""} added${r.skipped ? `. ${r.skipped} already on the sheet, left as they were.` : "."}`
        : "Nothing new: every order in that file is already on the sheet.");
    } catch { setError("That file could not be read. Use the orders export from Shopify (CSV or Excel)."); }
  }
  function newDay() {
    if (!confirm("Start a new day? This clears every order on the current sheet. Download it first if you still need it.")) return;
    setBatches([]); setDate(today()); setMessage("");
  }
  function persist(next: Record<string, string>) { setSaved(next); localStorage.setItem(SAVED, JSON.stringify(next)); }
  const assign = (title: string, col: string) => persist({ ...saved, [norm(title)]: col });
  const unassign = (k: string) => { const n = { ...saved }; delete n[k]; persist(n); };
  async function copy(url: string, id: string) {
    try { await navigator.clipboard.writeText(url); setCopied(id); setTimeout(() => setCopied(""), 1600); } catch {}
  }
  async function download() {
    if (!tpl) return;
    const [y, m, d] = date.split("-");
    for (let i = 0; i < pages.length; i++) {
      const bytes = await fillTemplate(tpl, pages[i], { date: `${d}/${m}/${y}`, pageNo: i + 1, pageCount: pages.length });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(new Blob([bytes as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      a.download = `Greenfish daily sales ${date}${pages.length > 1 ? ` (page ${i + 1})` : ""}.xlsx`;
      a.click();
    }
  }

  return (
    <>
      <header className="hero">
        <div className="hero-inner">
          <div className="brand">Greenfish Home</div>
          <h1>Daily delivery sheet</h1>
          <p>Upload each hourly Shopify export. New orders join the sheet, earlier ones stay exactly as they are.</p>
        </div>
      </header>

      <main className="wrap">
        <section className="panel">
          <label className={`drop${over ? " over" : ""}`}
            onDragOver={e => { e.preventDefault(); setOver(true); }} onDragLeave={() => setOver(false)}
            onDrop={e => { e.preventDefault(); setOver(false); load(e.dataTransfer.files[0]); }}>
            <input type="file" accept=".csv,.xlsx,.xls" onChange={e => { load(e.target.files?.[0]); e.target.value = ""; }} />
            <strong>{batches.length ? "Add the next export" : "Drop the Shopify orders export here"}</strong>
            <span>or click to choose a CSV or Excel file</span>
          </label>
          {message && <div className="ok" role="status">{message}</div>}
          {error && <div className="error" role="alert">{error}</div>}
          <div className="controls">
            <label className="field">Sheet date
              <input type="date" value={date} onChange={e => setDate(e.target.value)} />
            </label>
            <label className="field">Area column shows
              <select value={areaMode} onChange={e => setAreaMode(e.target.value as AreaMode)}>
                <option value="zip">Postal code</option>
                <option value="suburb">Suburb (address line 2)</option>
                <option value="city">City</option>
                <option value="street">Street address</option>
              </select>
            </label>
            <label className="check">
              <input type="checkbox" checked={includeIgnored} onChange={e => setIncludeIgnored(e.target.checked)} />
              Also note products marked IGNORE on the sheet as OTHER
            </label>
          </div>
        </section>

        {staleDay && (
          <div className="notice">
            <h3>This sheet is from {date.split("-").reverse().join("/")}</h3>
            <div className="row"><span>Download it if you still need it, then start today&apos;s sheet.</span>
              <button className="btn" onClick={newDay}>Start a new day</button></div>
          </div>
        )}

        {batches.length > 0 && (
          <>
            <div className="stats">
              <div className="stat"><b>{placed.length}</b><span>orders</span></div>
              <div className="stat"><b>{placed.length - collections}</b><span>deliveries</span></div>
              <div className="stat"><b>{collections}</b><span>collections</span></div>
              <div className="stat"><b>{pages.length}</b><span>{pages.length === 1 ? "sheet" : "sheets"} (29 orders each)</span></div>
            </div>

            <ol className="uploads" aria-label="Uploads today">
              {batches.map(b => (
                <li key={b.id}><span aria-hidden>🔒</span> {hhmm(b.at)} upload <span className="muted">{b.orders.length} order{b.orders.length > 1 ? "s" : ""}, {b.fileName}</span></li>
              ))}
            </ol>

            {unlisted.length > 0 && (
              <section className="notice">
                <h3>{unlisted.length} product{unlisted.length > 1 ? "s are" : " is"} not on your list</h3>
                <span>They still go on the sheet as &quot;OTHER&quot; text on that order&apos;s row. Choose a column to put them in the right place.</span>
                {unlisted.map(t => (
                  <div className="row" key={t}>
                    <span>{t}</span>
                    <select defaultValue="" aria-label={`Column for ${t}`} onChange={e => e.target.value && assign(t, e.target.value)}>
                      <option value="">Choose a column…</option>
                      {options.map(o => <option key={o}>{o}</option>)}
                    </select>
                  </div>
                ))}
              </section>
            )}

            <div className="orders">
              {placed.map(({ order: o, lines }) => (
                <article key={o.order} className={`card${o.collection ? " collect" : ""}`}>
                  <div className="card-top">
                    <h3>{o.order}</h3>
                    {o.customer && <span className="who">{o.customer}</span>}
                    <span className="who">🔒 {hhmm(batchOf.get(o.order)!.at)}</span>
                    <span className={`badge${o.collection ? " collect" : ""}`}>{o.collection ? "Collection" : "Delivery"}</span>
                  </div>
                  <div className="where">
                    {o.collection ? (
                      <b>Collection<small>{o.collectionReason === "no-address" ? "No street address in the export" : o.method}</small></b>
                    ) : (
                      <>
                        <b>{o.area}</b>
                        <a className="btn" href={o.mapsUrl} target="_blank" rel="noreferrer">Open in Maps</a>
                        <button className="btn" onClick={() => copy(o.mapsUrl, o.order)}>{copied === o.order ? "Copied" : "Copy link"}</button>
                      </>
                    )}
                  </div>
                  <div className="chips">
                    {lines.map((l, i) => l.how === "none"
                      ? <span className="chip other" key={i}>{l.title.split(" | ")[0]} <b>× {l.qty}</b> not on list</span>
                      : l.ignored
                        ? <span className="chip skip" key={i}>{l.title.split(" | ")[0]} <b>× {l.qty}</b> {includeIgnored ? "as OTHER" : "not on sheet"}</span>
                        : <span className="chip" key={i}>{l.header} <b>{l.entree ? `${l.qty}e` : `× ${l.qty}`}</b></span>)}
                  </div>
                </article>
              ))}
            </div>
          </>
        )}

        {Object.keys(saved).length > 0 && (
          <details className="saved">
            <summary>Saved column choices ({Object.keys(saved).length})</summary>
            <ul>{Object.entries(saved).map(([k, v]) => (
              <li key={k}>{k} → <b>{v}</b> <button className="btn link" onClick={() => unassign(k)}>Remove</button></li>
            ))}</ul>
          </details>
        )}
      </main>

      {batches.length > 0 && (
        <div className="bar">
          <div className="bar-inner">
            <span className="muted">{placed.length} orders for {date.split("-").reverse().join("/")}</span>
            <span style={{ display: "flex", gap: 10 }}>
              <button className="btn" onClick={newDay}>Start a new day</button>
              <button className="btn primary" onClick={download}>Download filled sheet{pages.length > 1 ? "s" : ""}</button>
            </span>
          </div>
        </div>
      )}
    </>
  );
}
