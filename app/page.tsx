"use client";
import { useEffect, useMemo, useState } from "react";
import { Template, Order, AreaMode, readExport, groupOrders, readTemplate, place, paginate, fillTemplate, norm } from "@/lib/greenfish";

const KEY = "greenfish-extra-mapping";
const today = () => new Date().toISOString().slice(0, 10);

export default function Home() {
  const [tpl, setTpl] = useState<Template | null>(null);
  const [orders, setOrders] = useState<Order[]>([]);
  const [saved, setSaved] = useState<Record<string, string>>({});
  const [date, setDate] = useState(today());
  const [areaMode, setAreaMode] = useState<AreaMode>("suburb");
  const [areaEdits, setAreaEdits] = useState<Record<string, string>>({});
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [over, setOver] = useState(false);
  const [copied, setCopied] = useState("");

  useEffect(() => {
    fetch("/template.xlsx").then(r => r.arrayBuffer()).then(b => setTpl(readTemplate(b)))
      .catch(() => setError("The sheet template could not be loaded. Check that public/template.xlsx exists."));
    try { setSaved(JSON.parse(localStorage.getItem(KEY) || "{}")); } catch {}
  }, []);

  const placed = useMemo(() => (tpl ? place(orders, tpl, saved, areaMode, areaEdits) : []), [orders, tpl, saved, areaMode, areaEdits]);
  const pages = useMemo(() => paginate(placed), [placed]);
  const others = useMemo(() => [...new Set(placed.flatMap(p => p.other.map(l => l.title)))], [placed]);
  const options = useMemo(() => tpl ? ["IGNORE", ...Array.from(new Set(Object.values(tpl.headers))).filter(h => !["DELIVERED", "ORDER#", "AREA"].includes(h.toUpperCase()))] : [], [tpl]);
  const collections = placed.filter(p => p.order.collection).length;

  async function load(f?: File) {
    if (!f) return;
    setError("");
    try {
      const o = groupOrders(await readExport(f));
      if (!o.length) throw new Error("empty");
      setOrders(o); setFileName(f.name); setAreaEdits({});
    } catch { setError("That file could not be read. Use the orders export from Shopify (CSV or Excel)."); }
  }
  function persist(next: Record<string, string>) { setSaved(next); localStorage.setItem(KEY, JSON.stringify(next)); }
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
          <p>Drop in today&apos;s Shopify orders and your sheet fills itself in, with a map link for every delivery.</p>
        </div>
      </header>

      <main className="wrap">
        <section className="panel">
          <label className={`drop${over ? " over" : ""}`}
            onDragOver={e => { e.preventDefault(); setOver(true); }} onDragLeave={() => setOver(false)}
            onDrop={e => { e.preventDefault(); setOver(false); load(e.dataTransfer.files[0]); }}>
            <input type="file" accept=".csv,.xlsx,.xls" onChange={e => load(e.target.files?.[0])} />
            <strong>{fileName ? "Choose a different export" : "Drop the Shopify orders export here"}</strong>
            <span>{fileName ? `Loaded: ${fileName}` : "or click to choose a CSV or Excel file"}</span>
          </label>
          <div className="controls">
            <label className="field">Sheet date
              <input type="date" value={date} onChange={e => setDate(e.target.value)} />
            </label>
            <label className="field">Area comes from
              <select value={areaMode} onChange={e => setAreaMode(e.target.value as AreaMode)}>
                <option value="suburb">Suburb (falls back to city)</option>
                <option value="city">City only</option>
              </select>
            </label>
          </div>
          {error && <div className="error" role="alert">{error}</div>}
        </section>

        {orders.length > 0 && (
          <>
            <div className="stats">
              <div className="stat"><b>{placed.length}</b><span>orders</span></div>
              <div className="stat"><b>{placed.length - collections}</b><span>deliveries</span></div>
              <div className="stat"><b>{collections}</b><span>collections</span></div>
              <div className="stat"><b>{pages.length}</b><span>{pages.length === 1 ? "sheet" : "sheets"} (29 orders each)</span></div>
            </div>

            {others.length > 0 && (
              <section className="notice">
                <h3>{others.length} product{others.length > 1 ? "s are" : " is"} not on your list</h3>
                <span>They will still appear on the sheet, written as &quot;OTHER&quot; text on the order&apos;s row. Choose a column to put them in the right place instead.</span>
                {others.map(t => (
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
              {placed.map(p => { const { order: o, lines } = p; return (
                <article key={o.order} className={`card${o.collection ? " collect" : ""}`}>
                  <div className="card-top">
                    <h3>{o.order}</h3>
                    {o.customer && <span className="who">{o.customer}</span>}
                    <span className={`badge${o.collection ? " collect" : ""}`}>{o.collection ? "Collection" : "Delivery"}</span>
                  </div>
                  <div className="where">
                    {o.collection ? (
                      <b>Collection<small>{o.collectionReason === "no-address" ? "No street address in the export" : o.method}</small></b>
                    ) : (
                      <>
                        <input className="area" aria-label={`Area for ${o.order}`} value={o.area} placeholder="Type the area"
                          onChange={e => setAreaEdits(a => ({ ...a, [o.order]: e.target.value }))} />
                        <a className="btn" href={o.mapsUrl} target="_blank" rel="noreferrer">Open in Maps</a>
                        <button className="btn" onClick={() => copy(o.mapsUrl, o.order)}>{copied === o.order ? "Copied" : "Copy link"}</button>
                      </>
                    )}
                  </div>
                  {p.otherCut > 0 && <div className="warn">{p.otherCut} unlisted product(s) did not fit on this row of the sheet and print as “+{p.otherCut} more”. Check the list below.</div>}
                  <div className="chips">
                    {lines.map((l, i) => l.how === "none"
                      ? <span className="chip other" key={i}>{l.title.split(" | ")[0]} <b>× {l.qty}</b> not on list</span>
                      : l.ignored
                        ? <span className="chip skip" key={i}>{l.title.split(" | ")[0]} not on sheet</span>
                        : <span className="chip" key={i}>{l.header} <b>{l.entree ? `${l.qty}e` : `× ${l.qty}`}</b></span>)}
                  </div>
                </article>
              ); })}
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

      {orders.length > 0 && (
        <div className="bar">
          <div className="bar-inner">
            <span className="muted">{`${placed.length} orders ready for ${date.split("-").reverse().join("/")}`}{others.length ? `, ${others.length} product(s) will print as OTHER` : ""}</span>
            <button className="btn primary" onClick={download} disabled={!placed.length}>Download filled sheet{pages.length > 1 ? "s" : ""}</button>
          </div>
        </div>
      )}
    </>
  );
}
