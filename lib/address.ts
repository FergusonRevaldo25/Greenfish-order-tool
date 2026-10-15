export type Row = Record<string, string>;

export type Destination =
  | { kind: "collection"; label: string }
  | { kind: "delivery"; label: string; full: string; url: string }
  | { kind: "unknown"; label: string };

// Shopify prefixes postal codes with an apostrophe in CSV exports ('7600) so Excel keeps them as text.
const clean = (s: string | undefined) => (s || "").replace(/\s+/g, " ").replace(/^'+/, "").trim();

/** Shopify uses a pickup shipping method for collections, and these orders usually have no street address. */
// "Greenfish Factory" is how Greenfish's own collections appear in the shipping method column.
const COLLECTION_WORDS = /collect|pick[\s-]?up|in[\s-]?store|greenfish factory/i;
const WALK_IN = /walk[\s-]?in/i;

/** First non-empty value among several possible Shopify column names (exports differ between versions). */
const pick = (r: Row, names: string[]) => {
  for (const n of names) if (clean(r[n])) return clean(r[n]);
  return "";
};

export function isCollection(r: Row): boolean {
  const method = pick(r, ["Shipping Method", "Delivery Method", "Delivery method"]);
  if (COLLECTION_WORDS.test(method)) return true;
  // Counter sales entered by hand in Shopify are named "walk-in" and have no address.
  if (WALK_IN.test(pick(r, ["Shipping Name", "Billing Name"])) && !pick(r, ["Shipping Address1", "Shipping Street"])) return true;
  const note = pick(r, ["Notes", "Note"]);
  return COLLECTION_WORDS.test(note) && !pick(r, ["Shipping Address1", "Shipping Street"]);
}

/** Builds the one-line address from whichever shipping columns the export has. */
export function fullAddress(r: Row): { street: string; suburb: string; full: string } {
  const a1 = pick(r, ["Shipping Address1", "Shipping Address 1"]);
  const a2 = pick(r, ["Shipping Address2", "Shipping Address 2"]);
  // Newer exports have one combined "Shipping Street" column; use it only when Address1 is missing.
  const street = a1 ? [a1, a2].filter(Boolean).join(", ") : pick(r, ["Shipping Street"]);
  const city = pick(r, ["Shipping City"]);
  const zip = pick(r, ["Shipping Zip", "Shipping Postal Code"]);
  const province = pick(r, ["Shipping Province Name", "Shipping Province"]);
  const cc = pick(r, ["Shipping Country", "Shipping Country Code"]);
  const country = cc.toUpperCase() === "ZA" ? "South Africa" : cc;
  const parts = [street, city, province, zip, country].filter(Boolean);
  // drop exact repeats (some stores type the city into both address2 and city)
  const seen = new Set<string>();
  const uniq = parts.filter(p => { const k = p.toLowerCase(); if (seen.has(k)) return false; seen.add(k); return true; });
  return { street, suburb: a2 || city, full: uniq.join(", ") };
}

/** Google Maps search link. Spaces become "+" so the link stays short enough for Excel's 255-character limit. */
export function mapsUrl(address: string): string {
  const q = encodeURIComponent(address).replace(/%20/g, "+");
  return `https://www.google.com/maps/search/?api=1&query=${q}`;
}

export function destinationFor(r: Row): Destination {
  if (isCollection(r)) return { kind: "collection", label: "COLLECTION" };
  const { street, suburb, full } = fullAddress(r);
  if (!full) return { kind: "unknown", label: "" };
  // What is printed in the AREA cell: the street and suburb, short enough to read on the sheet.
  const city = pick(r, ["Shipping City"]);
  const hasCity = city && street.toLowerCase().includes(city.toLowerCase());
  const label = [street, hasCity ? "" : suburb === street ? "" : city].filter(Boolean).join(", ") || full;
  let url = mapsUrl(full);
  if (url.length > 240) url = mapsUrl([street, pick(r, ["Shipping City"])].filter(Boolean).join(", "));
  if (url.length > 240) url = mapsUrl(pick(r, ["Shipping City"]) || full.slice(0, 80));
  return { kind: "delivery", label, full, url };
}
