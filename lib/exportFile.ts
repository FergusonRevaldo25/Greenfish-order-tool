import * as XLSX from "xlsx";
import Papa from "papaparse";
import type { Row } from "./address";

/** Reads a Shopify orders export (CSV or Excel) into rows keyed by column name. */
export async function readExport(file: File): Promise<Row[]> {
  if (/\.csv$/i.test(file.name)) {
    const text = await file.text();
    return Papa.parse<Row>(text, { header: true, skipEmptyLines: true }).data;
  }
  const wb = XLSX.read(await file.arrayBuffer());
  return XLSX.utils.sheet_to_json<Row>(wb.Sheets[wb.SheetNames[0]], { defval: "", raw: false });
}
