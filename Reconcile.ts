/**
 * Run against Reconciliation_Working.xlsx.
 * Sheets Sales, Ledger, Bank must contain data in row 1 onward.
 * For the demo, the three input datasets have already been staged there.
 * Invoice IDs must be unique per source; duplicate IDs trigger an error.
 */
function main(workbook: ExcelScript.Workbook): string {
  const sales = readRows(workbook, "Sales", 3);
  const ledger = readRows(workbook, "Ledger", 2);
  const bank = readRows(workbook, "Bank", 2);
  const salesMap = new Map<string, (string | number | boolean)[]>();
  const ledgerMap = new Map<string, (string | number | boolean)[]>();
  const bankMap = new Map<string, (string | number | boolean)[]>();
  function addRows(rows: (string | number | boolean)[][], target: Map<string, (string | number | boolean)[]>, source: string) {
    for (const row of rows) {
      const id = String(row[0]).trim();
      if (!id) continue;
      if (target.has(id)) throw new Error(`Duplicate InvoiceID ${id} in ${source}`);
      target.set(id, row);
    }
  }
  addRows(sales, salesMap, "Sales");
  addRows(ledger, ledgerMap, "Ledger");
  addRows(bank, bankMap, "Bank");
  const keys = Array.from(new Set([...salesMap.keys(), ...ledgerMap.keys(), ...bankMap.keys()])).sort();
  const output: (string | number)[][] = [["InvoiceID", "Customer", "SalesAmount", "LedgerAmount", "BankAmount", "SalesVsLedger", "SalesVsBank", "Status"]];
  let matched = 0, exceptions = 0;
  for (const id of keys) {
    const s = salesMap.get(id), l = ledgerMap.get(id), b = bankMap.get(id);
    const sv = s ? amount(s[2], id, "Sales") : null;
    const lv = l ? amount(l[1], id, "Ledger") : null;
    const bv = b ? amount(b[1], id, "Bank") : null;
    let status: string;
    if (!s) status = "MISSING SALES";
    else if (!l) status = "MISSING LEDGER";
    else if (!b) status = "MISSING BANK";
    else if (Math.abs(sv! - lv!) < 0.005 && Math.abs(sv! - bv!) < 0.005) status = "MATCHED";
    else status = "VARIANCE";
    if (status === "MATCHED") matched++; else exceptions++;
    output.push([id, s ? String(s[1]) : "", sv ?? "", lv ?? "", bv ?? "", sv !== null && lv !== null ? round2(sv - lv) : "", sv !== null && bv !== null ? round2(sv - bv) : "", status]);
  }
  const report = workbook.getWorksheet("Reconciliation") ?? workbook.addWorksheet("Reconciliation");
  const existing = report.getUsedRange();
  if (existing) existing.clear(ExcelScript.ClearApplyTo.all);
  const target = report.getRangeByIndexes(0, 0, output.length, 8);
  target.setValues(output);
  report.getRange("A1:H1").getFormat().getFill().setColor("#D75F00");
  report.getRange("A1:H1").getFormat().getFont().setColor("#FFFFFF");
  report.getRange("A1:H1").getFormat().getFont().setBold(true);
  report.getRange(`C2:G${output.length}`).setNumberFormat("#,##0.00;[Red](#,##0.00)");
  report.getUsedRange().getFormat().autofitColumns();
  return `Reconciled ${keys.length} invoices: ${matched} matched, ${exceptions} exceptions`;
}
function readRows(workbook: ExcelScript.Workbook, sheetName: string, columns: number): (string | number | boolean)[][] {
  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) throw new Error(`Missing worksheet: ${sheetName}`);
  const used = sheet.getUsedRange();
  if (!used || used.getRowCount() < 2) return [];
  return sheet.getRangeByIndexes(1, 0, used.getRowCount() - 1, columns).getValues();
}
function amount(value: string | number | boolean, id: string, source: string): number {
  if (value === "" || value === null || typeof value === "boolean") throw new Error(`Invalid amount for ${id} in ${source}`);
  const n = Number(value);
  if (!Number.isFinite(n)) throw new Error(`Invalid amount for ${id} in ${source}`);
  return n;
}
function round2(n: number): number { return Math.round((n + Number.EPSILON) * 100) / 100; }
