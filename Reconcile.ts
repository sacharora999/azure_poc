/**
 * Credit Line Utilization Report
 * Input worksheets: FBL5n, VendorMaster, BUMaster, CreditLine
 * Output worksheet: Credit Utilization Report
 *
 * Business-rule assumptions to confirm:
 * - Open items: Clearing Doc blank (no clearing-date field supplied).
 * - Posting Date <= asOfDate.
 * - Utilization: sum of Amount in Local Currency, preserving source sign.
 * - Credit line values are in INR millions; multiply by 1,000,000.
 */
function main(workbook: ExcelScript.Workbook, asOfDate: string = "2026-09-30"): string {
  type Cell = string | number | boolean;
  type Row = Cell[];
  const clean = (v: Cell): string => String(v ?? "").trim();
  const key = (v: Cell): string => clean(v).replace(/\s+/g, " ").toUpperCase();
  const read = (name: string): Row[] => {
    const sheet = workbook.getWorksheet(name);
    if (!sheet) throw new Error("Missing worksheet: " + name);
    const range = sheet.getUsedRange();
    if (!range) throw new Error("Empty worksheet: " + name);
    const rows = range.getValues();
    if (rows.length < 2) throw new Error("No data rows in: " + name);
    return rows;
  };
  const index = (headers: Row, aliases: string[], required: boolean = true): number => {
    const normalized = headers.map(v => key(v).replace(/[^A-Z0-9]/g, ""));
    for (const alias of aliases) {
      const found = normalized.indexOf(key(alias).replace(/[^A-Z0-9]/g, ""));
      if (found >= 0) return found;
    }
    if (required) throw new Error("Missing column: " + aliases[0] + ". Available: " + headers.map(clean).join(", "));
    return -1;
  };
  const amount = (v: Cell): number => {
    if (typeof v === "number") return Number.isFinite(v) ? v : 0;
    let s = clean(v).replace(/[,₹\s]/g, "");
    if (!s || s === "-") return 0;
    if (/^\(.*\)$/.test(s)) s = "-" + s.slice(1, -1);
    const n = Number(s);
    if (!Number.isFinite(n)) throw new Error("Invalid numeric amount: " + clean(v));
    return n;
  };
  const excelDate = (v: Cell): number | null => {
    const s = clean(v);
    if (!s) return null;
    if (typeof v === "number" || /^\d{5}(\.\d+)?$/.test(s)) {
      const serial = Number(v);
      return Number.isFinite(serial) ? Math.floor(serial) : null;
    }
    let year = 0, month = 0, day = 0;
    let m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s);
    if (m) { year = Number(m[1]); month = Number(m[2]); day = Number(m[3]); }
    else {
      m = /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/.exec(s);
      if (!m) return null;
      day = Number(m[1]); month = Number(m[2]); year = Number(m[3]);
    }
    const utc = Date.UTC(year, month - 1, day);
    const dt = new Date(utc);
    if (dt.getUTCFullYear() !== year || dt.getUTCMonth() !== month - 1 || dt.getUTCDate() !== day) return null;
    return Math.floor((utc - Date.UTC(1899, 11, 30)) / 86400000);
  };
  const cutoff = excelDate(asOfDate);
  if (cutoff === null) throw new Error("Invalid asOfDate; use YYYY-MM-DD: " + asOfDate);

  const vendor = read("VendorMaster");
  const va = index(vendor[0], ["Account"]);
  const vn = index(vendor[0], ["Customer Name", "Company Name"]);
  const customers = new Map<string, string>();
  for (const r of vendor.slice(1)) if (key(r[va])) customers.set(key(r[va]), clean(r[vn]));

  const bu = read("BUMaster");
  const ba = index(bu[0], ["Account"]);
  const bn = index(bu[0], ["Branch Name", "Branch"]);
  const branches = new Map<string, string>();
  for (const r of bu.slice(1)) if (key(r[ba])) branches.set(key(r[ba]), clean(r[bn]));

  const credit = read("CreditLine");
  const cb = index(credit[0], ["Branch"]);
  const cn = index(credit[0], ["Company Name", "Company name", "Customer Name"]);
  const cl = index(credit[0], ["Approved Credit Line Amt for FY", "Approved Credit line amount for FY26 (in INR Million)", "Approved Credit Line Amount"]);
  const cv = index(credit[0], ["Credit Insurance Limit Validity", "Credit Insurance Limit validity"], false);
  type Credit = { branch: string; customer: string; limit: number; validity: string };
  const credits = new Map<string, Credit>();
  for (const r of credit.slice(1)) {
    const branch = clean(r[cb]); const customer = clean(r[cn]);
    if (!branch || !customer) continue;
    const composite = key(branch) + "|" + key(customer);
    if (credits.has(composite)) throw new Error("Duplicate CreditLine branch/company: " + composite);
    credits.set(composite, { branch, customer, limit: amount(r[cl]) * 1000000, validity: cv < 0 ? "" : clean(r[cv]) });
  }

  const fbl = read("FBL5n");
  const fa = index(fbl[0], ["Account"]);
  const fp = index(fbl[0], ["Posting Date"]);
  const fc = index(fbl[0], ["Clearing Doc", "Clearing Document"]);
  const fm = index(fbl[0], ["Amount in Local Currency"]);
  const sums = new Map<string, number>();
  const missing = new Set<string>();
  let included = 0;
  for (const r of fbl.slice(1)) {
    const account = key(r[fa]);
    if (!account) continue;
    if (clean(r[fc]) !== "") continue;
    const posted = excelDate(r[fp]);
    if (posted === null) throw new Error("Invalid Posting Date for account " + account + ": " + clean(r[fp]));
    if (posted > cutoff) continue;
    const customer = customers.get(account);
    const branch = branches.get(account);
    if (!customer || !branch) { missing.add(account); continue; }
    const composite = key(branch) + "|" + key(customer);
    sums.set(composite, (sums.get(composite) ?? 0) + amount(r[fm]));
    included++;
  }
  if (missing.size) throw new Error("Missing master mappings for accounts: " + Array.from(missing).slice(0, 15).join(", "));

  let report = workbook.getWorksheet("Credit Utilization Report");
  if (!report) report = workbook.addWorksheet("Credit Utilization Report");
  const old = report.getUsedRange();
  if (old) old.clear(ExcelScript.ClearApplyTo.all);

  const header: Cell[] = ["Branch", "Customer (Open Credit)", "Credit Line", "Validity of limit", "Utilization", "%"];
  const output: Cell[][] = [header];
  const ordered = Array.from(credits.values()).sort((a, b) => (a.branch + "|" + a.customer).localeCompare(b.branch + "|" + b.customer));
  for (const c of ordered) {
    const utilization = sums.get(key(c.branch) + "|" + key(c.customer)) ?? 0;
    const percent = c.limit > 0 ? utilization / c.limit * 100 : 0;
    output.push([c.branch, c.customer, c.limit, c.validity, utilization, percent]);
  }
  if (output.length < 1 || output.length > 1000000) throw new Error("Invalid output row count: " + output.length);
  const width = 6;
  for (let i = 0; i < output.length; i++) if (output[i].length !== width) throw new Error("Invalid output width at row " + i);
  report.getRangeByIndexes(0, 0, output.length, width).setValues(output);

  const head = report.getRange("A1:F1");
  head.getFormat().getFill().setColor("#D9E2F3");
  head.getFormat().getFont().setBold(true);
  head.getFormat().setWrapText(true);
  report.getRange("A:A").getFormat().setColumnWidth(95);
  report.getRange("B:B").getFormat().setColumnWidth(340);
  report.getRange("C:E").getFormat().setColumnWidth(135);
  report.getRange("F:F").getFormat().setColumnWidth(90);
  if (output.length > 1) {
    const last = output.length;
    report.getRange(`C2:C${last}`).setNumberFormat("#,##0.00");
    report.getRange(`E2:E${last}`).setNumberFormat("#,##0.00");
    report.getRange(`F2:F${last}`).setNumberFormat("0.00");
    const percentRange = report.getRange(`F2:F${last}`);
    const low = percentRange.addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue);
    low.getCellValue().setRule({ formula1: "0", formula2: "20", operator: ExcelScript.ConditionalCellValueOperator.between });
    low.getCellValue().getFormat().getFill().setColor("#F4B183");
    const mid = percentRange.addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue);
    mid.getCellValue().setRule({ formula1: "20", formula2: "70", operator: ExcelScript.ConditionalCellValueOperator.between });
    mid.getCellValue().getFormat().getFill().setColor("#B7DEE8");
    const high = percentRange.addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue);
    high.getCellValue().setRule({ formula1: "70", operator: ExcelScript.ConditionalCellValueOperator.greaterThan });
    high.getCellValue().getFormat().getFill().setColor("#DAA0C1");
  }
  report.getRange("A1:F1").getFormat().setRowHeight(32);
  return `Report generated: ${ordered.length} credit-line records, ${included} eligible FBL5n rows, as of ${asOfDate}.`;
}
