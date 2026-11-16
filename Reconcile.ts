/**
 * Credit Line Utilization Report — Office Scripts
 * Workbook must contain sheets: FBL5n, VendorMaster, BUMaster, CreditLine.
 * All four source files must first be converted from .xls to .xlsx and
 * copied into these worksheets (e.g. by Power Automate / a staging step).
 * Amounts are in INR; credit line input is INR millions.
 * ASSUMPTION: open items have blank Clearing Doc, and posting date <= asOfDate.
 */
function main(workbook: ExcelScript.Workbook, asOfDate: string = "2026-09-30"): string {
  type Cell = string | number | boolean;
  type RecordRow = { [key: string]: Cell };
  const norm = (v: Cell | undefined): string => String(v ?? "").trim().replace(/\s+/g, " ").toUpperCase();
  const companyKey = (v: Cell | undefined): string => norm(v).replace(/[^A-Z0-9]/g, "");
  const amount = (v: Cell | undefined): number => {
    if (typeof v === "number") return v;
    const s = String(v ?? "").trim();
    if (!s) return 0;
    const negative = /^\(.*\)$/.test(s);
    const n = Number(s.replace(/[₹,()\s]/g, ""));
    if (!Number.isFinite(n)) throw new Error(`Invalid amount: ${s}`);
    return negative ? -n : n;
  };
  function readSheet(name: string): RecordRow[] {
    const sheet = workbook.getWorksheet(name);
    if (!sheet) throw new Error(`Missing worksheet: ${name}`);
    const range = sheet.getUsedRange();
    if (!range) throw new Error(`Empty worksheet: ${name}`);
    const cells = range.getValues() as Cell[][];
    const headers = cells[0].map(x => norm(x));
    return cells.slice(1).filter(row => row.some(x => norm(x) !== "")).map(row => {
      const obj: RecordRow = {};
      headers.forEach((h, i) => { if (h) obj[h] = row[i] ?? ""; });
      return obj;
    });
  }
  const field = (r: RecordRow, ...names: string[]): Cell => {
    for (const name of names) {
      const value = r[norm(name)];
      if (value !== undefined) return value;
    }
    return "";
  };
  const requireFields = (sheet: string, fields: string[]): void => {
    const ws = workbook.getWorksheet(sheet);
    if (!ws) throw new Error(`Missing worksheet: ${sheet}`);
    const used = ws.getUsedRange();
    if (!used) throw new Error(`Empty worksheet: ${sheet}`);
    const hdr = used.getRow(0).getValues()[0].map(x => norm(x));
    for (const f of fields) if (!hdr.includes(norm(f))) throw new Error(`${sheet}: missing column '${f}'`);
  };
  requireFields("FBL5n", ["Account", "Posting Date", "Amount in Local Currency", "Clearing Doc"]);
  requireFields("VendorMaster", ["Account", "Customer Name"]);
  requireFields("BUMaster", ["Account", "Branch Name"]);
  requireFields("CreditLine", ["Branch", "Company Name"]);
  const txns = readSheet("FBL5n");
  const vendors = readSheet("VendorMaster");
  const branches = readSheet("BUMaster");
  const limits = readSheet("CreditLine");
  const vendorMap = new Map<string, string>();
  vendors.forEach(r => vendorMap.set(norm(field(r,"Account")), String(field(r,"Customer Name")).trim()));
  const branchMap = new Map<string, string>();
  branches.forEach(r => branchMap.set(norm(field(r,"Account")), String(field(r,"Branch Name")).trim()));
  const reportDate = new Date(asOfDate + "T00:00:00Z");
  if (Number.isNaN(reportDate.getTime())) throw new Error("asOfDate must be YYYY-MM-DD");
  function dateValue(v: Cell): number {
    if (typeof v === "number") return Date.UTC(1899,11,30) + Math.floor(v) * 86400000;
    const s = String(v ?? "").trim();
    if (!s) throw new Error("Posting Date is blank");
    let m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
    if (m) return Date.UTC(Number(m[3]), Number(m[2])-1, Number(m[1]));
    const t = Date.parse(s);
    if (Number.isNaN(t)) throw new Error(`Invalid posting date: ${s}`);
    return t;
  }
  // Sum open balances by normalized Branch + Company Name.
  const utilization = new Map<string, number>();
  let missingMasters = 0;
  for (const t of txns) {
    if (norm(field(t,"Clearing Doc")) !== "") continue;
    if (dateValue(field(t,"Posting Date")) > reportDate.getTime()) continue;
    const acct = norm(field(t,"Account"));
    const customer = vendorMap.get(acct);
    const branch = branchMap.get(acct);
    if (!customer || !branch) { missingMasters++; continue; }
    const key = norm(branch) + "|" + companyKey(customer);
    utilization.set(key, (utilization.get(key) ?? 0) + amount(field(t,"Amount in Local Currency")));
  }
  const output: Cell[][] = [];
  const matched = new Set<string>();
  for (const r of limits) {
    const branch = String(field(r,"Branch")).trim();
    const company = String(field(r,"Company Name", "Company name")).trim();
    if (!branch || !company) continue;
    const creditMillion = amount(field(r,"Approved Credit Line Amt for FY", "Approved Credit Line Amount for FY26 (INR Million)", "Approved Credit Line Amount for FY26", "Approved Credit Line Amt for FY26"));
    const creditINR = creditMillion * 1000000;
    const key = norm(branch) + "|" + companyKey(company);
    if (matched.has(key)) throw new Error(`Duplicate CreditLine branch/company: ${branch} / ${company}`);
    matched.add(key);
    const used = utilization.get(key) ?? 0;
    output.push([branch.toUpperCase(), company.toUpperCase(), creditINR, String(field(r,"Credit Insurance Limit Validity")).trim(), used, creditINR > 0 ? used / creditINR * 100 : ""]);
  }
  output.sort((a,b) => String(a[0]).localeCompare(String(b[0])) || String(a[1]).localeCompare(String(b[1])));
  const old = workbook.getWorksheet("Credit Utilization Report");
  if (old) old.delete();
  const ws = workbook.addWorksheet("Credit Utilization Report");
  ws.getRange("A1:F1").merge();
  ws.getRange("A1").setValue("Credit Line Utilization Report");
  ws.getRange("A2:F2").merge();
  ws.getRange("A2").setValue(`As on ${asOfDate}`);
  ws.getRange("A4:F4").setValues([["Branch","Customer (Open Credit)","Credit Line","Validity of limit","Utilization","%"]]);
  const header = ws.getRange("A4:F4");
  header.getFormat().getFill().setColor("#DCE6EF");
  header.getFormat().getFont().setBold(true);
  header.getFormat().setRowHeight(32);
  if (output.length) {
    const data = ws.getRangeByIndexes(4,0,output.length,6);
    data.setValues(output);
    ws.getRange(`C5:C${output.length+4}`).setNumberFormatLocal("#,##0;[Red](#,##0);–");
    ws.getRange(`E5:E${output.length+4}`).setNumberFormatLocal("#,##0;[Red](#,##0);–");
    ws.getRange(`F5:F${output.length+4}`).setNumberFormatLocal("0.00");
    const pct = ws.getRange(`F5:F${output.length+4}`);
    const c1 = pct.addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
    c1.getCustom().getRule().setFormula('=AND(ISNUMBER(F5),F5>=0,F5<=20)');
    c1.getCustom().getFormat().getFill().setColor("#E9D38C");
    const c2 = pct.addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
    c2.getCustom().getRule().setFormula('=AND(ISNUMBER(F5),F5>20,F5<=70)');
    c2.getCustom().getFormat().getFill().setColor("#B6DAD1");
    const c3 = pct.addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
    c3.getCustom().getRule().setFormula('=AND(ISNUMBER(F5),F5>70)');
    c3.getCustom().getFormat().getFill().setColor("#E3AAC7");
  }
  ws.getRange("A:A").getFormat().setColumnWidth(90);
  ws.getRange("B:B").getFormat().setColumnWidth(350);
  ws.getRange("C:C").getFormat().setColumnWidth(130);
  ws.getRange("D:D").getFormat().setColumnWidth(180);
  ws.getRange("E:E").getFormat().setColumnWidth(130);
  ws.getRange("F:F").getFormat().setColumnWidth(85);
  ws.getRange("A1:F1").getFormat().getFont().setBold(true);
  ws.getRange("A1:F1").getFormat().getFont().setSize(15);
  ws.getRange("A4:F4").getFormat().setWrapText(true);
  ws.getFreezePanes().freezeRows(4);
  ws.getAutoFilter().apply(ws.getRange(`A4:F${Math.max(5,output.length+4)}`));
  return `Created ${output.length} customer credit line rows; ${missingMasters} open transactions skipped due to missing master data; ${utilization.size - [...matched].filter(k => utilization.has(k)).length} utilization keys had no credit-line match.`;
}

