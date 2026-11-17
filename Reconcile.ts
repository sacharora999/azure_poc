/**
 * DEBTORS AGEING REPORT — OFFICE SCRIPTS
 *
 * Required worksheets in the same Excel workbook:
 *   1. FBL5n        — Transaction amounts, due dates and Special G/L indicator
 *   2. VendorMaster — Customer names and countries, linked by Account
 *   3. BUMaster     — Branch names, linked by Account
 *   4. CreditLine   — MOJ SBU, linked by Branch and Company Name
 *
 * reportDate must be in YYYY-MM-DD format, e.g. 2026-09-30.
 * All report amounts are displayed in thousands of INR.
 * Transactions with Special G/L Ind = "T" are excluded.
 */
function main(workbook: ExcelScript.Workbook, reportDate: string = "2026-09-30") {
  // Cell values returned by Office Scripts.
  type Cell = string | number | boolean;

  // One consolidated customer record per branch.
  // amounts indexes: 0=Outstanding, 1=Before Due, 2=1-15, 3=16-30,
  // 4=31-60, 5=61-90, 6=Over 90 days.
  type Bucket = {
    code: string;
    sbu: string;
    customer: string;
    country: string;
    branch: string;
    amounts: number[];
  };

  // Convert a cell to trimmed text; normalize text for reliable joins.
  const clean = (v: Cell | undefined): string => String(v ?? "").trim();
  const norm = (v: Cell | undefined): string =>
    clean(v).replace(/\s+/g, " ").toUpperCase();

  // Convert Excel amount cells to numbers, including negatives in parentheses.
  // Empty or nonnumeric values are treated as zero.
  const numberValue = (v: Cell | undefined): number => {
    if (typeof v === "number") return v;
    const s = clean(v).replace(/,/g, "");
    const n = Number(
      s.startsWith("(") && s.endsWith(")")
        ? "-" + s.slice(1, -1)
        : s
    );
    return Number.isFinite(n) ? n : 0;
  };

  // Convert Net Due Date to an Excel date serial (whole days).
  // Numeric Excel dates are already serials; text dates are parsed as dates.
  const dateSerial = (v: Cell | undefined): number => {
    if (typeof v === "number") return Math.floor(v);
    const s = clean(v);
    if (!s) throw new Error("Missing Net Due Date");
    const d = new Date(s);
    if (isNaN(d.getTime())) throw new Error("Invalid Net Due Date: " + s);
    return Math.floor(
      (Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()) -
        Date.UTC(1899, 11, 30)) / 86400000
    );
  };

  // Validate the requested as-on date and convert it to an Excel date serial.
  const dateMatch = /^(\d{4})-(\d{2})-(\d{2})$/.exec(reportDate);
  if (!dateMatch) throw new Error("Report date must be YYYY-MM-DD");

  const y = Number(dateMatch[1]);
  const m = Number(dateMatch[2]);
  const d = Number(dateMatch[3]);
  const utc = Date.UTC(y, m - 1, d);
  const check = new Date(utc);

  if (
    check.getUTCFullYear() !== y ||
    check.getUTCMonth() !== m - 1 ||
    check.getUTCDate() !== d
  ) {
    throw new Error("Invalid report date");
  }

  const asOn = Math.floor((utc - Date.UTC(1899, 11, 30)) / 86400000);
  const formattedDate = `${dateMatch[3]}.${dateMatch[2]}.${dateMatch[1]}`;

  // Read a source worksheet, check required headers, and return its data rows.
  // Header comparisons ignore case and extra spaces.
  const read = (
    name: string,
    required: string[]
  ): { headers: string[]; rows: Cell[][] } => {
    const ws = workbook.getWorksheet(name);
    if (!ws) throw new Error(`Missing worksheet: ${name}`);

    const range = ws.getUsedRange();
    if (!range) throw new Error(`Empty worksheet: ${name}`);

    const values = range.getValues() as Cell[][];
    const headers = values[0].map(v => norm(v));

    required.forEach(h => {
      if (!headers.includes(norm(h))) {
        throw new Error(`${name}: missing column ${h}`);
      }
    });

    return { headers, rows: values.slice(1) };
  };

  // Retrieve a row value by column name instead of hardcoding column positions.
  const at = (
    data: { headers: string[] },
    row: Cell[],
    header: string
  ): Cell | undefined => row[data.headers.indexOf(norm(header))];

  // Load all four input datasets.
  const fbl = read("FBL5n", [
    "Account", "Special G/L Ind", "Net Due Date", "Amount in Local Currency"
  ]);
  const vendors = read("VendorMaster", ["Account", "Customer Name", "Country"]);
  const bu = read("BUMaster", ["Account", "Branch Name"]);
  const credit = read("CreditLine", ["Branch", "Company Name", "MOJ SBU"]);

  // Build Account -> customer name/country lookup from VendorMaster.
  const vendorByAccount = new Map<string, { customer: string; country: string }>();
  vendors.rows.forEach(r => {
    const key = norm(at(vendors, r, "Account"));
    if (key) {
      vendorByAccount.set(key, {
        customer: clean(at(vendors, r, "Customer Name")),
        country: clean(at(vendors, r, "Country"))
      });
    }
  });

  // Build Account -> normalized branch lookup from BUMaster.
  const branchByAccount = new Map<string, string>();
  bu.rows.forEach(r => {
    const key = norm(at(bu, r, "Account"));
    if (key) branchByAccount.set(key, norm(at(bu, r, "Branch Name")));
  });

  // Build Branch + Company Name -> MOJ SBU lookup from CreditLine.
  const sbuByBranchCustomer = new Map<string, string>();
  credit.rows.forEach(r => {
    const key =
      norm(at(credit, r, "Branch")) + "|" +
      norm(at(credit, r, "Company Name"));
    sbuByBranchCustomer.set(key, clean(at(credit, r, "MOJ SBU")));
  });

  // Group eligible transactions by Branch + Account.
  const grouped = new Map<string, Bucket>();
  fbl.rows.forEach(r => {
    const code = clean(at(fbl, r, "Account"));

    // Exclude blank Account and Special G/L indicator T.
    if (!code || norm(at(fbl, r, "Special G/L Ind")) === "T") return;

    // Enrich each transaction using the three master-data lookups.
    const vendor = vendorByAccount.get(norm(code));
    const branch = branchByAccount.get(norm(code)) ?? "UNMAPPED";
    const customer = vendor?.customer ?? "";
    const sbu = sbuByBranchCustomer.get(branch + "|" + norm(customer)) ?? "";
    const key = branch + "|" + norm(code);

    // Create the customer record the first time it appears.
    let item = grouped.get(key);
    if (!item) {
      item = {
        code,
        sbu,
        customer,
        country: vendor?.country ?? "",
        branch,
        amounts: [0, 0, 0, 0, 0, 0, 0]
      };
      grouped.set(key, item);
    }

    // Convert INR to INR thousands and calculate days past due.
    const amount = numberValue(at(fbl, r, "Amount in Local Currency")) / 1000;
    const days = asOn - dateSerial(at(fbl, r, "Net Due Date"));

    // Outstanding includes every eligible transaction amount.
    item.amounts[0] += amount;

    // Assign amount to exactly one aging bucket.
    // Day 0 is included in Before Due to reconcile with Outstanding.
    const bucket =
      days <= 0 ? 1 :
      days <= 15 ? 2 :
      days <= 30 ? 3 :
      days <= 60 ? 4 :
      days <= 90 ? 5 : 6;

    item.amounts[bucket] += amount;
  });

  // Replace any previous report with a fresh output worksheet.
  const existing = workbook.getWorksheet("Debtor Aging Report");
  if (existing) existing.delete();
  const ws = workbook.addWorksheet("Debtor Aging Report");

  // Define the 12 report columns in screenshot order.
  const headers = [
    "Customer Code", "SBU", "Customer Name", "Country",
    "Outstanding", "Before Due", "Up to 15 days", "16-30 days",
    "31-60 days", "61-90 days", "Over 90 days", "Total Overdue"
  ];

  // Create the merged title, blue header background, and units label.
  ws.getRange("A1:L2").merge();
  const title = ws.getRange("A1");
  title.setValue(`DEBTORS AGEING REPORT AS ON ${formattedDate}`);
  title.getFormat().getFont().setBold(true);
  title.getFormat().getFont().setSize(16);
  title.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  ws.getRange("A1:L3").getFormat().getFill().setColor("#D7EAF3");

  ws.getRange("L3").setValue("INR'1000");
  ws.getRange("L3").getFormat().getFont().setBold(true);
  ws.getRange("L3").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);

  // Create and format the column headings in row 4.
  ws.getRange("A4:L4").setValues([headers]);
  ws.getRange("A4:L4").getFormat().getFill().setColor("#B7D9EA");
  ws.getRange("A4:L4").getFormat().getFont().setBold(true);
  ws.getRange("A4:L4").getFormat().setRowHeight(34);
  ws.getRange("A4:L4").getFormat().setWrapText(true);

  // Create one section per branch, sorted alphabetically.
  const branchNames = Array.from(
    new Set(Array.from(grouped.values()).map(x => x.branch))
  ).sort();
  let rowNumber = 5;

  for (const branch of branchNames) {
    // Display the branch section heading.
    const branchRow = ws.getRange(`A${rowNumber}:L${rowNumber}`);
    branchRow.merge();
    ws.getRange(`A${rowNumber}`).setValue(
      "Branch: " + (
        branch === "DELHI" ? "New Delhi" :
        branch === "UNMAPPED" ? "Unmapped" :
        branch.charAt(0) + branch.slice(1).toLowerCase()
      )
    );
    branchRow.getFormat().getFill().setColor("#D7EAF3");
    branchRow.getFormat().getFont().setBold(true);
    rowNumber++;

    // Sort the branch's customer records by customer name.
    const records = Array.from(grouped.values())
      .filter(x => x.branch === branch)
      .sort((a, b) => a.customer.localeCompare(b.customer));
    const start = rowNumber;

    // Construct report rows. Total Overdue excludes Before Due and Outstanding.
    const output: (string | number)[][] = records.map(x => {
      const a = x.amounts;
      return [
        x.code, x.sbu, x.customer, x.country,
        ...a,
        a.slice(2).reduce((sum, v) => sum + v, 0)
      ];
    });

    // Write customer rows in a single operation and format the amount columns.
    if (output.length) {
      ws.getRangeByIndexes(rowNumber - 1, 0, output.length, 12).setValues(output);
      ws.getRange(`E${start}:L${start + output.length - 1}`)
        .setNumberFormat('#,##0;[Red]\\(#,##0);–');
      ws.getRange(`L${start}:L${start + output.length - 1}`)
        .getFormat().getFill().setColor("#F2DDD4");
      rowNumber += output.length;
    }

    // Calculate branch totals for Outstanding, all buckets, and Total Overdue.
    const totals = new Array<number>(8).fill(0);
    records.forEach(x => {
      x.amounts.forEach((v, i) => totals[i] += v);
      totals[7] += x.amounts.slice(2).reduce((sum, v) => sum + v, 0);
    });

    // Display a bold, blue branch-total row.
    ws.getRange(`A${rowNumber}:D${rowNumber}`).merge();
    ws.getRange(`A${rowNumber}`).setValue("Branch Total");
    ws.getRange(`E${rowNumber}:L${rowNumber}`).setValues([totals]);
    ws.getRange(`E${rowNumber}:L${rowNumber}`)
      .setNumberFormat('#,##0;[Red]\\(#,##0);–');
    ws.getRange(`A${rowNumber}:L${rowNumber}`)
      .getFormat().getFill().setColor("#D7EAF3");
    ws.getRange(`A${rowNumber}:L${rowNumber}`)
      .getFormat().getFont().setBold(true);

    // Leave one blank row before the next branch.
    rowNumber += 2;
  }

  // Set column widths for readability and keep headers visible while scrolling.
  ws.getRange("A:A").getFormat().setColumnWidth(115);
  ws.getRange("B:B").getFormat().setColumnWidth(55);
  ws.getRange("C:C").getFormat().setColumnWidth(315);
  ws.getRange("D:D").getFormat().setColumnWidth(65);
  ws.getRange("E:L").getFormat().setColumnWidth(100);
  ws.getFreezePanes().freezeRows(4);

  // Return a concise execution summary to the caller / Power Automate.
  return {
    reportDate,
    customers: grouped.size,
    branches: branchNames.length,
    excludedSpecialGL: "T",
    outputSheet: ws.getName()
  };
}
