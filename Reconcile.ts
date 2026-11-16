
/**
 * Credit Line Utilization Report
 *
 * Input worksheets:
 * 1. FBL5n
 * 2. VendorMaster
 * 3. BUMaster
 * 4. CreditLine
 *
 * Output worksheet:
 * Credit Utilization Report
 *
 * Fixed column positions.
 * No dynamic header lookup.
 *
 * Credit Line is in INR millions.
 * Utilization is in INR.
 */

function main(workbook: ExcelScript.Workbook): string {

  // ============================================
  // 1. DEFINE DATA TYPES
  // ============================================

  type Cell = string | number | boolean;
  type Row = Cell[];

  // ============================================
  // 2. HELPER FUNCTIONS
  // ============================================

  // Remove leading and trailing spaces.
  const clean = (value: Cell): string => {
    return String(value ?? "").trim();
  };

  // Normalize values for matching.
  const key = (value: Cell): string => {
    return clean(value)
      .replace(/\s+/g, " ")
      .toUpperCase();
  };

  // Convert Excel values into numbers.
  const toNumber = (value: Cell): number => {

    if (typeof value === "number") {
      if (!Number.isFinite(value)) {
        throw new Error("Invalid numeric value");
      }
      return value;
    }

    let text = clean(value)
      .replace(/[,\s]/g, "");

    if (text === "" || text === "-") {
      return 0;
    }

    // Handle negative amounts such as (5000).
    if (/^\(.*\)$/.test(text)) {
      text = "-" + text.slice(1, -1);
    }

    const result = Number(text);

    if (!Number.isFinite(result)) {
      throw new Error(
        "Invalid numeric amount: " + clean(value)
      );
    }

    return result;
  };

  // Read worksheet data.
  const readSheet = (sheetName: string): Row[] => {

    const sheet = workbook.getWorksheet(sheetName);

    if (!sheet) {
      throw new Error(
        "Worksheet not found: " + sheetName
      );
    }

    const usedRange = sheet.getUsedRange();

    if (!usedRange) {
      throw new Error(
        "Worksheet is empty: " + sheetName
      );
    }

    const rows = usedRange.getValues();

    if (rows.length < 2) {
      throw new Error(
        "No data records in: " + sheetName
      );
    }

    return rows;
  };

  // ============================================
  // 3. FIXED COLUMN INDEXES
  // ============================================

  // JavaScript indexes start from zero.

  // FBL5n
  const FBL_ACCOUNT = 0;   // Column A
  const FBL_AMOUNT = 9;    // Column J

  // VendorMaster
  const VENDOR_ACCOUNT = 0; // Column A
  const VENDOR_NAME = 1;    // Column B

  // BUMaster
  const BU_ACCOUNT = 0;  // Column A
  const BU_BRANCH = 1;   // Column B

  // CreditLine
  const CREDIT_BRANCH = 0;  // Column A
  const CREDIT_COMPANY = 2; // Column C
  const CREDIT_AMOUNT = 3;  // Column D

  // ============================================
  // 4. READ INPUT WORKSHEETS
  // ============================================

  const fblRows = readSheet("FBL5n");
  const vendorRows = readSheet("VendorMaster");
  const buRows = readSheet("BUMaster");
  const creditRows = readSheet("CreditLine");

  // Check that required columns exist.
  if (fblRows[0].length <= FBL_AMOUNT) {
    throw new Error("FBL5n must contain column J");
  }

  if (vendorRows[0].length <= VENDOR_NAME) {
    throw new Error("VendorMaster must contain column B");
  }

  if (buRows[0].length <= BU_BRANCH) {
    throw new Error("BUMaster must contain column B");
  }

  if (creditRows[0].length <= CREDIT_AMOUNT) {
    throw new Error("CreditLine must contain column D");
  }

  // ============================================
  // 5. BUILD ACCOUNT -> CUSTOMER LOOKUP
  // ============================================

  const customerByAccount = new Map<string, string>();

  for (const row of vendorRows.slice(1)) {

    const account = key(row[VENDOR_ACCOUNT]);
    const customer = clean(row[VENDOR_NAME]);

    if (!account) {
      continue;
    }

    if (customerByAccount.has(account)) {
      throw new Error(
        "Duplicate Account in VendorMaster: " + account
      );
    }

    customerByAccount.set(account, customer);
  }

  // ============================================
  // 6. BUILD ACCOUNT -> BRANCH LOOKUP
  // ============================================

  const branchByAccount = new Map<string, string>();

  for (const row of buRows.slice(1)) {

    const account = key(row[BU_ACCOUNT]);
    const branch = clean(row[BU_BRANCH]);

    if (!account) {
      continue;
    }

    if (branchByAccount.has(account)) {
      throw new Error(
        "Duplicate Account in BUMaster: " + account
      );
    }

    branchByAccount.set(account, branch);
  }

  // ============================================
  // 7. BUILD CREDIT LINE LOOKUP
  // ============================================

  // Key: Branch + Company Name
  // Value: Approved Credit Line in INR millions.

  const creditByCustomer = new Map<string, number>();

  for (const row of creditRows.slice(1)) {

    const branch = key(row[CREDIT_BRANCH]);
    const company = key(row[CREDIT_COMPANY]);

    if (!branch || !company) {
      continue;
    }

    const compositeKey = branch + "|" + company;

    if (creditByCustomer.has(compositeKey)) {
      throw new Error(
        "Duplicate CreditLine entry: " + compositeKey
      );
    }

    const creditMillion = toNumber(
      row[CREDIT_AMOUNT]
    );

    creditByCustomer.set(
      compositeKey,
      creditMillion
    );
  }

  // ============================================
  // 8. PROCESS FBL5n TRANSACTIONS
  // ============================================

  type Summary = {
    branch: string;
    customer: string;
    utilization: number;
  };

  const summaries = new Map<string, Summary>();

  const missingAccounts = new Set<string>();

  let processedRows = 0;

  for (const row of fblRows.slice(1)) {

    const account = key(row[FBL_ACCOUNT]);

    if (!account) {
      continue;
    }

    // Get Customer from VendorMaster.
    const customer = customerByAccount.get(account);

    // Get Branch from BUMaster.
    const branch = branchByAccount.get(account);

    if (!customer || !branch) {
      missingAccounts.add(account);
      continue;
    }

    // Read transaction amount.
    const amount = toNumber(row[FBL_AMOUNT]);

    // Group by Branch + Customer.
    const compositeKey =
      key(branch) + "|" + key(customer);

    const existing = summaries.get(compositeKey);

    if (existing) {

      // Add to existing utilization.
      existing.utilization += amount;

    } else {

      // Create new utilization record.
      summaries.set(compositeKey, {
        branch: branch,
        customer: customer,
        utilization: amount
      });
    }

    processedRows++;
  }

  // Stop if Account mappings are missing.
  if (missingAccounts.size > 0) {

    const examples = Array.from(missingAccounts)
      .slice(0, 15)
      .join(", ");

    throw new Error(
      "Missing Account mappings in VendorMaster " +
      "or BUMaster. Accounts: " + examples
    );
  }

  // ============================================
  // 9. PREPARE OUTPUT DATA
  // ============================================

  const output: Cell[][] = [[
    "Branch",
    "Customer (Open Credit)",
    "Credit Line",
    "Utilization",
    "%",
    "Remark"
  ]];

  const ordered = Array.from(summaries.values())
    .sort((a, b) => {

      const first =
        key(a.branch) + "|" + key(a.customer);

      const second =
        key(b.branch) + "|" + key(b.customer);

      return first.localeCompare(second);
    });

  let unmatchedCreditLines = 0;

  for (const item of ordered) {

    const compositeKey =
      key(item.branch) + "|" + key(item.customer);

    // Credit line in millions.
    const creditMillion =
      creditByCustomer.get(compositeKey);

    if (creditMillion === undefined) {
      unmatchedCreditLines++;
    }

    // Convert millions to INR for calculation.
    const creditINR =
      (creditMillion ?? 0) * 1000000;

    // Calculate utilization percentage.
    const percentage = creditINR > 0
      ? (item.utilization / creditINR) * 100
      : 0;

    // Add report row.
    output.push([
      item.branch,
      item.customer,
      creditMillion ?? 0,
      item.utilization,
      percentage,
      "" // Remark remains blank.
    ]);
  }

  // ============================================
  // 10. VALIDATE OUTPUT DIMENSIONS
  // ============================================

  if (output.length > 1000000) {
    throw new Error(
      "Too many output rows: " + output.length
    );
  }

  for (const row of output) {
    if (row.length !== 6) {
      throw new Error(
        "Output must contain exactly 6 columns"
      );
    }
  }

  // ============================================
  // 11. CREATE OR RESET REPORT WORKSHEET
  // ============================================

  let report = workbook.getWorksheet(
    "Credit Utilization Report"
  );

  if (!report) {
    report = workbook.addWorksheet(
      "Credit Utilization Report"
    );
  }

  const oldRange = report.getUsedRange();

  if (oldRange) {
    oldRange.clear(ExcelScript.ClearApplyTo.all);
  }

  // Write all report data in one operation.
	// Report title
  const titleRange = report.getRange("A1:F1");
  titleRange.merge(false);
  titleRange.setValue("Credit Utilization Report");

  titleRange.getFormat().getFont().setBold(true);
  titleRange.getFormat().getFont().setSize(18);
  titleRange.getFormat().setRowHeight(35);

  // Colour coding legend
  report.getRange("A3:B3").merge(false);
  report.getRange("A3").setValue("0–20%: Yellow/Orange"); 
  report.getRange("A3:B3").getFormat().getFill()
	  .setColor("#F4B183");

  report.getRange("C3:D3").merge(false);
  report.getRange("C3").setValue("21–70%: Light Blue");
  report.getRange("C3:D3").getFormat().getFill()
	  .setColor("#B7DEE8");

  report.getRange("E3:F3").merge(false);
  report.getRange("E3").setValue("Above 70%: Pink");
  report.getRange("E3:F3").getFormat().getFill()
	  .setColor("#DAA0C1");

	// Write headers and data from row 5
  report.getRangeByIndexes(
	  4, 0, output.length, 6
	).setValues(output);

  // ============================================
  // 12. FORMAT REPORT HEADER
  // ============================================

  const headerRange = report.getRange("A5:F5");

  headerRange.getFormat()
    .getFill()
    .setColor("#D9E2F3");

  headerRange.getFormat()
    .getFont()
    .setBold(true);

  headerRange.getFormat().setWrapText(true);
  headerRange.getFormat().setRowHeight(32);

  // Column widths.
  report.getRange("A:A")
    .getFormat().setColumnWidth(110);

  report.getRange("B:B")
    .getFormat().setColumnWidth(300);

  report.getRange("C:D")
    .getFormat().setColumnWidth(140);

  report.getRange("E:E")
    .getFormat().setColumnWidth(100);

  report.getRange("F:F")
    .getFormat().setColumnWidth(150);

  // ============================================
  // 13. NUMBER AND PERCENTAGE FORMATTING
  // ============================================

  if (output.length > 1) {

    const lastRow = output.length + 4;

    // Credit Line and Utilization.
  report.getRange(`C6:D${lastRow}`)
  .setNumberFormat("#,##0.00");

  report.getRange(`E6:E${lastRow}`)
  .setNumberFormat('0.00"%"');

    // ==========================================
    // 14. PERCENTAGE COLOR CODING
    // ==========================================

    for (let i = 1; i < output.length; i++) {

      const percentage = Number(output[i][4]);

      const cell = report.getCell(i + 4, 4);

      if (percentage >= 0 && percentage <= 20) {

        // 0-20%: Yellow/orange.
        cell.getFormat().getFill()
          .setColor("#F4B183");

      } else if (percentage > 20 && percentage <= 70) {

        // Above 20% through 70%: Light blue.
        cell.getFormat().getFill()
          .setColor("#B7DEE8");

      } else if (percentage > 70) {

        // Above 70%: Pink.
        cell.getFormat().getFill()
          .setColor("#DAA0C1");
      }
    }
  }

  // ============================================
  // 15. RETURN SUCCESS MESSAGE
  // ============================================

  return (
    "Report generated successfully. " +
    "Customers: " + ordered.length +
    ", FBL5n rows processed: " + processedRows +
    ", Missing credit lines: " + unmatchedCreditLines
  );
}
