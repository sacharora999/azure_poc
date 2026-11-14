/** Office Scripts solution: Expense Reconciliation and Exceptions */
function main(workbook: ExcelScript.Workbook) {
  const source = workbook.getWorksheet("Input");
  if (!source) throw new Error("Input worksheet missing");
  const used = source.getUsedRange();
  if (!used) throw new Error("No input data");
  const data = used.getValues().slice(1);
  let report = workbook.getWorksheet("Report");
  if (report) report.delete();
  report = workbook.addWorksheet("Report");
  const output: (string|number)[][]=[["Transaction ID","Department","Claimed","Approved","Difference","Status"]];
  for(const r of data){const claimed=Number(r[2])||0;const approved=Number(r[3])||0;const diff=claimed-approved;output.push([String(r[0]),String(r[1]),claimed,approved,diff,Math.abs(diff)>1000?"EXCEPTION":"OK"]);}
  report.getRangeByIndexes(0,0,output.length,output[0].length).setValues(output);
  const header=report.getRangeByIndexes(0,0,1,output[0].length);
  header.getFormat().getFill().setColor("#155E75");
  header.getFormat().getFont().setColor("#FFFFFF");
  header.getFormat().getFont().setBold(true);
  report.getUsedRange().getFormat().autofitColumns();
  report.getUsedRange().getFormat().autofitRows();
  report.getFreezePanes().freezeRows(1);
  report.getUsedRange().setNumberFormat('#,##0;[Red](#,##0);0');
  
}
