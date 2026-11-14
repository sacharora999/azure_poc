/** Office Scripts solution: Monthly Profit and Loss Consolidation */
function main(workbook: ExcelScript.Workbook) {
  const source = workbook.getWorksheet("Input");
  if (!source) throw new Error("Input worksheet missing");
  const used = source.getUsedRange();
  if (!used) throw new Error("No input data");
  const data = used.getValues().slice(1);
  let report = workbook.getWorksheet("Report");
  if (report) report.delete();
  report = workbook.addWorksheet("Report");
  const map = new Map<string, number[]>();
  for(const r of data) { const month=String(r[0]); const type=String(r[2]); const a=map.get(month)??[0,0,0]; const amt=Number(r[3])||0; if(type==="Revenue")a[0]+=amt; else if(type==="COGS")a[1]+=amt; else a[2]+=amt; map.set(month,a); }
  const output: (string|number)[][]=[["Month","Revenue","COGS","Gross Profit","Other Expenses","Net Profit"]];
  for(const [m,a] of map)output.push([m,a[0],a[1],a[0]-a[1],a[2],a[0]-a[1]-a[2]]);
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
