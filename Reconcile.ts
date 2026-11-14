/** Office Scripts solution: Departmental Budget vs Actual MIS */
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
  for (const r of data) { const key=String(r[0]); const a=map.get(key)??[0,0]; a[0]+=Number(r[2])||0; a[1]+=Number(r[3])||0; map.set(key,a); }
  const output: (string|number)[][]=[["Department","Budget","Actual","Variance","Variance %"]];
  for(const [key,v] of map) output.push([key,v[0],v[1],v[1]-v[0],v[0]===0?0:(v[1]-v[0])/v[0]]);
  const n=output.length+1; output.push(["GRAND TOTAL",0,0,0,0]);

  report.getRangeByIndexes(0,0,output.length,output[0].length).setValues(output);
  const header=report.getRangeByIndexes(0,0,1,output[0].length);
  header.getFormat().getFill().setColor("#155E75");
  header.getFormat().getFont().setColor("#FFFFFF");
  header.getFormat().getFont().setBold(true);
  report.getRange(`B${n}:E${n}`).setFormulas([[`=SUM(B2:B${n-1})`,`=SUM(C2:C${n-1})`,`=C${n}-B${n}`,`=IFERROR(D${n}/B${n},0)`]]);
  report.getUsedRange().getFormat().autofitColumns();
  report.getUsedRange().getFormat().autofitRows();
  report.getFreezePanes().freezeRows(1);
  report.getUsedRange().setNumberFormat('#,##0;[Red](#,##0);0');
  report.getRange(`E2:E${output.length+1}`).setNumberFormat("0.0%");
}
