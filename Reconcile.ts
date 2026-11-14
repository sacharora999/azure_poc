/** Office Scripts solution: YTD Branch MIS and Trend Report */
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
  for(const r of data){const key=`${String(r[1])}|${String(r[0])}`;const a=map.get(key)??[0,0];a[0]+=Number(r[2])||0;a[1]+=Number(r[3])||0;map.set(key,a);}
  const output: (string|number)[][]=[["Branch","Month","Revenue","Expense","Profit","YTD Profit"]];
  const months=["Jan","Feb","Mar","Apr","May","Jun"];
  for(const branch of ["NCL","MM","CHN","CORP"]){let ytd=0;for(const month of months){const a=map.get(`${branch}|${month}`)??[0,0];const profit=a[0]-a[1];ytd+=profit;output.push([branch,month,a[0],a[1],profit,ytd]);}}
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
