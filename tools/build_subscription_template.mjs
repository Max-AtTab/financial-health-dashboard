import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outDir = "/Users/max/Desktop/ MB/.subscription-build/outputs";
const outFile = `${outDir}/陪跑訂閱_月度營運儀表板模板.xlsx`;
const navy = "#1F4E78";
const blue = "#D9EAF7";
const amber = "#FFF2CC";
const green = "#E2F0D9";
const red = "#FCE4D6";
const grey = "#F2F2F2";
const font = "Arial";

const wb = Workbook.create();
const dashboard = wb.worksheets.add("Dashboard");
const monthly = wb.worksheets.add("Monthly Input");
const actions = wb.worksheets.add("Action Plan");
const meeting = wb.worksheets.add("Monthly Review");
const report = wb.worksheets.add("Client Report");
const guide = wb.worksheets.add("Instructions");

for (const ws of [dashboard, monthly, actions, meeting, report, guide]) {
  ws.showGridLines = false;
  ws.getRange("A1:Z80").format.font = { name: font, size: 10, color: "#1F2937" };
}

function header(ws, range) {
  ws.getRange(range).format = { fill: navy, font: { name: font, size: 10, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true };
}
function section(ws, range) {
  ws.getRange(range).format = { fill: blue, font: { name: font, size: 10, bold: true, color: "#1F2937" }, verticalAlignment: "center" };
}
function input(ws, range, numberFormat) {
  ws.getRange(range).format = { fill: amber, font: { name: font, size: 10, color: "#0000FF" }, verticalAlignment: "center" };
  if (numberFormat) ws.getRange(range).format.numberFormat = numberFormat;
}
function formula(ws, range, numberFormat) {
  ws.getRange(range).format = { font: { name: font, size: 10, color: "#000000" }, verticalAlignment: "center" };
  if (numberFormat) ws.getRange(range).format.numberFormat = numberFormat;
}
function borders(ws, range) {
  ws.getRange(range).format.borders = { preset: "all", style: "thin", color: "#D9E2F3" };
}

// Dashboard
dashboard.getRange("A2:H2").merge();
dashboard.getRange("A2").values = [["陪跑訂閱：月度營運儀表板"]];
dashboard.getRange("A2:H2").format = { font: { name: font, size: 15, bold: true, color: navy }, verticalAlignment: "center" };
dashboard.getRange("A3").values = [["客戶名稱"]];
dashboard.getRange("B3:D3").merge();
dashboard.getRange("B3").values = [[""]];
dashboard.getRange("E3").values = [["報告月份"]];
dashboard.getRange("F3").formulas = [["=IF(COUNT('Monthly Input'!$A$5:$A$40)=0,\"\",MAX('Monthly Input'!$A$5:$A$40))"]];
dashboard.getRange("G3:H3").merge();
dashboard.getRange("G3").values = [["黃色欄位請填；其餘自動計算"]];
section(dashboard,"A3:A3"); section(dashboard,"E3:E3");
input(dashboard,"B3:D3"); formula(dashboard,"F3","yyyy-mm");
dashboard.getRange("G3:H3").format = { fill: grey, font: { name: font, size: 9, italic: true, color: "#595959" }, wrapText: true };
dashboard.getRange("A5:H5").merge(); dashboard.getRange("A5").values = [["本月關鍵指標"]]; section(dashboard,"A5:H5");
dashboard.getRange("A6:H6").values = [["營收","毛利率","營業利益率","期末現金","現金可支撐月數","目標達成率","淨客戶變動","需處理事項"]]; header(dashboard,"A6:H6");
dashboard.getRange("A7").formulas = [["=IF($F$3=\"\",\"\",SUMIFS('Monthly Input'!$B$5:$B$40,'Monthly Input'!$A$5:$A$40,$F$3))"]];
dashboard.getRange("B7").formulas = [["=IF(OR($F$3=\"\",A7=0),\"\",SUMIFS('Monthly Input'!$C$5:$C$40,'Monthly Input'!$A$5:$A$40,$F$3)/A7)"]];
dashboard.getRange("C7").formulas = [["=IF(OR($F$3=\"\",A7=0),\"\",(SUMIFS('Monthly Input'!$C$5:$C$40,'Monthly Input'!$A$5:$A$40,$F$3)-SUMIFS('Monthly Input'!$D$5:$D$40,'Monthly Input'!$A$5:$A$40,$F$3))/A7)"]];
dashboard.getRange("D7").formulas = [["=IF($F$3=\"\",\"\",SUMIFS('Monthly Input'!$E$5:$E$40,'Monthly Input'!$A$5:$A$40,$F$3))"]];
dashboard.getRange("E7").formulas = [["=IF(OR($F$3=\"\",SUMIFS('Monthly Input'!$D$5:$D$40,'Monthly Input'!$A$5:$A$40,$F$3)=0),\"\",D7/SUMIFS('Monthly Input'!$D$5:$D$40,'Monthly Input'!$A$5:$A$40,$F$3))"]];
dashboard.getRange("F7").formulas = [["=IF(OR($F$3=\"\",SUMIFS('Monthly Input'!$I$5:$I$40,'Monthly Input'!$A$5:$A$40,$F$3)=0),\"\",SUMIFS('Monthly Input'!$J$5:$J$40,'Monthly Input'!$A$5:$A$40,$F$3)/SUMIFS('Monthly Input'!$I$5:$I$40,'Monthly Input'!$A$5:$A$40,$F$3))"]];
dashboard.getRange("G7").formulas = [["=IF($F$3=\"\",\"\",SUMIFS('Monthly Input'!$K$5:$K$40,'Monthly Input'!$A$5:$A$40,$F$3)-SUMIFS('Monthly Input'!$L$5:$L$40,'Monthly Input'!$A$5:$A$40,$F$3))"]];
dashboard.getRange("H7").formulas = [["=IF($F$3=\"\",\"\",COUNTIFS('Action Plan'!$H$5:$H$50,\"<>完成\",'Action Plan'!$H$5:$H$50,\"<>\"))"]];
formula(dashboard,"A7","#,##0;[Red](#,##0);-"); formula(dashboard,"B7:C7","0.0%"); formula(dashboard,"D7","#,##0;[Red](#,##0);-"); formula(dashboard,"E7","0.0"); formula(dashboard,"F7","0.0%"); formula(dashboard,"G7:H7","#,##0;[Red](#,##0);-");
borders(dashboard,"A6:H7");
dashboard.getRange("A9:H9").merge(); dashboard.getRange("A9").values = [["最近 12 個月趨勢"]]; section(dashboard,"A9:H9");
dashboard.getRange("A10:H10").values = [["月份","營收","毛利率","期末現金","現金可支撐月數","目標達成率","新客戶","流失客戶"]]; header(dashboard,"A10:H10");
for (let r=0; r<12; r++) {
  const sourceRow = 5+r;
  const targetRow = 11+r;
  dashboard.getRange(`A${targetRow}:H${targetRow}`).formulas = [[
    `=IF('Monthly Input'!A${sourceRow}=\"\",\"\",'Monthly Input'!A${sourceRow})`,
    `=IF(A${targetRow}=\"\",\"\",'Monthly Input'!B${sourceRow})`,
    `=IF(OR(A${targetRow}=\"\",B${targetRow}=0),\"\",'Monthly Input'!C${sourceRow}/B${targetRow})`,
    `=IF(A${targetRow}=\"\",\"\",'Monthly Input'!E${sourceRow})`,
    `=IF(OR(A${targetRow}=\"\",'Monthly Input'!D${sourceRow}=0),\"\",D${targetRow}/'Monthly Input'!D${sourceRow})`,
    `=IF(OR(A${targetRow}=\"\",'Monthly Input'!I${sourceRow}=0),\"\",'Monthly Input'!J${sourceRow}/'Monthly Input'!I${sourceRow})`,
    `=IF(A${targetRow}=\"\",\"\",'Monthly Input'!K${sourceRow})`,
    `=IF(A${targetRow}=\"\",\"\",'Monthly Input'!L${sourceRow})`
  ]];
}
formula(dashboard,"A11:A22","yyyy-mm"); formula(dashboard,"B11:B22","#,##0;[Red](#,##0);-"); formula(dashboard,"C11:C22","0.0%"); formula(dashboard,"D11:D22","#,##0;[Red](#,##0);-"); formula(dashboard,"E11:E22","0.0"); formula(dashboard,"F11:F22","0.0%"); formula(dashboard,"G11:H22","#,##0;[Red](#,##0);-"); borders(dashboard,"A10:H22");
dashboard.getRange("A24:H24").merge(); dashboard.getRange("A24").values = [["顧問判讀與下月重點"]]; section(dashboard,"A24:H24");
dashboard.getRange("A25").values = [["本月結論"]]; dashboard.getRange("B25:H26").merge(); dashboard.getRange("B25").values = [[""]];
dashboard.getRange("A27").values = [["下月優先事項"]]; dashboard.getRange("B27:H28").merge(); dashboard.getRange("B27").values = [[""]];
section(dashboard,"A25:A25"); section(dashboard,"A27:A27"); input(dashboard,"B25:H26"); input(dashboard,"B27:H28"); borders(dashboard,"A25:H28");
dashboard.getRange("A2:H28").format.wrapText = true;
dashboard.getRange("A:A").format.columnWidth = 15; dashboard.getRange("B:B").format.columnWidth = 17; dashboard.getRange("C:C").format.columnWidth = 14; dashboard.getRange("D:D").format.columnWidth = 17; dashboard.getRange("E:F").format.columnWidth = 17; dashboard.getRange("G:H").format.columnWidth = 18;
dashboard.getRange("A2:H28").format.rowHeight = 21; dashboard.getRange("A2:H2").format.rowHeight = 30;

// Monthly Input
monthly.getRange("A2:M2").merge(); monthly.getRange("A2").values = [["月度營運資料輸入"]]; monthly.getRange("A2:M2").format = { font: { name: font, size: 15, bold: true, color: navy } };
monthly.getRange("A3:M3").merge(); monthly.getRange("A3").values = [["每月填一列。請以完整月份、同一會計口徑填寫；黃色欄位為輸入。"]]; monthly.getRange("A3:M3").format = { fill: grey, font: { name: font, size: 9, italic: true, color: "#595959" } };
monthly.getRange("A4:M4").values = [["月份","營收","毛利","營業費用","期末現金","應收帳款","應付帳款","存貨","業績目標","實際業績","新客戶","流失客戶","備註"]]; header(monthly,"A4:M4");
for (let r=5; r<=40; r++) { input(monthly,`A${r}:M${r}`); }
monthly.getRange("A5:A40").format.numberFormat = "yyyy-mm";
monthly.getRange("B5:J40").format.numberFormat = "#,##0;[Red](#,##0);-";
monthly.getRange("K5:L40").format.numberFormat = "#,##0;[Red](#,##0);-";
borders(monthly,"A4:M40");
monthly.getRange("A:A").format.columnWidth=13; monthly.getRange("B:J").format.columnWidth=14; monthly.getRange("K:L").format.columnWidth=12; monthly.getRange("M:M").format.columnWidth=30;
monthly.freezePanes.freezeRows(4);

// Action plan
actions.getRange("A2:J2").merge(); actions.getRange("A2").values = [["陪跑行動計畫"]]; actions.getRange("A2:J2").format = { font: { name: font, size: 15, bold: true, color: navy } };
actions.getRange("A3:J3").merge(); actions.getRange("A3").values = [["每次月度檢討後更新。完成前的事項會計入 Dashboard 的需處理事項。"]]; actions.getRange("A3:J3").format = { fill: grey, font: { name: font, size: 9, italic: true, color: "#595959" } };
actions.getRange("A4:J4").values = [["行動項目","主題","問題／機會","量化目標","負責人","開始日","到期日","狀態","完成率","下次檢核證據"]]; header(actions,"A4:J4");
for (let r=5; r<=50; r++) input(actions,`A${r}:J${r}`);
actions.getRange("F5:G50").format.numberFormat="yyyy-mm-dd"; actions.getRange("I5:I50").format.numberFormat="0%";
actions.getRange("B5:B50").dataValidation={rule:{type:"list",values:["財務","獲利","業績","客戶","組織","流程","策略"]}};
actions.getRange("H5:H50").dataValidation={rule:{type:"list",values:["未開始","進行中","待確認","完成","暫停"]}};
borders(actions,"A4:J50");
actions.getRange("A:A").format.columnWidth=22; actions.getRange("B:B").format.columnWidth=12; actions.getRange("C:D").format.columnWidth=25; actions.getRange("E:E").format.columnWidth=14; actions.getRange("F:G").format.columnWidth=13; actions.getRange("H:H").format.columnWidth=12; actions.getRange("I:I").format.columnWidth=10; actions.getRange("J:J").format.columnWidth=28; actions.freezePanes.freezeRows(4);

// Monthly Review
meeting.getRange("A2:H2").merge(); meeting.getRange("A2").values=[["月度檢討會議紀錄"]]; meeting.getRange("A2:H2").format={font:{name:font,size:15,bold:true,color:navy}};
meeting.getRange("A4:B7").values=[["客戶名稱", ""],["會議月份", ""],["出席者", ""],["下次會議", ""]]; section(meeting,"A4:A7"); input(meeting,"B4:B7"); meeting.getRange("B5:B5").format.numberFormat="yyyy-mm"; meeting.getRange("B7:B7").format.numberFormat="yyyy-mm-dd";
meeting.getRange("A9:H9").merge(); meeting.getRange("A9").values=[["固定議程"]]; section(meeting,"A9:H9");
meeting.getRange("A10:D10").values=[["順序","議題","討論重點","決議"]]; header(meeting,"A10:D10");
meeting.getRange("A11:D15").values=[[1,"上月行動回顧","",""],[2,"本月指標與異常","",""],[3,"客戶／業績／毛利變化","",""],[4,"現金與營運風險","",""],[5,"下月行動與資源需求","",""]];
input(meeting,"C11:D15"); borders(meeting,"A10:D15");
meeting.getRange("A17:H17").merge(); meeting.getRange("A17").values=[["本月需客戶決策或支援事項"]]; section(meeting,"A17:H17"); meeting.getRange("A18:H22").merge(); input(meeting,"A18:H22"); borders(meeting,"A17:H22");
meeting.getRange("A:A").format.columnWidth=10; meeting.getRange("B:B").format.columnWidth=25; meeting.getRange("C:D").format.columnWidth=34; meeting.getRange("A2:H22").format.wrapText=true;

// Client Report
report.getRange("A2:H2").merge(); report.getRange("A2").values=[["月度陪跑報告"]]; report.getRange("A2:H2").format={font:{name:font,size:16,bold:true,color:navy}};
report.getRange("A3").values=[["客戶名稱"]]; report.getRange("B3:D3").merge(); report.getRange("B3").formulas=[["=Dashboard!B3"]]; report.getRange("E3").values=[["報告月份"]]; report.getRange("F3").formulas=[["=Dashboard!F3"]]; section(report,"A3:A3"); section(report,"E3:E3"); formula(report,"B3:D3"); formula(report,"F3","yyyy-mm");
report.getRange("A5:H5").merge(); report.getRange("A5").values=[["經營摘要"]]; section(report,"A5:H5");
report.getRange("A6:H6").values=[["營收","毛利率","營業利益率","期末現金","現金可支撐月數","目標達成率","淨客戶變動","未完成行動"]]; header(report,"A6:H6");
report.getRange("A7:H7").formulas=[["=Dashboard!A7","=Dashboard!B7","=Dashboard!C7","=Dashboard!D7","=Dashboard!E7","=Dashboard!F7","=Dashboard!G7","=Dashboard!H7"]]; formula(report,"A7","#,##0;[Red](#,##0);-"); formula(report,"B7:C7","0.0%"); formula(report,"D7","#,##0;[Red](#,##0);-"); formula(report,"E7","0.0"); formula(report,"F7","0.0%"); formula(report,"G7:H7","#,##0;[Red](#,##0);-"); borders(report,"A6:H7");
report.getRange("A9:H9").merge(); report.getRange("A9").values=[["顧問結論"]]; section(report,"A9:H9"); report.getRange("A10:H12").merge(); report.getRange("A10").formulas=[["=Dashboard!B25"]]; formula(report,"A10:H12"); borders(report,"A9:H12");
report.getRange("A14:H14").merge(); report.getRange("A14").values=[["下月優先事項"]]; section(report,"A14:H14"); report.getRange("A15:H17").merge(); report.getRange("A15").formulas=[["=Dashboard!B27"]]; formula(report,"A15:H17"); borders(report,"A14:H17");
report.getRange("A19:H19").merge(); report.getRange("A19").values=[["行動追蹤"]]; section(report,"A19:H19"); report.getRange("A20:E20").values=[["行動項目","負責人","到期日","狀態","完成率"]]; header(report,"A20:E20");
for (let r=0;r<8;r++) { const src=5+r, dst=21+r; report.getRange(`A${dst}:E${dst}`).formulas=[[`=IF('Action Plan'!A${src}=\"\",\"\",'Action Plan'!A${src})`,`=IF(A${dst}=\"\",\"\",'Action Plan'!E${src})`,`=IF(A${dst}=\"\",\"\",'Action Plan'!G${src})`,`=IF(A${dst}=\"\",\"\",'Action Plan'!H${src})`,`=IF(A${dst}=\"\",\"\",'Action Plan'!I${src})`]]; }
formula(report,"C21:C28","yyyy-mm-dd"); formula(report,"E21:E28","0%"); borders(report,"A20:E28"); report.getRange("A2:H28").format.wrapText=true;
report.getRange("A:A").format.columnWidth=22; report.getRange("B:B").format.columnWidth=15; report.getRange("C:C").format.columnWidth=14; report.getRange("D:D").format.columnWidth=15; report.getRange("E:H").format.columnWidth=17;

// Instructions
guide.getRange("A2:H2").merge(); guide.getRange("A2").values=[["使用說明"]]; guide.getRange("A2:H2").format={font:{name:font,size:15,bold:true,color:navy}};
guide.getRange("A4:B9").values=[
 ["步驟","操作"],
 ["1","每月在 Monthly Input 新增一列，填入實際數字。"],
 ["2","在 Action Plan 更新行動、負責人、期限與完成率。"],
 ["3","Dashboard 自動抓取最新月份；顧問填寫本月結論與下月優先事項。"],
 ["4","月度檢討時填 Monthly Review；確認決議和需客戶支援事項。"],
 ["5","交付前檢查 Client Report，列印或匯出 PDF 給客戶。"]
]; header(guide,"A4:B4"); borders(guide,"A4:B9"); guide.getRange("A5:B9").format.wrapText=true;
guide.getRange("A11:H11").merge(); guide.getRange("A11").values=[["每月需提供的資料"]]; section(guide,"A11:H11");
guide.getRange("A12:D18").values=[
 ["資料","來源","用途","頻率"],
 ["營收、毛利、營業費用","月結損益表","獲利趨勢與目標達成","每月"],
 ["期末現金、應收、應付、存貨","銀行與帳齡表","流動性與營運資金","每月"],
 ["業績目標與實績","業務報表","業績差異與客戶變動","每月"],
 ["新客戶、流失客戶","CRM 或業務名單","客戶健康度","每月"],
 ["行動進度與證據","各負責人","改善計畫追蹤","每月"]
]; header(guide,"A12:D12"); borders(guide,"A12:D18"); guide.getRange("A12:D18").format.wrapText=true;
guide.getRange("A:A").format.columnWidth=18; guide.getRange("B:B").format.columnWidth=34; guide.getRange("C:C").format.columnWidth=30; guide.getRange("D:D").format.columnWidth=14;

// Conditional formatting only applies when a meaningful status has been entered.
actions.getRange("H5:H50").conditionalFormats.add("containsText",{text:"完成",format:{fill:green}});
actions.getRange("H5:H50").conditionalFormats.add("containsText",{text:"暫停",format:{fill:red}});

wb.recalculate();
const checks = await wb.inspect({kind:"table",range:"Dashboard!A2:H28",include:"values,formulas",tableMaxRows:30,tableMaxCols:10});
const errors = await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",options:{useRegex:true,maxResults:100},summary:"formula error scan"});
await fs.mkdir(outDir,{recursive:true});
const preview = await wb.render({sheetName:"Dashboard",range:"A2:H28",scale:1.2,format:"png"});
await fs.writeFile(`${outDir}/dashboard_preview.png`,new Uint8Array(await preview.arrayBuffer()));
const file = await SpreadsheetFile.exportXlsx(wb);
await file.save(outFile);
console.log(JSON.stringify({outFile,checks:checks.ndjson,errors:errors.ndjson}));
