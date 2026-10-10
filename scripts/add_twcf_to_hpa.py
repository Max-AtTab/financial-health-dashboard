from copy import copy
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

SOURCE = "/tmp/yaoxi-finance-hpa-current.xlsx"
OUTPUT = "/tmp/yaoxi-finance-hpa-twcf.xlsx"
SHEET_NAME = "7. 13週現金流量表"


def copy_style(source, target):
    target._style = copy(source._style)
    if source.number_format:
        target.number_format = source.number_format
    target.alignment = copy(source.alignment)


wb = load_workbook(SOURCE, data_only=False)
if SHEET_NAME in wb.sheetnames:
    del wb[SHEET_NAME]

ws = wb.create_sheet(SHEET_NAME)
ws.sheet_view.showGridLines = False
ws.freeze_panes = "B6"
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.page_margins.left = 0.25
ws.page_margins.right = 0.25
ws.page_margins.top = 0.5
ws.page_margins.bottom = 0.5

# Reuse the existing HPA visual system rather than introducing a new design.
style_sheet = wb["6. 應付帳款與CCC HPA"]
title_style = style_sheet["A1"]
header_style = style_sheet["A3"]
label_style = style_sheet["A4"]
formula_style = style_sheet["B4"]
note_style = style_sheet["F4"]
section_style = style_sheet["A7"]

for col, width in {"A": 34, **{get_column_letter(i): 14 for i in range(2, 15)}, "O": 42}.items():
    ws.column_dimensions[col].width = width

for row, height in {1: 34.5, 2: 42, 3: 24, 5: 26, 6: 24}.items():
    ws.row_dimensions[row].height = height

# Title and operating instructions.
ws.merge_cells("A1:O1")
ws["A1"] = "表七：13週現金流量表（TWCF）｜直接法・滾動式"
copy_style(title_style, ws["A1"])
ws["A1"].alignment = Alignment(horizontal="left", vertical="center")

ws.merge_cells("A2:O2")
ws["A2"] = (
    "啟動時機：CCC＞90天、現金可支撐月數黃／紅燈，或任一生存風險紅燈。黃底為輸入格；每週結束以實際值覆蓋預測，並向右延伸一週，永遠保留未來13週能見度。"
)
copy_style(note_style, ws["A2"])
ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")

ws["A3"] = "第1週週一（請填）"
copy_style(label_style, ws["A3"])
ws["B3"] = None
copy_style(formula_style, ws["B3"])
ws["B3"].number_format = "yyyy/m/d"
ws["B3"].comment = Comment("請填第1週週一日期。第2週起會自動每週加7天。", "曜析數據")
ws.merge_cells("C3:O3")
ws["C3"] = "期初現金請使用銀行「實際可用餘額」（各帳戶加總、扣除已開立未兌現支票），不是 ERP 帳面數。"
copy_style(note_style, ws["C3"])
ws["C3"].alignment = Alignment(wrap_text=True, vertical="center")

# Period headers.
ws["A5"] = "週次"
copy_style(header_style, ws["A5"])
for week, col in enumerate(range(2, 15), start=1):
    cell = ws.cell(5, col)
    cell.value = f"第{week}週"
    copy_style(header_style, cell)
ws["O5"] = "說明"
copy_style(header_style, ws["O5"])

ws["A6"] = "週起始日"
copy_style(label_style, ws["A6"])
for col in range(2, 15):
    cell = ws.cell(6, col)
    cell.value = "=$B$3" if col == 2 else f"={get_column_letter(col - 1)}6+7"
    copy_style(formula_style, cell)
    cell.number_format = "m/d"
ws["O6"] = "自動：起始日＋7天遞延"
copy_style(note_style, ws["O6"])

# Helpers.
yellow = PatternFill(fill_type="solid", fgColor="FFFFF2CC")
blue = PatternFill(fill_type="solid", fgColor="FFD9E1F2")
red = PatternFill(fill_type="solid", fgColor="FFFADBD8")
thin = Side(style="thin", color="FFB7C9D6")
bottom_border = Border(bottom=thin)

def set_label(row, label, style_source=label_style):
    cell = ws.cell(row, 1)
    cell.value = label
    copy_style(style_source, cell)

def set_note(row, text):
    cell = ws.cell(row, 15)
    cell.value = text
    copy_style(note_style, cell)
    cell.alignment = Alignment(wrap_text=True, vertical="center")

def set_inputs(row, note, *, starting_col=2, ending_col=14):
    for col in range(starting_col, ending_col + 1):
        cell = ws.cell(row, col)
        copy_style(formula_style, cell)
        cell.fill = copy(yellow)
        cell.number_format = '#,##0;[Red](#,##0);-'
        cell.alignment = Alignment(horizontal="right", vertical="center")
    set_note(row, note)

def set_formula_row(row, formula_template, note, fmt='#,##0;[Red](#,##0);-'):
    for col in range(2, 15):
        letter = get_column_letter(col)
        cell = ws.cell(row, col)
        cell.value = formula_template.format(col=letter, prev=get_column_letter(col - 1))
        copy_style(formula_style, cell)
        cell.number_format = fmt
        cell.alignment = Alignment(horizontal="right", vertical="center")
    set_note(row, note)

def section(row, title):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=15)
    cell = ws.cell(row, 1)
    cell.value = title
    copy_style(section_style, cell)
    cell.font = Font(name="Microsoft JhengHei", size=10, bold=True, color="FF1F1F1F")
    cell.alignment = Alignment(horizontal="left", vertical="center")

# Core cash waterfall.
set_label(7, "期初現金餘額")
for col in range(2, 15):
    cell = ws.cell(7, col)
    if col == 2:
        cell.value = None
        copy_style(formula_style, cell)
        cell.fill = copy(yellow)
        cell.comment = Comment("請輸入銀行實際可用現金。", "曜析數據")
    else:
        cell.value = f"={get_column_letter(col - 1)}23"
        copy_style(formula_style, cell)
    cell.number_format = '#,##0;[Red](#,##0);-'
    cell.alignment = Alignment(horizontal="right", vertical="center")
set_note(7, "第1週請填；第2週起自動承接前一週期末現金")

section(8, "── 現金流入 Receipts（依實際「入帳」週次填，不是開發票週次）──")
set_label(9, "應收帳款收現（帳齡×收現率）")
set_inputs(9, "依 AR 帳齡與歷史收現率排入；現銷為主可留空")
set_label(10, "現金銷售收入（現銷）")
set_inputs(10, "餐飲／零售主要流入；商場月結款請填實際入帳週")
set_label(11, "其他現金流入")
set_inputs(11, "退稅、利息、資產出售、股東投入等")
set_label(12, "流入小計", style_sheet["A10"])
set_formula_row(12, '=IF(COUNT({col}9:{col}11)=0,"",SUM({col}9:{col}11))', "自動加總")

section(13, "── 現金流出 Disbursements（依實際「付款」週次填；發薪日、債務到期日不可平均攤）──")
for row, label, note in [
    (14, "員工薪資／勞健保", "依實際發薪日與勞健保扣繳日填入"),
    (15, "租金／水電／固定費", "依合約付款週次填入"),
    (16, "供應商／採購付款", "依 AP 帳齡與付款條件排入"),
    (17, "稅款", "營業稅、營所稅、扣繳等確切付款週次"),
    (18, "債務本息償還", "依還款表填；錯過到期日可能構成違約"),
    (19, "資本支出／非營運支出", "設備、裝修、法律費、一次性支出"),
    (20, "應急緩衝", "保留突發設備維修／退貨／法律費等預備額"),
]:
    set_label(row, label)
    set_inputs(row, note)
set_label(21, "流出小計", style_sheet["A10"])
set_formula_row(21, '=IF(COUNT({col}14:{col}20)=0,"",SUM({col}14:{col}20))', "自動加總")

set_label(22, "淨現金流", style_sheet["A10"])
set_formula_row(22, '=IF(AND({col}12="",{col}21=""),"",N({col}12)-N({col}21))', "流入小計－流出小計")
set_label(23, "期末現金餘額", style_sheet["A10"])
set_formula_row(23, '=IF(OR({col}7="",{col}22=""),"",{col}7+{col}22)', "期初現金＋淨現金流；第2週會帶入下一週期初")

section(25, "── 流動性防線與預警 ──")
set_label(26, "最低流動性門檻（請填）")
for col in range(2, 15):
    cell = ws.cell(26, col)
    if col == 2:
        cell.value = None
        copy_style(formula_style, cell)
        cell.fill = copy(yellow)
        cell.comment = Comment("請填銀行契約或管理階層設定的最低可用現金；沒有硬性門檻，可先填一個月固定支出。", "曜析數據")
    else:
        cell.value = "=$B$26"
        copy_style(formula_style, cell)
    cell.number_format = '#,##0;[Red](#,##0);-'
    cell.alignment = Alignment(horizontal="right", vertical="center")
set_note(26, "銀行 covenant 或內部最低可用現金；沒有時先用一個月固定支出")
set_label(27, "寬裕度 Headroom")
set_formula_row(27, '=IF(OR({col}23="",{col}26=""),"",{col}23-{col}26)', "期末現金－最低門檻；負數代表跌破防線")
set_label(28, "寬裕度 %")
set_formula_row(28, '=IFERROR({col}27/{col}26,"")', "低於20%即應啟動防禦行動", fmt="0.0%")
set_label(29, "系統判讀")
for col in range(2, 15):
    letter = get_column_letter(col)
    prev = get_column_letter(col - 1)
    if col == 2:
        formula = f'=IF({letter}23="","",IF({letter}23<0,"🚨 現金轉負",IF({letter}27<0,"🚨 低於最低流動性",IF({letter}28<0.2,"⚠️ 寬裕度<20%","✅ 安全"))))'
    else:
        formula = f'=IF({letter}23="","",IF({letter}23<0,"🚨 現金轉負",IF({letter}27<0,"🚨 低於最低流動性",IF(AND({prev}22<0,{letter}22<0),"⚠️ 連兩週淨流出",IF({letter}28<0.2,"⚠️ 寬裕度<20%","✅ 安全")))))'
    cell = ws.cell(29, col)
    cell.value = formula
    copy_style(formula_style, cell)
    cell.alignment = Alignment(horizontal="center", vertical="center")
set_note(29, "依序檢查：轉負、跌破門檻、連兩週淨流出、寬裕度<20%")

# AR collection assumptions make the model explainable without forcing data duplication.
section(31, "── 應收帳款帳齡收現率假設（可依過去12個月實績校準；僅作第9列排程參考）──")
for col, label in enumerate(["0–30天", "31–60天", "61–90天", ">90天"], start=1):
    cell = ws.cell(32, col)
    cell.value = label
    copy_style(header_style, cell)
for col, value in enumerate([0.90, 0.60, 0.30, 0.10], start=1):
    cell = ws.cell(33, col)
    cell.value = value
    copy_style(formula_style, cell)
    cell.fill = copy(yellow)
    cell.number_format = "0%"
    cell.alignment = Alignment(horizontal="center", vertical="center")
ws.merge_cells("E32:O33")
ws["E32"] = "參考假設：請用企業過去12個月「實際收款率」校準；不要直接把業務預估當作現金流入。"
copy_style(note_style, ws["E32"])
ws["E32"].alignment = Alignment(wrap_text=True, vertical="center")

# Align visual treatment, numeric formatting and row heights.
for row in range(5, 30):
    ws.row_dimensions[row].height = 22
    for col in range(1, 16):
        ws.cell(row, col).border = copy(bottom_border)

for row in [12, 21, 22, 23, 27, 28, 29]:
    for col in range(1, 16):
        ws.cell(row, col).font = copy(ws.cell(row, col).font)
        ws.cell(row, col).font = Font(
            name=ws.cell(row, col).font.name or "Microsoft JhengHei",
            size=ws.cell(row, col).font.sz or 10,
            bold=True,
            italic=ws.cell(row, col).font.italic,
            color=ws.cell(row, col).font.color,
        )

for row in [23, 27, 29]:
    for col in range(1, 16):
        ws.cell(row, col).fill = copy(red if row == 23 else blue)

# Data validation protects the two date/amount controls without constraining finance values.
date_validation = DataValidation(type="date", operator="between", formula1="DATE(2020,1,1)", formula2="DATE(2100,12,31)", allow_blank=True)
ws.add_data_validation(date_validation)
date_validation.add(ws["B3"])

ws.auto_filter.ref = "A5:O29"
wb.calculation.fullCalcOnLoad = True
wb.calculation.forceFullCalc = True
wb.calculation.calcMode = "auto"
wb.save(OUTPUT)
print(OUTPUT)
