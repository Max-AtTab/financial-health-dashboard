from openpyxl import Workbook

output = "/tmp/twcf_formula_verification.xlsx"
wb = Workbook()
ws = wb.active
ws.title = "TWCF verification"

# Week 1 — healthy with 20% headroom.
ws["B7"] = 1000
ws["B9"] = 800
ws["B12"] = '=IF(COUNT(B9:B11)=0,"",SUM(B9:B11))'
ws["B14"] = 500
ws["B21"] = '=IF(COUNT(B14:B20)=0,"",SUM(B14:B20))'
ws["B22"] = '=IF(AND(B12="",B21=""),"",N(B12)-N(B21))'
ws["B23"] = '=IF(OR(B7="",B22=""),"",B7+B22)'
ws["B26"] = 1000
ws["B27"] = '=IF(OR(B23="",B26=""),"",B23-B26)'
ws["B28"] = '=IFERROR(B27/B26,"")'
ws["B29"] = '=IF(B23="","",IF(B23<0,"NEGATIVE",IF(B27<0,"BELOW FLOOR",IF(B28<0.2,"LOW HEADROOM","SAFE"))))'

# Week 2 — two consecutive negative weeks, but still above the floor.
ws["C7"] = '=B23'
ws["C9"] = 100
ws["C12"] = '=IF(COUNT(C9:C11)=0,"",SUM(C9:C11))'
ws["C14"] = 300
ws["C21"] = '=IF(COUNT(C14:C20)=0,"",SUM(C14:C20))'
ws["C22"] = '=IF(AND(C12="",C21=""),"",N(C12)-N(C21))'
ws["C23"] = '=IF(OR(C7="",C22=""),"",C7+C22)'
ws["C26"] = '=B26'
ws["C27"] = '=IF(OR(C23="",C26=""),"",C23-C26)'
ws["C28"] = '=IFERROR(C27/C26,"")'
ws["C29"] = '=IF(C23="","",IF(C23<0,"NEGATIVE",IF(C27<0,"BELOW FLOOR",IF(AND(B22<0,C22<0),"TWO WEEKS OUTFLOW",IF(C28<0.2,"LOW HEADROOM","SAFE")))))'

wb.save(output)
print(output)
