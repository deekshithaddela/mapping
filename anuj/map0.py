import xlrd
import xlwt
import re
from difflib import SequenceMatcher

def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

# Give the location of the file
loc2 = ("cmdb_ci_appl1.xlsx")
loc1 = ("NIST- product List.xlsx")
loc3 = ("cmdb_ci_appl.xls")

# To open Workbook
wb1 = xlrd.open_workbook(loc1)
sheet1 = wb1.sheet_by_index(0)
wb2 = xlrd.open_workbook(loc2)
sheet2 = wb2.sheet_by_index(0)
wb3 = xlwt.Workbook()
sheet3 = wb3.add_sheet("Sheet 1", cell_overwrite_ok=True)

# For
for i in range(sheet2.nrows):
    p=re.sub(r'_', '', ''.join(a for a in ''.join(sheet2.cell_value(i, 0).lower().split()) if not a.isdigit()))
    v=''.join(sheet2.cell_value(i, 4).lower().split())
    c=0
    for j in range(sheet1.nrows):
        v1=''.join(sheet1.cell_value(j, 0).lower().split())
        p1=re.sub(r'_', '', ''.join(a for a in ''.join(sheet1.cell_value(j, 1).lower().split()) if not a.isdigit()))
        if v == v1:
            if p == p1:
                sheet3.write(i, 0, sheet2.cell_value(i, 0))
                sheet3.write(i, 1, sheet2.cell_value(i, 1))
                sheet3.write(i, 2, sheet2.cell_value(i, 2))
                sheet3.write(i, 3, sheet2.cell_value(i, 3))
                sheet3.write(i, 4, sheet2.cell_value(i, 4))
                sheet3.write(i, 6, sheet1.cell_value(j, 0))
                sheet3.write(i, 7, sheet1.cell_value(j, 1))
            elif similar(p,p1) >= 0.7:
                sheet3.write(i, 0, sheet2.cell_value(i, 0))
                sheet3.write(i, 1, sheet2.cell_value(i, 1))
                sheet3.write(i, 2, sheet2.cell_value(i, 2))
                sheet3.write(i, 3, sheet2.cell_value(i, 3))
                sheet3.write(i, 4, sheet2.cell_value(i, 4))
                sheet3.write(i, 6, sheet1.cell_value(j, 0))
                sheet3.write(i, 7, sheet1.cell_value(j, 1))
            else:
                c=c+1
        elif c+1 == sheet1.nrows:
            sheet3.write(i, 0, sheet2.cell_value(i, 0))
            sheet3.write(i, 1, sheet2.cell_value(i, 1))
            sheet3.write(i, 2, sheet2.cell_value(i, 2))
            sheet3.write(i, 3, sheet2.cell_value(i, 3))
            sheet3.write(i, 4, sheet2.cell_value(i, 4))
        else:
            c=c+1
wb3.save(loc3)
