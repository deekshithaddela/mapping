import xlrd
import xlwt

# Give the location of the file
loc1 = ("C:/Users/daddela/Downloads/mapping/1.xlsx")
loc2 = ("C:/Users/daddela/Downloads/mapping/2.xlsx")
loc3 = ("C:/Users/daddela/Downloads/mapping/3.xls")

# To open Workbook
wb1 = xlrd.open_workbook(loc1)
sheet1 = wb1.sheet_by_index(0)
wb2 = xlrd.open_workbook(loc2)
sheet2 = wb2.sheet_by_index(0)
wb3 = xlwt.Workbook()
sheet3 = wb3.add_sheet("sheet")

# For
for i in range(sheet1.nrows):
    v=''.join(a for a in ''.join(sheet1.cell_value(i, 0).lower().split()) if not a.isdigit())
    p=''.join(a for a in ''.join(sheet1.cell_value(i, 1).lower().split()) if not a.isdigit())
    for j in range(sheet2.nrows):
        v1=''.join(a for a in ''.join(sheet1.cell_value(j, 0).lower().split()) if not a.isdigit())
        p1=''.join(a for a in ''.join(sheet1.cell_value(j, 1).lower().split()) if not a.isdigit())
        if v == v1 and p == p1:
            sheet3.write(i, 0, v)
            sheet3.write(i, 1, p)
            sheet3.write(i, 2, v1)
            sheet3.write(i, 3, p1)
wb3.save(loc3)
