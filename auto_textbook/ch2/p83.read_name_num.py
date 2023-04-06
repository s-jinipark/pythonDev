
# 셀 주소와 행ㆍ열 번호를 변환하는 프로그램

import openpyxl as excel

book = excel.Workbook()
sheet = book.active

# 셀 주소에서 행ㆍ열 번호 얻기
cell = sheet["C2"]
(row, col) = (cell.row, cell.column)
print("C2=({},{})".format(row, col))

print('--------------------')

# 행ㆍ열 번호 에서 셀 주소 얻기
cell = sheet.cell(row=2, column=3)
cdt = cell.coordinate
print("(2,3)={}".format(cdt))
