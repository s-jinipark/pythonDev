
# 다양한 방법으로 워크시트 값 읽기

import openpyxl as excel


# 워크북 열기
book = excel.load_workbook('write_cellname.xlsx')
# 워크시트 읽기
sheet = book.active

# 셀 H2의 값 읽기
print(sheet['H2'].value)

# 셀 H2의 값 읽기
cell = sheet.cell(row=2, column=8)
print(cell.value)

print('--------------------')

