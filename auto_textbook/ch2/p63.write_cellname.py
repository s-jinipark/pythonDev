
# 셀주소를 표현하는 방법

import openpyxl as excel

# 워크북을 생성하고 활성화된 워크시트 가져오기
book = excel.Workbook()
sheet = book.active

# 시트에 셀 주소 채우기
for y in range(1,101):
    for x in range(1,101):
        cell = sheet.cell(row=y, column=x)
        cell.value = cell.coordinate   # 셀 주소 가져오기

# 파일 (워크북) 저장
book.save('output/write_cellname.xlsx')

print('--------------------')

