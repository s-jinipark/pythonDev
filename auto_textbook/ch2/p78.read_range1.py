
# for 문을 이용해 범위 내의 셀 얻기

import openpyxl as excel

# 워크북 열기
book = excel.load_workbook('output/write_cellname.xlsx')
# 워크시트 읽기
sheet = book.active

# 연속 데이터를 얻어 출력하기
for y in range(2,5):
    r = []
    for x in range(2,5):
        v = sheet.cell(row=y, column=x).value
        r.append(v)
    print(r)

print('--------------------')

