
# 셀 주소를 전달해 범위 내의 셀 얻기

import openpyxl as excel

# 워크북 열고 시트를 가져오기
book = excel.load_workbook('write_cellname.xlsx')
sheet = book.active

# 연속 데이터를 읽어서 출력하기
for row in sheet["B2":"D4"]:
    r = []
    for cell in row:
        r.append(cell.value)
    print(r)

print('--------------------')

