
# 워크시트의 데이터를 전부 읽어보기

import openpyxl as excel

#book = excel.load_workbook('input/monthly_sales.xlsx')
# 수식이 설정되어 있는 셀 있음. 계산이 끝난 값 얻으려면
book = excel.load_workbook('input/monthly_sales.xlsx', data_only=True)

sheet = book.active

# A3 에서 F9 범위의 셀을 가져오기
rows = sheet["A3":"F9"]
for row in rows:
    # 셀의 값을 리스트로 얻기
    values = [cell.value for cell in row]  # (2)
    # 리스트 출력
    print(values)

print('--------------------')
'''
위 (2) - 리스트 컴프리헨션 기법. 
이는 축약된 형태의 표현식이기 때문에 한눈에 이해하기 어려울 수 있다
다시 작성하면 다음과 같다
'''
for row in rows:
    # 셀의 값을 리스트에 저장
    values = []
    for cell in row:
        values.append(cell.value)
    # 리스트 출력
    print(values)