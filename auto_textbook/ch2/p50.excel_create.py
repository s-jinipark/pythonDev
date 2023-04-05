# 2-2 기본적인 엑셀 데이터 다루기

# 라이브러리 불러오기
import openpyxl as excel

# 새 워크북 생성
book = excel.Workbook()

# 활성화된 워크시크 가져오기
sheet = book.active

# 셀 A1 에 값 입력
sheet["A1"] = "안녕하세요"

# 파일 저장
book.save('hello.xlsx')

print('--------------------')

# 워크북 열기
book2 = excel.load_workbook('hello.xlsx')

# 워크북에서 첫 번째 워크시트 가져오기
sheet2 = book2.worksheets[0]

# 시트에서 셀 A1 가져오기
cell = sheet2['A1']

# A1 데이터 화면에 출력
print(cell.value)