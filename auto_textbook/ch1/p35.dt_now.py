# 1-5 파이썬에서 날짜ㆍ시간 계산하기

# 날짜ㆍ시간 관련한 처리를 하는 모듈
import datetime

# 현재 시각 구하기
t = datetime.datetime.now()
print(t)

print('--------------------')
# 모듈명을 생략할 수 있다
from datetime import datetime
t2 = datetime.now()
print(t2)

print('--------------------')
# 날짜ㆍ시간 을 특정 형식으로 출력
fmt = t2.strftime('%Y년%m월%d일 %H시%M분%S초')
print(fmt)
