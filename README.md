# python-mysql-slow


## Python으로 pt-query-digest의 output file을 Excel로 바꿔준다 (Top5 Query)


1. pt-query-digest로 mysql slow query 분석


2. python을 이용하여 해당 file excel로 변환
- 함수로 만들어놨기때문에 앞에는 query_log 파일 / 뒤에는 Excel 이름


pt_query_digest_slow("C:/Users/qwerr/Desktop/python/mysql/lee.log","C:/Users/qwerr/Desktop/python/mysql/kovea1.xlsx") 

