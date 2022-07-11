from openpyxl import Workbook
from openpyxl.styles import Font,Alignment
from openpyxl.styles.borders import Border, Side
import xlsxwriter
import os

##1. Sheet 1번 List
title = list()
contents = list()

##2. Sheet 2번 List
query_list = list()
query_title = list()
lee = list()
del_str = '='
query_title_list = list()
title_1 = list()
title_2 = list()
con = list()
sheet2_query = list()

def pt_query_digest_slow(input,output):

    ##기존 파일 삭제
    filepath = output
    if os.path.exists(filepath):
        os.remove(filepath)

    lee22 = input
    ##LOG 파일 읽어오기
    slow_text = ''
    read_file = open(input, 'r', encoding="UTF-8")
    while True:
        line = read_file.readline()
        if not line: break
        slow_text += line
    read_file.close()

    ## Sheet1 번 RANK 부분 ) TOP10개의 문장을 List 형식 rank_excel에 넣음 ( RANK,Query1번까지 slow_text를 자름 ) )
    rank_1 = slow_text.find('Rank')
    rank_2 = slow_text.find('Query 1')
    RANK = slow_text[rank_1:rank_2]
    rank_list = RANK.split('#')
    del rank_list[1]
    rank_excel = rank_list[0:11]

    # Excel 생성
    workbook = xlsxwriter.Workbook(output)


    # list 문자열로 변환 함수
    def listToString(str_list):
        result = ""
        for s in str_list:
            result += s + " "
        return result.strip()

    ## 2번 Sheet Query Table Excel 만들어주는 함수
    def rankquery(lst, sheet_name, real_query):
        a = 2
        b = 0
        query = workbook.add_worksheet(sheet_name)
        title_format = workbook.add_format({'bold': True})
        title_format.set_bg_color('#a0a0a0')
        title_format.set_align('center')
        title_format.set_color('#ffffff')
        query.merge_range('C13:K24', sheet2_query[real_query])
        query.write(2, 2, "Attribute", title_format)
        query.write(2, 3, "pct", title_format)
        query.write(2, 4, "total", title_format)
        query.write(2, 5, "min", title_format)
        query.write(2, 6, "max", title_format)
        query.write(2, 7, "avg", title_format)
        query.write(2, 8, "95%", title_format)
        query.write(2, 9, "stddev", title_format)
        query.write(2, 10, "median", title_format)
        for i in range(0, 63):
            format = workbook.add_format({'border': 1})
            b = b + 1
            query.write(a + 1, b + 1, str(lst[i]), format)
            if i % 9 == 8:
                a = a + 1
                b = 0

    ## SHEET 1번 Excel 값 넣기
    for a in range(0, 10):
        list = rank_excel[a].split(' ')
        list = ' '.join(list).split()
        ##1번줄 list[0],list[1]+list[2],list[3]+list[4],list[5],list[6]
        if a == 0:
            title.append(list[0])
            title.append(list[1] + list[2])
            title.append(list[3] + list[4])
            title.append('%')
            title.append(list[5])
            title.append(list[6])
            title.append("SQL Command")
            sheet1 = workbook.add_worksheet('Slow_RANK')
            title_format = workbook.add_format({'bold': True})
            title_format.set_bg_color('#a0a0a0')
            title_format.set_align('center')
            title_format.set_color('#ffffff')
            title_format.set_border(1)
            for i in range(0, len(title)):
                sheet1.write(1, i + 2, title[i], title_format)
            ###여기에 FOR문으로 Excel에 title데이터 넣을 예정 ( 셀병합등이 필요해 보인다. Responsetime은 2개의 cell )
        if a > 0:
            contents.append(list.pop(0))
            contents.append(list.pop(0))
            contents.append(list.pop(0))
            contents.append(list.pop(0))
            contents.append(list.pop(0))
            contents.append(list.pop(0))
            del (list[0])
            contents.append(listToString(list))
            contents_format = workbook.add_format({'bold': False})
            contents_format.set_border(1)
            for b in range(0, len(contents)):
                sheet1.write(a + 1, b + 2, str(contents[b]), contents_format)
            contents.clear()
        ##2번줄 list[0],list[1],list[2]+list[3],list[4],list[5]+list[6],나머지
        # print(list)
    sheet1.set_column('D:D', 50)
    sheet1.set_column('I:I', 100)
    sheet1.set_column('E:E', 15)

    # SHEET 2번
    for i in range(1, 11):
        number = "Query %d" % i
        number2 = "Query %s" % str(i + 1)
        first = slow_text.find(number)
        two = slow_text.find(number2)
        query_list.append(slow_text[first:two])

    for i in range(1, 11):
        number = "Attribute"
        number2 = "String"
        query_list[i - 1] = query_list[i - 1].replace('#', '')
        mm = query_list[i - 1]
        query_title_1 = mm.find(number)
        query_title_2 = mm.find(number2)
        query_title.append(mm[query_title_1:query_title_2])

    three = 0
    for i in range(1, 11):
        number = "10s+"
        number2 = "Query %s" % str(i + 1)
        first = slow_text.find(number, three)
        two = slow_text.find(number2)
        three = first + 1
        sheet2_query.append(slow_text[first:two])
        print(sheet2_query[i - 1])

    for i in range(0, 8):
        query_title_list.append(str(query_title[i]).split('\n'))
    print(query_title_list)

    for i in range(0, len(query_title_list)):
        del (query_title_list[i][0])

    print(query_title_list)

    for a in range(0, len(query_title_list)):
        # print('\n')
        for b in range(0, 8):
            lee = (query_title_list[a][b].split(' '))
            lee = ' '.join(lee).split()
            title_1.append(lee)

    for a in range(0, len(title_1)):
        for b in range(0, len(title_1[a])):
            title_2.append(title_1[a][b])

    print(title_2)

    remove_content = "==="
    for word in title_2:
        if remove_content in word:
            title_2.remove(word)
    for word in title_2:
        if remove_content in word:
            title_2.remove(word)

    print(title_1)
    print(title_2)

    for a in range(0, 7):
        con.append(title_2.pop(0))
        con.append(title_2.pop(0))
        con.append(title_2.pop(0))
        con.append("-")
        con.append("-")
        con.append("-")
        con.append("-")
        con.append("-")
        con.append("-")
        for i in range(0, 6):
            con.append(title_2.pop(0) + title_2.pop(0))
            con.append(title_2.pop(0))
            con.append(title_2.pop(0))
            con.append(title_2.pop(0))
            con.append(title_2.pop(0))
            con.append(title_2.pop(0))
            con.append(title_2.pop(0))
            con.append(title_2.pop(0))
            con.append(title_2.pop(0))

    query1 = con[0:63]
    query2 = con[63:126]
    query3 = con[126:189]
    query4 = con[189:252]
    query5 = con[252:315]

    rankquery(query1, "Query1", 0)
    rankquery(query2, "Query2", 1)
    rankquery(query3, "Query3", 2)
    rankquery(query4, "Query4", 3)
    rankquery(query5, "Query5", 4)

    workbook.close()


pt_query_digest_slow("C:/Users/qwerr/Desktop/python/mysql/lee.log","C:/Users/qwerr/Desktop/python/mysql/kovea1.xlsx")