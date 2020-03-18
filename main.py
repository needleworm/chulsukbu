#-*-coding:euc-kr
"""
Author : Byunghyun Ban
GitHub : https://github.com/needleworm
Book : 일반인을 위한 업무 자동화
Last Modification : 2020.02.13.
"""
import time
import pyexcel as px
import sys


# 작업 시작 메시지를 출력합니다.
print("Process Start")

# 시작 시점의 시간을 기록합니다.
start_time = time.time()

# 엑셀 파일 이름을 입력받습니다.
filename = sys.argv[1]

# 엑셀 파일을 book 형태로 불러옵니다.
book = px.get_book(file_name=filename)

# 각 반별 정보만 뽑아옵니다
class1_1 = book[1].get_array()
class1_2 = book[2].get_array()
class1_3 = book[3].get_array()
class1_4 = book[4].get_array()
class1_5 = book[5].get_array()
class1_6 = book[6].get_array()

# 이걸 묶어서 하나의 리스트로 만듭니다.
classes_book = [class1_1, class1_2, class1_3, class1_4, class1_5, class1_6]

classes = []
# 1번 시트를 읽어와 과목명만 뽑아옵니다.
book_array = book[0].get_array()
for row in book_array:
    for col in row:
        if col in "월화수목금":
            continue
        if col and col not in classes:
            splt = col.split("\n")
            col = "".join(splt)
            classes.append(col)

# 요일
days = ["월", "화", "수", "목", "금"]

# 각 과목별 출석부 샘플을 저장할 딕셔너리를 만든다
result_book = {}
# 각 과목별 출석부에 삽입된 학생 이름을 카운트하기 위한 딕셔너리를 만든다
dictionary_for_book = {}

# 각 과목별로 출석부를 만들어 리스트에 집어넣습니다.
for el in classes:
    new_book = {}
    new_dictionary = {}
    line_1 = ["" for i in range(19)]
    line_2 = ["", "", "  월  일"] + ["/" for i in range(16)]
    line_3 = ["", "", "  교  시"] + ["8" for i in range(16)]
    line_4 = ["", "", "  과  목"] + [" " for i in range(16)]
    line_5 = ["연번", "반", "번호", "학생명, 강사명"] + ["" for i in range(16)]
    line_template = ["" for i in range(18)]

    for day in days:
        # 새 시트가 들어갈 리스트
        new_sheet = []
        header_line = "2020학년도 방과후학교 출석부(" + el + "-"+ day + "요일)"
        header = [header_line] + ["" for i in range(18)]
        new_sheet.append(line_1)
        new_sheet.append(line_2)
        new_sheet.append(line_3)
        new_sheet.append(line_4)
        new_sheet.append(line_5)
        new_dictionary[day] = len(new_sheet) + 1
        for i in range(37):
            new_line = [str(i+1)] + line_template
            new_sheet.append(new_line)
        new_book[day] = new_sheet

    result_book[el] = new_book
    dictionary_for_book[el] = new_dictionary


# 이제 모든 학생들을 한명씩 읽어와 출석부에 삽입합니다.
for cls in classes_book:
    # 첫 세줄 버리고 한 줄씩 읽어오기
    for line in cls[3:]:
        class_number = line[1]
        student_id = line[2]
        name = line[3]

        if not name:
            continue

        # 과목들 뽑아오기
        enrolled_classes = line[4:]
        for i, cls_name in enumerate(enrolled_classes):
            day = days[i]
            # 과목 뽑아오기
            classbook = result_book[cls_name]
            # 요일별 출석부 뽑기
            daybook = classbook[day]
            # 반 기재하기
            daybook[dictionary_for_book[cls_name][day]][1] = class_number
            # 번호 기재하기
            daybook[dictionary_for_book[cls_name][day]][2] = student_id
            # 학생명 기재하기
            daybook[dictionary_for_book[cls_name][day]][3] = name
            dictionary_for_book[cls_name][day] += 1

for cls_name in classes:
    class_sheet = result_book[cls_name]
    result = px.get_book(bookdict=class_sheet)
    result.save_as("출석부_" + cls_name + ".xlsx")

# 작업 종료 메시지를 출력합니다.
print("Process Done.")

# 작업에 총 몇 초가 걸렸는지 출력합니다.
end_time = time.time()
print("The Job Took " + str(end_time - start_time) + " seconds.")
