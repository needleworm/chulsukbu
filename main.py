#-*-coding:euc-kr
"""
Author : Byunghyun Ban
GitHub : https://github.com/needleworm
Book : �Ϲ����� ���� ���� �ڵ�ȭ
Last Modification : 2020.02.13.
"""
import time
import pyexcel as px
import sys


# �۾� ���� �޽����� ����մϴ�.
print("Process Start")

# ���� ������ �ð��� ����մϴ�.
start_time = time.time()

# ���� ���� �̸��� �Է¹޽��ϴ�.
filename = sys.argv[1]

# ���� ������ book ���·� �ҷ��ɴϴ�.
book = px.get_book(file_name=filename)

# �� �ݺ� ������ �̾ƿɴϴ�
class1_1 = book[1].get_array()
class1_2 = book[2].get_array()
class1_3 = book[3].get_array()
class1_4 = book[4].get_array()
class1_5 = book[5].get_array()
class1_6 = book[6].get_array()

# �̰� ��� �ϳ��� ����Ʈ�� ����ϴ�.
classes_book = [class1_1, class1_2, class1_3, class1_4, class1_5, class1_6]

classes = []
# 1�� ��Ʈ�� �о�� ����� �̾ƿɴϴ�.
book_array = book[0].get_array()
for row in book_array:
    for col in row:
        if col in "��ȭ�����":
            continue
        if col and col not in classes:
            splt = col.split("\n")
            col = "".join(splt)
            classes.append(col)

# ����
days = ["��", "ȭ", "��", "��", "��"]

# �� ���� �⼮�� ������ ������ ��ųʸ��� �����
result_book = {}
# �� ���� �⼮�ο� ���Ե� �л� �̸��� ī��Ʈ�ϱ� ���� ��ųʸ��� �����
dictionary_for_book = {}

# �� ���񺰷� �⼮�θ� ����� ����Ʈ�� ����ֽ��ϴ�.
for el in classes:
    new_book = {}
    new_dictionary = {}
    line_1 = ["" for i in range(19)]
    line_2 = ["", "", "  ��  ��"] + ["/" for i in range(16)]
    line_3 = ["", "", "  ��  ��"] + ["8" for i in range(16)]
    line_4 = ["", "", "  ��  ��"] + [" " for i in range(16)]
    line_5 = ["����", "��", "��ȣ", "�л���, �����"] + ["" for i in range(16)]
    line_template = ["" for i in range(18)]

    for day in days:
        # �� ��Ʈ�� �� ����Ʈ
        new_sheet = []
        header_line = "2020�г⵵ ������б� �⼮��(" + el + "-"+ day + "����)"
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


# ���� ��� �л����� �Ѹ� �о�� �⼮�ο� �����մϴ�.
for cls in classes_book:
    # ù ���� ������ �� �پ� �о����
    for line in cls[3:]:
        class_number = line[1]
        student_id = line[2]
        name = line[3]

        if not name:
            continue

        # ����� �̾ƿ���
        enrolled_classes = line[4:]
        for i, cls_name in enumerate(enrolled_classes):
            day = days[i]
            # ���� �̾ƿ���
            classbook = result_book[cls_name]
            # ���Ϻ� �⼮�� �̱�
            daybook = classbook[day]
            # �� �����ϱ�
            daybook[dictionary_for_book[cls_name][day]][1] = class_number
            # ��ȣ �����ϱ�
            daybook[dictionary_for_book[cls_name][day]][2] = student_id
            # �л��� �����ϱ�
            daybook[dictionary_for_book[cls_name][day]][3] = name
            dictionary_for_book[cls_name][day] += 1

for cls_name in classes:
    class_sheet = result_book[cls_name]
    result = px.get_book(bookdict=class_sheet)
    result.save_as("�⼮��_" + cls_name + ".xlsx")

# �۾� ���� �޽����� ����մϴ�.
print("Process Done.")

# �۾��� �� �� �ʰ� �ɷȴ��� ����մϴ�.
end_time = time.time()
print("The Job Took " + str(end_time - start_time) + " seconds.")
