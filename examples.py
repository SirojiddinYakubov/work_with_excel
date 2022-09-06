# TODO Exceldagi ma'lumotlarni o'qish

import openpyxl
import string

book = openpyxl.open('users.xlsx', read_only=False)
sheet = book.worksheets[0]

datas = []
keys = ['person_id', 'first_name', 'last_name', 'email', 'gender', 'salary', 'birthday']
for row in sheet.iter_rows(min_col=2, max_col=8, min_row=2, max_row=10):
    data = dict(zip(keys, (cell.value for cell in row)))
    datas.append(data)

print(datas)


def filter_data(data):
    if len(str(data['person_id'])) != 14:
        return False
    if len(data['first_name']) < 5:
        return False
    return True


print(list(filter(filter_data, datas)))
