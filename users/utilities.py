from typing import Union
import openpyxl
import os
from openpyxl.styles import Font, Alignment
from openpyxl.styles.colors import BLACK
from conf import settings
from datetime import datetime
from users.models import CustomUser


def get_file_path(file_name: str) -> str:
    return os.path.join(settings.BASE_DIR, file_name)


def filter_user_data(data: dict) -> bool:
    if len(str(data['person_id'])) != 14:
        return False
    # if len(data['first_name']) < 5:
    #     return False
    return all([bool(value) for value in data.values()])


str_to_date = lambda value: datetime.strptime(value, "%d.%m.%Y").date()
get_gender_value = lambda value: CustomUser.GENDER_MALE if value == 'erkak' else CustomUser.GENDER_FEMALE


def get_real_value(data: dict) -> dict:
    if data['birthday']:
        data['birthday'] = str_to_date(data['birthday'])
    if data['gender']:
        data['gender'] = get_gender_value(data['gender'])
    return data


def users_data_parse_from_excel(file_name: str) -> Union[list[dict], tuple[dict]]:
    file_path = get_file_path(file_name=file_name)
    book = openpyxl.open(filename=file_path, read_only=True)
    sheet = book.active

    datas = []
    keys = ['person_id', 'first_name', 'last_name', 'email', 'gender', 'salary', 'birthday']
    for row in sheet.iter_rows(min_col=2, max_col=8, min_row=2, max_row=sheet.max_row + 1):
        data = dict(zip(keys, (cell.value for cell in row)))
        datas.append(data)

    return list(map(get_real_value, filter(filter_user_data, datas)))


def write_users_data_to_excel(file_name: str, sheet_name: str, users_data: Union[tuple, list]):
    # file_path = get_file_path(file_name=file_name)
    # book = openpyxl.open(filename=file_path)
    #
    # if not sheet_name in book.sheetnames:
    #     book.create_sheet(sheet_name)
    # sheet = book[sheet_name]

    book = openpyxl.Workbook()
    sheet = book.active

    sheet.column_dimensions['B'].width = 25
    sheet.column_dimensions['C'].width = 30
    sheet.column_dimensions['D'].width = 15
    sheet.column_dimensions['E'].width = 15
    sheet.column_dimensions['F'].width = 15
    sheet.column_dimensions['G'].width = 15

    # set headers
    headers = ['ID', 'person_id', 'email', 'last_name', 'first_name', 'salary', 'birthday']
    for i_col, header in enumerate(headers, start=1):
        cell = sheet.cell(row=1, column=i_col, value=header)
        cell.font = Font(size=11, color=BLACK, bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for i_row, user_data in enumerate(users_data, start=2):
        for i_col, value in enumerate(user_data, start=1):
            cell = sheet.cell(row=i_row, column=i_col, value=value)

            if i_col == 6:
                cell.alignment = Alignment(horizontal='left', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='center', vertical='center')
    book.save(file_name)
