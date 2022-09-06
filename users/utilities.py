from typing import Union
import openpyxl
import os
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
