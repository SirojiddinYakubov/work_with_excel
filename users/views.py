from django.http import HttpResponse
from users.models import CustomUser
from . import utilities
from .services import UserService


def create_users(request):
    # datas = utilities.users_data_parse_from_excel(file_name='users.xlsx')
    # users = CustomUser.objects.bulk_create([CustomUser(email=data['email'],
    #                                                    last_name=data['last_name'],
    #                                                    first_name=data['first_name'],
    #                                                    person_id=data['person_id'],
    #                                                    gender=data['gender'],
    #                                                    salary=data['salary'],
    #                                                    birthday=data['birthday']) for data in datas])
    users_data = UserService.get_users_list(
        fields=['id', 'person_id', 'email', 'last_name', 'first_name', 'salary', 'birthday'])
    utilities.write_users_data_to_excel(file_name='new_users.xlsx', sheet_name='users_data', users_data=users_data)
    return HttpResponse(f"{len(users_data)} objects were written to a file!")
