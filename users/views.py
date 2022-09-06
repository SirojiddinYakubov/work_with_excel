from django.http import HttpResponse
from django.shortcuts import render
from . import utilities
from users.models import CustomUser


def create_users(request):
    datas = utilities.users_data_parse_from_excel(file_name='users.xlsx')
    users = CustomUser.objects.bulk_create([CustomUser(email=data['email'],
                                                       last_name=data['last_name'],
                                                       first_name=data['first_name'],
                                                       person_id=data['person_id'],
                                                       gender=data['gender'],
                                                       salary=data['salary'],
                                                       birthday=data['birthday']) for data in datas])
    print(users)
    return HttpResponse(f"{len(users)} objects created!")
