from django.contrib import messages
from django.http import HttpResponse, FileResponse
from django.shortcuts import redirect
from django.urls import reverse_lazy

from users.models import CustomUser
from . import utilities
from .services import UserService
from django.utils.timezone import localtime


def create_users(request):
    datas = utilities.users_data_parse_from_excel(file_name='users.xlsx')
    users = CustomUser.objects.bulk_create([CustomUser(email=data['email'],
                                                       last_name=data['last_name'],
                                                       first_name=data['first_name'],
                                                       person_id=data['person_id'],
                                                       gender=data['gender'],
                                                       salary=data['salary'],
                                                       birthday=data['birthday']) for data in datas])

    return HttpResponse(f"{len(users)} objects created!")


def write_users_to_file(request):
    users_data = UserService.get_users_list(
        fields=['id', 'person_id', 'email', 'last_name', 'first_name', 'salary', 'birthday'])
    utilities.write_users_data_to_excel(file_name='new_users.xlsx', sheet_name='users_data', users_data=users_data)
    return HttpResponse(f"{len(users_data)} objects were written to a file!")


def users_salary_report_to_file(request):
    users = UserService.get_users_list(fields=['email', 'last_name', 'first_name', 'salary'], count=7)
    file = utilities.users_salary_report_build(users=users)
    file_name = 'users_salary_report_' + localtime().strftime('%d-%m-%Y_%H-%M-%S') + '.xlsx'
    return FileResponse(
        file,
        as_attachment=True,
        filename=file_name,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


def send_report_file_with_email(request):
    users = UserService.get_users_list(fields=['email', 'last_name', 'first_name', 'salary'], count=7)
    file = utilities.users_salary_report_file_build_for_email_send(users=users)
    superusers = UserService.get_superusers_list()
    server = utilities.SmtpServer(recipients=superusers, file=file)
    server.send()
    return redirect(reverse_lazy('users:index_page'))
