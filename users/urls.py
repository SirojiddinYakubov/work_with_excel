from django.urls import path
from django.views.generic import TemplateView
from users import views

app_name = 'users'

urlpatterns = [
    path('', TemplateView.as_view(template_name='index.html')),
    path('create-users/', views.create_users),
    path('write-users-to-file/', views.write_users_to_file),
    path('users-salary-report-to-file/', views.users_salary_report_to_file, name='users_salary_report_to_file'),
]
