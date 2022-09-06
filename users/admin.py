from django.contrib import admin
from users.models import CustomUser


@admin.register(CustomUser)
class CustomUserAdmin(admin.ModelAdmin):
    list_display = ['id', 'email', 'first_name', 'last_name', 'birthday',
                    'is_active', 'is_superuser', 'is_staff', 'date_joined', 'last_login']
    list_display_links = ['email', 'first_name', 'last_name', ]
    list_filter = ['is_active', 'is_superuser', 'is_staff']
    search_fields = ['first_name', 'last_name', 'email']
    save_on_top = True