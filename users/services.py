from users.models import CustomUser


class UserService:
    @classmethod
    def get_users_list(cls, fields):
        return CustomUser.objects.filter(is_superuser=False, is_active=True).values_list(*fields)
