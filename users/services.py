from users.models import CustomUser
from django.db.models import QuerySet

class UserService:
    """
    Класс UserService
    Этот класс используется для извлечения пользовательских данных из база данных

    Методы:
       get_users_list
    """

    @classmethod
    def get_users_list(cls, fields: list, count: int = 10) -> QuerySet:
        """
        Метод: get_users_list
        отфильтровывает с нужными полями пользователей
        """
        return CustomUser.objects.filter(is_superuser=False, is_active=True).values_list(*fields)[:count]
