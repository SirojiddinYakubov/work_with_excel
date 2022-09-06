from django.contrib.auth.models import AbstractUser
from django.core.validators import MinValueValidator, MaxValueValidator
from django.db import models
from django.utils.translation import gettext_lazy as _

from .managers import CustomUserManager


class CustomUser(AbstractUser):
    GENDER_MALE = 0
    GENDER_FEMALE = 1
    GENDER_CHOICES = [(GENDER_MALE, 'Male'), (GENDER_FEMALE, 'Female')]

    username = None
    email = models.EmailField(_('email address'), unique=True)

    USERNAME_FIELD = 'email'
    REQUIRED_FIELDS = []

    objects = CustomUserManager()

    person_id = models.BigIntegerField(null=True,
                                       validators=[MinValueValidator(10000000000000),
                                                   MaxValueValidator(99999999999999)])
    gender = models.IntegerField(choices=GENDER_CHOICES, null=True, )
    salary = models.DecimalField(decimal_places=2, max_digits=9, null=True, )
    birthday = models.DateField(blank=True, null=True)

    def __str__(self):
        return self.email
