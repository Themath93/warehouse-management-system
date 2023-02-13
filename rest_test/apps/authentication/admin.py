# -*- encoding: utf-8 -*-
"""
Copyright (c) 2019 - present AppSeed.us
"""

from django.contrib import admin
from django.contrib.auth.admin import UserAdmin as BaseUserAdmin
from django.contrib.auth.models import User

from .models import UserDetail

# Define an inline admin descriptor for Employee model
# which acts a bit like a singleton
class UserDetailInline(admin.StackedInline):
    model = UserDetail
    can_delete = False
    verbose_name_plural = 'userdetail'

# Define a new User admin
class UserAdmin(BaseUserAdmin):
    inlines = (UserDetailInline,)

# Re-register UserAdmin
admin.site.unregister(User)
admin.site.register(User, UserAdmin)

# admin.site.register(User)
# Register your models here.
