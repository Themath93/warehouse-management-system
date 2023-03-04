# -*- encoding: utf-8 -*-
"""
Copyright (c) 2019 - present AppSeed.us
"""

from django.urls import path
from .views import login_view, register_user, profile, case_detail
from django.contrib.auth.views import LogoutView


app_name = 'authentication'

urlpatterns = [
    path('login/', login_view, name="login"),
    path('profile/', profile, name='profile'),
    path('profile/<str:case_id>', case_detail, name='case_detail'),
    path('register/', register_user, name="register"),
    path("logout/", LogoutView.as_view(), name="logout")
]
