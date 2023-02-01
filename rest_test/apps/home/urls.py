# -*- encoding: utf-8 -*-
"""
Copyright (c) 2019 - present AppSeed.us
"""

from django.urls import path, re_path
from apps.home import views

# app_name = 'home'

urlpatterns = [

    # The home page
    path('', views.index, name='home'),
    path('request_parts/', views.req_parts, name='req_parts'),
    # Matches any html file
    re_path(r'^.*\.*', views.pages, name='pages'),

]
