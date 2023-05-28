# -*- encoding: utf-8 -*-
"""
Copyright (c) 2019 - present AppSeed.us
"""

from django.urls import path, re_path
from apps.home import views

app_name = 'home'

urlpatterns = [

    # The home page
    path('', views.index, name='home'),
    path('kakao_test', views.kakao_test, name='kakao'),
    path('kafka_test', views.kafka_test, name='kafka_test'),
    path('request_parts/', views.req_parts, name='req_parts'),
    path('request_parts/process', views.svc_process, name='svc_process'),
    # Matches any html file
    re_path(r'^.*\.*', views.pages, name='pages'),

]
