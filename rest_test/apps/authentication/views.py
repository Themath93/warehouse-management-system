# -*- encoding: utf-8 -*-
"""
Copyright (c) 2019 - present AppSeed.us
"""

# Create your views here.
from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login
from .forms import LoginForm, SignUpForm
from django.contrib.auth.decorators import login_required
from .models import *
from django.contrib.auth.models import User
from django.http import HttpResponse
from django.forms.models import model_to_dict

import pandas as pd
import json

def login_view(request):
    form = LoginForm(request.POST or None)

    msg = None

    if request.method == "POST":

        if form.is_valid():
            username = form.cleaned_data.get("username")
            password = form.cleaned_data.get("password")
            user = authenticate(username=username, password=password)
            if user is not None:
                login(request, user)
                return redirect("/")
            else:
                msg = 'Invalid credentials'
        else:
            msg = 'Error validating the form'

    return render(request, "accounts/login.html", {"form": form, "msg": msg})


def register_user(request):
    msg = None
    success = False

    if request.method == "POST":
        form = SignUpForm(request.POST)
        if form.is_valid():
            form.save()
            username = form.cleaned_data.get("username")
            raw_password = form.cleaned_data.get("password1")
            user = authenticate(username=username, password=raw_password)

            msg = 'User created successfully.'
            success = True

            # return redirect("/login/")

        else:
            msg = 'Form is not valid'
    else:
        form = SignUpForm()

    return render(request, "accounts/register.html", {"form": form, "msg": msg, "success": success})

@login_required(login_url="/login/")
def profile(request):
    #이름,이메일 데이터 보여주기
    if request.user.is_authenticated:
        get_request_type = request.GET.get('request_type')
        user = User.objects.get(id=request.user.id)
        uesr_detail = UserDetail.objects.filter(user_id=user.id)
        user_stock_id = uesr_detail[0].subinventory
        my_reqeusts = ServiceRequest.objects.filter(fe_initial=user_stock_id).using('dw')
        print(get_request_type=='all')
        if (get_request_type != None) &(get_request_type != 'all'):
            my_reqeusts=my_reqeusts.filter(state=get_request_type)
        elif get_request_type == 'all':
            my_reqeusts = my_reqeusts

        return render(request,'home/user.html',{'forms':user,'my_requests':my_reqeusts})

@login_required(login_url="/login/")
def my_stock(request,subinventory):
    if request.method == 'GET':
        user = User.objects.get(id=request.user.id)
        my_part = TotalStock.objects.filter(subinventory=subinventory)
        my_tool = SvcTool.objects.filter(on_hand=subinventory).using('dw')
        return render(request,'home/mystocks.html',{'my_parts':my_part,'my_tools':my_tool,'forms':user})
    return HttpResponse("내재고")

@login_required(login_url="/login/")
def send_return_request(request):
    posts= request.POST
    print(posts)
    return HttpResponse("응답완료")

@login_required(login_url="/login/")
def case_detail(request,case_id):
    each_case = ServiceRequest.objects.using('dw').filter(svc_key=case_id)
    df_parts= pd.DataFrame(list(each_case.values()))['parts']
    parts_str = df_parts[0].replace("'",'"')
    dict_parts = json.loads(parts_str)
    return render(request,'home/request_detail.html',{'dict_parts':dict_parts,'req_detail':each_case})