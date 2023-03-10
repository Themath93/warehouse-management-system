# -*- encoding: utf-8 -*-
"""
Copyright (c) 2019 - present AppSeed.us
"""

# Create your views here.
from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login

from core.infra import create_db_timeline

from .forms import LoginForm, SignUpForm
from django.contrib.auth.decorators import login_required
from .models import *
from django.contrib.auth.models import User
from django.template import loader
from django.http import HttpResponse
from django.core.mail import send_mail
from django.forms.models import model_to_dict

import pandas as pd
import json
import datetime as dt

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
    if request.method == 'POST':
        std_day = str(dt.datetime.today().date())
        fe_name = request.user.first_name + request.user.last_name,
        subinventory = request.POST['subinventory']
        user_email = request.POST['email']
        del_method = request.POST['del_method']
        courier_nm = request.POST['courier_nm']
        contact_1 = request.POST['contact']
        address = request.POST['address_1'] + " " + request.POST['address_2']
        req_date = request.POST['req_date']
        req_time = request.POST['req_time']
        del_instruction = request.POST['del_instruction']


        cols = ['part_no','serial_or_IR_no','qty','currnet_stock','transfer_to','BIN']
        data=[]
        for k,v in request.POST.items():
            if v == 'True':
                rows=[]
                if 'ts' in k:
                    ts_key = k.split('_')[2]
                    ts_obj = TotalStock.objects.get(ts_key=ts_key)
                    try:
                        bin = SvcBin.objects.using('dw').get(article_number=ts_obj.article_number).bin
                    except:
                        bin = 'NO_BIN'
                    qty =int(request.POST[f"parts_qty_{ts_key}"])
                    rows.append(ts_obj.article_number)
                    rows.append('')
                    rows.append(qty)
                    rows.append(subinventory)
                    rows.append('KR_SERV01')
                    rows.append(bin)
                    tmp = dict(zip(cols,rows))
                    data.append(tmp)

                    # part_db적용
                    total_db = TotalStock.objects.filter(std_day='2023-02-01')
                    try : # KR_SERV01에 해당 파트가 있을경우
                        search_db_kr = total_db.filter(article_number=ts_obj.article_number).get(subinventory='KR_SERV01')
                        search_db_qty_kr = search_db_kr.quantity
                        update_qty_kr = search_db_qty_kr + qty
                        search_db_kr.quantity=update_qty_kr
                        search_db_kr.save()

                    except: # KR_SERV01에 해당 파트가 없을 경우
                        TotalStock(
                            article_number = ts_obj.article_number,
                            subinventory = 'KR_SERV01',
                            quantity = qty,
                            country = '',
                            prod_centre = '',
                            prod_group = '',
                            description = '',
                            prod_status_type = '',
                            bin_cur = bin,
                            std_day = std_day,
                            state = 'GOOD_WR',
                            state_time = str(dt.datetime.now()).split('.')[0]
                    ).save()



                elif 'tool' in k:
                    tool_index = k.split('_')[2]
                    tool_obj = SvcTool.objects.using('dw').get(tool_index=tool_index)
                    rows.append(tool_obj.tool_nm)
                    rows.append('TOOL')
                    rows.append(subinventory)
                    rows.append('KR_SERV01')
                    rows.append('')
                    tmp = dict(zip(cols,rows))
                    data.append(tmp)
                    # tool_db 적용
                    tool_obj.on_hand = 'KR_SERV01'
                    tool_obj.state = 'GOOD_WR'
                    tool_obj.ship_date = std_day
                    tool_obj.save(using='dw')


        # SERVICE_REQUEST 파트요청 접수..
        key_contain = std_day.replace('-','')[2:]
        try:
            key_count = len(list(ServiceRequest.objects.filter(svc_key__icontains=key_contain).using('dw').values())) + 1 
        except:
            key_count = 1
            
        svc_key = 'RETURN_'+subinventory.split('_')[1]+'_'+key_contain+str(key_count)
        
        # DB적용
        ServiceRequest(
            svc_key=svc_key,
            fe_name = request.user.first_name + request.user.last_name,
            fe_initial = subinventory,
            req_day = req_date,
            req_time = req_time,
            address = address,
            del_met = del_method,
            is_return = 'is_return',
            is_urgent = '',
            recipient = '',
            contact = contact_1,
            contact_sub = '',
            del_instruction = del_instruction,
            parts = data,
            std_day = std_day,
            timeline = create_db_timeline(),
            state = 'requested'
        ).save(using='dw')


      # 요청메일 전송
        res = {
            'meta':{
                'desc':'service_part_reqeust',

                'cols':{
                    'fe_name':'fe_이름',
                    'fe_initial':'trunkstock_id',
                    'req_day':'요청일',
                    'req_time':'요청시간',
                    'address':'주소',
                    'del_met':'배송방법',
                    'is_return':'왕복배송여부',
                    'recipient':'수령인',
                    'is_urgent':'긴급여부',
                    'parts':'요청파트',
                    'del_instruction':'배송요청사항',
                    'contact':'전화번호'

                },
                'req_info':{
                    'fe_name':fe_name,
                    'fe_initial':subinventory,
                    'req_day':req_date,
                    'req_time':req_time,
                    'address':address,
                    'contact':contact_1,
                    'del_met':del_method,
                    'is_urgent':'',
                    'is_return':'is_return',
                    'recipient':'warehouse',
                    'del_instruction':del_instruction
                    

                },
                'std_day': std_day
            },
            'parts':data
        }
        req_json = json.dumps(res,ensure_ascii=False)
        print(req_json)
        __sending_outlook_mail(request, key_count, std_day, subinventory, req_json)    

        


        # return HttpResponse('반납요청 완료')
        return redirect(f'/profile/{svc_key}') 

@login_required(login_url="/login/")
def case_detail(request,case_id):
    each_case = ServiceRequest.objects.using('dw').filter(svc_key=case_id)
    df_parts= pd.DataFrame(list(each_case.values()))['parts']
    parts_str = df_parts[0].replace("'",'"')
    dict_parts = json.loads(parts_str)
    return render(request,'home/request_detail.html',{'dict_parts':dict_parts,'req_detail':each_case})


def __sending_outlook_mail(request, daily_count, std_day, fe_initial, req_json):
    req_json = json.loads(req_json)
    info_dict = req_json['meta']['req_info']
    info_dict['parts'] = req_json['parts']

    html_message = loader.render_to_string(
        'home/request_mail_form.html',
        context=info_dict

    )
    subject='RETURN_' +fe_initial.split('_')[1]+'_' +std_day.replace('-','')[2:]+str(daily_count)
    message = "test"
    from_email='deyoon@outlook.kr'
    to_list=['deyoon@outlook.kr',request.user.email]
    send_mail(subject,message,from_email,to_list,fail_silently=False,html_message=html_message)