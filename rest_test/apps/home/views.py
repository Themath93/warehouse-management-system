# -*- encoding: utf-8 -*-
"""
Copyright (c) 2019 - present AppSeed.us
"""

from django import template
from django.contrib.auth.decorators import login_required
from django.http import HttpResponse, HttpResponseRedirect
from django.template import loader
from django.urls import reverse
from django.shortcuts import render
from apps.authentication.models import *
from django.db.models import Q
from django.core.mail import send_mail
from core.infra import *

import json
import pandas as pd
import datetime as dt

@login_required(login_url="/login/")
def index(request):
    context = {'segment': 'index'}

    html_template = loader.get_template('home/index.html')
    return HttpResponse(html_template.render(context, request))


@login_required(login_url="/login/")
def pages(request):
    context = {}
    # All resource paths end in .html.
    # Pick out the html file name from the url. And load that template.
    try:

        load_template = request.path.split('/')[-1]

        if load_template == 'admin':
            return HttpResponseRedirect(reverse('admin:index'))
        context['segment'] = load_template

        html_template = loader.get_template('home/' + load_template)
        return HttpResponse(html_template.render(context, request))

    except template.TemplateDoesNotExist:

        html_template = loader.get_template('home/page-404.html')
        return HttpResponse(html_template.render(context, request))

    except:
        html_template = loader.get_template('home/page-500.html')
        return HttpResponse(html_template.render(context, request))



@login_required(login_url="/login/")
def req_parts(request):
    prod_pose = ProdPose.objects.all().using('dw')
    products = Products.objects.all().using('dw')

    # 매일 재고는 그날 그날 업데이트 분을 줘야 한다. 
    # 현재 개발단계에서는 그냥 특정일자 재고리스트를 참고하도록하자
    tmp_date = '2023-02-01'
    total_stock_db = list(TotalStock.objects.filter(Q(std_day=tmp_date)).values())
    json_ts = json.dumps(total_stock_db,ensure_ascii=False)

    return render(request, 'home/select_part.html',{'data_ts':json_ts,'datas_pose':prod_pose,'datas_products':products,'user':request.user})

@login_required(login_url="/login/")
def svc_process(request):

    if request.method =='POST':
        user_info = request.user.userdetail
        std_day = str(dt.datetime.today().date())
        req_user = request.user.username
        fe_initial = user_info.subinventory
        contact = user_info.contact_number_1
        sub_contact = user_info.contact_number_2
        req_date = request.POST.get('req_date')
        del_method = request.POST.get('del_method')
        req_time = request.POST.get('req_time')
        is_urgent = request.POST.get('is_urgent')
        is_return = request.POST.get('is_return')
        address = request.POST.get('search_address') + " " + request.POST.get('specific_address')
        recipient = 'need_fill'
        del_instruction = 'need_fill'
        state = 'requested'

        ### JSON 생성 ###
        cols = ['part_no','serial_no','qty','currnet_stock','transfer_to','BIN']
        data = []
        for key, value in request.POST.items():
            rows=[]
            if 'input_qty' in key:
                in_ts_key = int(key.split('_')[-1])
                req_part_db = TotalStock.objects.filter(Q(ts_key=in_ts_key))
                rows.append(list(req_part_db.values('article_number'))[0]['article_number'])
                rows.append('')
                rows.append(int(value))
                rows.append(list(req_part_db.values('subinventory'))[0]['subinventory'])
                rows.append(request.POST.get('to_sub_'+str(in_ts_key)))
                rows.append(list(req_part_db.values('bin_cur'))[0]['bin_cur'])
                tmp = dict(zip(cols,rows))
                data.append(tmp)

        # 가장 먼저해야하는 것은 SYSTEM_STOCK에 내용 반영 및 내용 전송
        # 임시사용 db 
        sys_db = SystemStock.objects.filter(std_day='2023-02-01').using('dw')
        # sys_db = SystemStock.objects.filter(std_day=std_day).using('dw')
        print(data)
        for d in data:
            search_db = sys_db.filter(Q(article_number=d['part_no']) and Q(subinventory=d['currnet_stock']))
            print(search_db.values('quantity'))
            



        # # SERVICE_REQUEST 파트요청 접수..
        # key_contain = std_day.replace('-','')[2:]
        # try:
        #     key_count = len(list(ServiceReqeust.objects.filter(svc_key__icontains=key_contain).using('dw').values())) + 1 
        # except:
        #     key_count = 1

        # req_db = ServiceReqeust(

        #     svc_key='SVC_'+fe_initial.split('_')[1]+'_'+key_contain+str(key_count),
        #     fe_name = request.user.first_name + request.user.last_name,
        #     fe_initial = fe_initial,
        #     req_day = req_date,
        #     req_time = req_time,
        #     address = address,
        #     del_met = del_method,
        #     is_return = is_return,
        #     is_urgent = is_urgent,
        #     recipient = recipient,
        #     contact = contact,
        #     contact_sub = sub_contact,
        #     del_instruction = del_instruction,
        #     parts = data,
        #     std_day = std_day,
        #     timeline = create_db_timeline(),
        #     state = state

        #     )
        # req_db.save(using='dw')

        # # 요청메일 전송
        # res = {
        #     'meta':{
        #         'desc':'service_part_reqeust',

        #         'cols':{
        #             'fe_name':'fe_이름',
        #             'fe_initial':'trunkstock_id',
        #             'req_day':'요청일',
        #             'req_time':'요청시간',
        #             'address':'주소',
        #             'del_met':'배송방법',
        #             'is_return':'왕복배송여부',
        #             'recipient':'수령인',
        #             'is_urgent':'긴급여부',
        #             'parts':'요청파트',
        #             'del_instruction':'배송요청사항',
        #             'contact':'전화번호'

        #         },
        #         'req_info':{
        #             'fe_name':req_user,
        #             'fe_initial':fe_initial,
        #             'req_day':req_date,
        #             'req_time':req_time,
        #             'address':address,
        #             'contact':contact,
        #             'del_met':del_method,
        #             'is_urgent':is_urgent,
        #             'is_return':is_return,
        #             'recipient':'need_update',
        #             'del_instruction':'need_update'
                    

        #         },
        #         'std_day': std_day
        #     },
        #     'parts':data
        # }
        # req_json = json.dumps(res,ensure_ascii=False)
        # __sending_outlook_mail(request, key_count, std_day, fe_initial, req_json)
        return HttpResponse("받았어요")

def __sending_outlook_mail(request, daily_count, std_day, fe_initial, req_json):
    send_mail(
            'SVC_' +fe_initial.split('_')[1]+'_' +std_day.replace('-','')[2:]+str(daily_count),
            req_json,
            'deyoon@outlook.kr',
            ['deyoon@outlook.kr',request.user.email],
            fail_silently=False,
        )

