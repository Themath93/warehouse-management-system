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

import json
import pandas as pd

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
    products_db = Products.objects.all()
    prod_pose_db = ProdPose.objects.all()
    
    print(SystemStock.objects.all()[0].article_number.prod_group)
    # 매일 재고는 그날 그날 업데이트 분을 줘야 한다. 
    # 현재 개발단계에서는 그냥 특정일자 재고리스트를 참고하도록하자
    tmp_date = '2023-02-01'
    system_stock_db = list(SystemStock.objects.filter(Q(std_day=tmp_date)).values())
    json_prods_db =json.dumps(list(products_db.values()),ensure_ascii=False)
    json_sys_db = json.dumps(system_stock_db,ensure_ascii=False)
    json_sys = json.loads(json_sys_db)

    json_tmp_list =[]
    for data in json_sys:
        data['article_number'] = data.pop('article_number_id')
        json_tmp_list.append(data)
    json_sys_fin = json.dumps(json_tmp_list,ensure_ascii=False)

    # df_sys = pd.DataFrame(json_tmp_list)
    # df_sys.to_csv('df_sys.csv',encoding='utf-8')

    df_prods = pd.DataFrame(json.loads(json_prods_db))
    df_prods.to_csv('df_prods.csv',encoding='utf-8')

    return render(request, 'home/select_part.html',{'datas_pose':prod_pose_db,'datas_products':products_db,'json_datas_sys':json_sys_fin,'json_datas_prods':json_prods_db})