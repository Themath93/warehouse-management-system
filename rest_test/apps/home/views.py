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
    prod_pose = ProdPose.objects.all().using('dw')
    products = Products.objects.all().using('dw')
    # 매일 재고는 그날 그날 업데이트 분을 줘야 한다. 
    # 현재 개발단계에서는 그냥 특정일자 재고리스트를 참고하도록하자
    tmp_date = '2023-02-01'
    total_stock_db = list(TotalStock.objects.filter(Q(std_day=tmp_date)).values())
    json_ts = json.dumps(total_stock_db,ensure_ascii=False)

    return render(request, 'home/select_part.html',{'data_ts':json_ts,'datas_pose':prod_pose,'datas_products':products})