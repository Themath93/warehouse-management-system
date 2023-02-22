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
from django.db.models import Q
from django.core.mail import send_mail

from apps.authentication.models import *
from core.infra import *

import json
import pandas as pd
import datetime as dt
from PyKakao import Message


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
        total_db = TotalStock.objects.filter(std_day='2023-02-01')
        # total_db = TotalStock.objects.filter(std_day=std_day).using('dw')
        for d in data:
            search_db = total_db.filter(article_number=d['part_no']).get(subinventory=d['currnet_stock'])
            search_db_qty = search_db.quantity
            update_qty = search_db_qty-d['qty']
            if update_qty < 0 : # 이 경우는 중복출고가 이뤄지던 도중 이미 품목이 가로채기 당한경우
                return HttpResponse(d['part_no']+'의 출고가능 수량은'+str(search_db_qty) + '개 입니다. 출고요청을 종료합니다.')
            search_db.quantity=update_qty
            if update_qty == 0: # 출고가능 개수가 0이면 db에서 삭제
                search_db.delete()
            else : 
                search_db.save()
            try : # 해당엔지니어에게 이미 파트가 있는 경우
                search_db_fe = total_db.filter(article_number=d['part_no']).get(subinventory=d['transfer_to'])
                search_db_qty_fe = search_db_fe.quantity
                update_qty_fe = search_db_qty_fe + d['qty']
                search_db_fe.quantity=update_qty_fe
                search_db_fe.save()

            except: # 해당 엔지니어가 해당 파트가 아예없는경우
                TotalStock(
                    article_number = d['part_no'],
                    subinventory = d['transfer_to'],
                    quantity = d['qty'],
                    country = 'None',
                    prod_centre = 'None',
                    prod_group = 'None',
                    description = 'Reqeusted Part',
                    prod_status_type = 'None',
                    bin_cur = 'None',
                    std_day = std_day,
                    state = 'Not Apply to System Yet',
                    state_time = str(dt.datetime.now()).split('.')[0]
               ).save()





        # SERVICE_REQUEST 파트요청 접수..
        key_contain = std_day.replace('-','')[2:]
        try:
            key_count = len(list(ServiceReqeust.objects.filter(svc_key__icontains=key_contain).using('dw').values())) + 1 
        except:
            key_count = 1

        req_db = ServiceReqeust(

            svc_key='SVC_'+fe_initial.split('_')[1]+'_'+key_contain+str(key_count),
            fe_name = request.user.first_name + request.user.last_name,
            fe_initial = fe_initial,
            req_day = req_date,
            req_time = req_time,
            address = address,
            del_met = del_method,
            is_return = is_return,
            is_urgent = is_urgent,
            recipient = recipient,
            contact = contact,
            contact_sub = sub_contact,
            del_instruction = del_instruction,
            parts = data,
            std_day = std_day,
            timeline = create_db_timeline(),
            state = state

            )
        req_db.save(using='dw')

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
                    'fe_name':req_user,
                    'fe_initial':fe_initial,
                    'req_day':req_date,
                    'req_time':req_time,
                    'address':address,
                    'contact':contact,
                    'del_met':del_method,
                    'is_urgent':is_urgent,
                    'is_return':is_return,
                    'recipient':'need_update',
                    'del_instruction':'need_update'
                    

                },
                'std_day': std_day
            },
            'parts':data
        }
        req_json = json.dumps(res,ensure_ascii=False)
        __sending_outlook_mail(request, key_count, std_day, fe_initial, req_json)




        return HttpResponse('요청이 완료되었습니다.')

def __sending_outlook_mail(request, daily_count, std_day, fe_initial, req_json):
    req_json = json.loads(req_json)
    # parts_df = pd.DataFrame(req_json['parts'])
    # parts_df.index = parts_df.index +1
    # parts_df = parts_df.drop(columns='BIN')
    # parts_html = parts_df.to_html()
    info_dict = req_json['meta']['req_info']
    info_dict['parts'] = req_json['parts']

    html_message = loader.render_to_string(
        'home/request_detail.html',
        context=info_dict

    )

    # send_mail(
    #         subject='SVC_' +fe_initial.split('_')[1]+'_' +std_day.replace('-','')[2:]+str(daily_count),
    #         html_message = parts_df.to_html(),
    #         from_email='deyoon@outlook.kr',
    #         recipient_list=['deyoon@outlook.kr',request.user.email],
    #         fail_silently=False,
    #     )

    subject='SVC_' +fe_initial.split('_')[1]+'_' +std_day.replace('-','')[2:]+str(daily_count)
    message = "test"
    from_email='deyoon@outlook.kr'
    to_list=['deyoon@outlook.kr',request.user.email]

    send_mail(subject,message,from_email,to_list,fail_silently=True,html_message=html_message)



def kakao_test(request):


    ## 카카오톡 메시지
    api = Message(service_key = "fb7a4a68eab473037f341e7cd7c73973")
    auth_url = api.get_url_for_generating_code()
    print(auth_url)
    url = "https://localhost:5000/?code=8C9sUYv5L56_XkxkKMVCQ-WX8crs6TzrFw3lzSlxsViamOwDjSDl2_wcVWg13ZGxo8F9sAo9dJcAAAGGd_AJCQ"
    access_token = api.get_access_token_by_redirected_url(url)
    api.set_access_token(access_token)

    content = {
                "title": "오늘의 디저트",
                "description": "아메리카노, 빵, 케익",
                "image_url": "https://mud-kage.kakao.com/dn/NTmhS/btqfEUdFAUf/FjKzkZsnoeE4o19klTOVI1/openlink_640x640s.jpg",
                "image_width": 640,
                "image_height": 640,
                "link": {
                    "web_url": "http://www.daum.net",
                    "mobile_web_url": "http://m.daum.net",
                    "android_execution_params": "contentId=100",
                    "ios_execution_params": "contentId=100"
                }
            }

    item_content = {
                "profile_text" :"Kakao",
                "profile_image_url" :"https://mud-kage.kakao.com/dn/Q2iNx/btqgeRgV54P/VLdBs9cvyn8BJXB3o7N8UK/kakaolink40_original.png",
                "title_image_url" : "https://mud-kage.kakao.com/dn/Q2iNx/btqgeRgV54P/VLdBs9cvyn8BJXB3o7N8UK/kakaolink40_original.png",
                "title_image_text" :"Cheese cake",
                "title_image_category" : "Cake",
                "items" : [
                    {
                        "item" :"Cake1",
                        "item_op" : "1000원"
                    },
                    {
                        "item" :"Cake2",
                        "item_op" : "2000원"
                    },
                    {
                        "item" :"Cake3",
                        "item_op" : "3000원"
                    },
                    {
                        "item" :"Cake4",
                        "item_op" : "4000원"
                    },
                    {
                        "item" :"Cake5",
                        "item_op" : "5000원"
                    }
                ],
                "sum" :"Total",
                "sum_op" : "15000원"
            }

    social = {
                "like_count": 100,
                "comment_count": 200,
                "shared_count": 300,
                "view_count": 400,
                "subscriber_count": 500
            }

    buttons = [
                {
                    "title": "웹으로 이동",
                    "link": {
                        "web_url": "http://www.daum.net",
                        "mobile_web_url": "http://m.daum.net"
                    }
                },
                {
                    "title": "앱으로 이동",
                    "link": {
                        "android_execution_params": "contentId=100",
                        "ios_execution_params": "contentId=100"
                    }
                }
            ]

    api.send_feed(content=content, item_content=item_content, social=social, buttons=buttons)
    return HttpResponse("sent message")