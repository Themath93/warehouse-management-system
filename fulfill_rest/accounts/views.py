from django.shortcuts import render
import json
from django.contrib.auth.decorators import login_required
from django.http import JsonResponse
from django.shortcuts import render, redirect
# from accounts.forms import UserForm
from django.contrib.auth import authenticate,login
from django.http import HttpResponse
# Create your views here.

# from rest_api.models import UserHistory

#회원가입 사용 X
# def signup(request):

   
#     if request.method == "GET":
#         form = UserForm()
#         return render(request, 'accounts/signup.html', {'form': form})

#     form = UserForm(request.POST)

#      #POST로 받았는데 유효성검사 결과 false인 경우 
#     if not form.is_valid():
#         print(form)
#         return render(request, 'accounts/signup.html', {'form': form})

#     #POST로 정상적으로 데이터 받은 경우 db에 user정보 저장
#     form.save()
#     print(form)
#     username = form.cleaned_data.get('username')
#     raw_password = form.cleaned_data.get('password1')
#     #신규사용자인증 및 자동로그인 기능
#     user = authenticate(username=username, password=raw_password)  # 사용자 인증
#     login(request, user)  # 로그인
#     return redirect('/')


