from django.shortcuts import render
from django.http import HttpResponse


# Create your views here.

def index(request):
    # return HttpResponse('메인페이지')
    return render(request,'index.html')