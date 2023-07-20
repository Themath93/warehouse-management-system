from kafka import KafkaProducer
import json
import datetime as dt
import time

class LogMiddleware:

    producer = KafkaProducer(bootstrap_servers='43.205.123.229:9092', value_serializer=lambda v: json.dumps(v).encode('utf-8'))

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        
        param_json =  dict(request.GET.lists()) if request.method == 'GET' else dict(request.POST.lists())

        # views.py 호출 이전에 실행 될 코드
        log_dict = {
            'ip' : request.META['REMOTE_ADDR'],
            'user': str(request.user),
            'http_method': request.method,
            'url': request.path,
            'parameter': param_json,
            'timestamp': int(time.time()),
            'session_id':'None',
            'std_day':str(dt.datetime.today()).split(' ')[0]
        }

        if request.session.keys():
            log_dict['session_id'] = request.session.session_key

        # print(log_dict)
        # print(request)
        self.producer.send('requests', log_dict)
        response = self.get_response(request)

        # views.py 호출 이후 실행 될 코드
        return response