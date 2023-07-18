from kafka import KafkaConsumer, consumer
import json
import datetime as dt
from time import sleep

# consumer 객체 생성
consumer = KafkaConsumer(
    'requests',
    bootstrap_servers=['43.201.103.136:9092'],
    auto_offset_reset='earliest',
    enable_auto_commit=True,
    consumer_timeout_ms=1000,
    value_deserializer=lambda x: json.loads(x.decode('utf-8'))
)

today = str(dt.datetime.today()).split(' ')[0]

while True:
    for message in consumer:
        if len(message) == 0: sleep(1)
        if  today in str(message.value) :
                jsoned_obj =json.dumps(message.value,ensure_ascii=False)
                print(jsoned_obj)

        # print(message.value)
        # print(type(message))