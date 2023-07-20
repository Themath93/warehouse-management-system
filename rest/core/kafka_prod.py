import time
from kafka import KafkaProducer
import json
# producer 객체 생성
# acks 0 -> 빠른 전송우선, acks 1 -> 데이터 정확성 우선
producer = KafkaProducer(acks=0, 
                            compression_type='gzip',
                            bootstrap_servers=['43.205.123.229:9092'],
                            value_serializer=lambda v: json.dumps(v).encode('utf-8'))

start = time.time()

producer.flush()
for i in range(2):
    producer.send(topic="test",value=str(i))
    producer.flush() #queue에 있는 데이터를 보냄

end = time.time() - start
print(end)