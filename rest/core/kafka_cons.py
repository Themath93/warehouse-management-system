from kafka import KafkaConsumer, consumer

# consumer 객체 생성
consumer = KafkaConsumer(
    'requests',
    bootstrap_servers=['127.0.0.1:9092'],
    auto_offset_reset='earliest',
    enable_auto_commit=True,
    consumer_timeout_ms=1000
)

while True:
    for message in consumer:
        # print(message.topic, message.partition, message.offset, message.key, message.value)
        print(message.value)