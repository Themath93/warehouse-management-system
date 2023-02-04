from hdfs import InsecureClient

def get_client():
    """
    docker run 시 내가 호스트 포트를 9870에 할당하여 datanodes에 저장되어있는 값을 볼수 있는 곳이다.
    사용유저는 root가 아닌 big 이다.
    여러 위험요소로 linux에서 root계정은 사용하지 않고, 보조 슈퍼계정을 생성하여 sudo로 권한을 끌어다쓴다.
    """
    return InsecureClient('http://localhost:9870', user='worker')