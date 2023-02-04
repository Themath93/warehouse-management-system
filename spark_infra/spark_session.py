import findspark
from pyspark.sql import SparkSession

def get_spark_session():

    findspark.init()
    return SparkSession.builder.getOrCreate() 
    # getOrCreate 는 진행중인 SparkSession을 반환하고, 만약 없다면 생성해서 반환해주는 method
    # spark의 다양한 매서드와 병렬처리기능을 사용하기위해서 SparkSession을 기점으로 해서 호출해야된다.
    # 판다스가 pandas로 호출하는 것처럼
    # vscode에서는 주피터처럼 yarn을 master로 하여 자동으로 spark를 선언해주지 못하기 때문에 
    # findspark.init()으로 강제 sparksession을 반환하게해준다.