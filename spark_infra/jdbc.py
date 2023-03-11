"""
JDBC Connect Info
"""
from spark_infra.spark_session import get_spark_session
from enum import Enum

class DataWarehouse(Enum):
    URL = 'jdbc:oracle:thin:@fulfill_high?TNS_ADMIN=/home/worker/project/db/fulfill_wallet'
    PROPS ={
        'user':'dw_fulfill'
       ,'password':'fulfillment123QWE!@#'
    }
    
class WebDB(Enum):
    URL = 'jdbc:oracle:thin:@fulfill_high?TNS_ADMIN=/home/worker/project/db/fulfill_wallet'
    PROPS ={
        'user':'web_fulfill'
       ,'password':'fulfillment123QWE!@#'
    }

class DataMart(Enum):
    URL = 'jdbc:oracle:thin:@fulfill_high?TNS_ADMIN=/home/worker/project/db/fulfill_wallet'
    PROPS ={
        'user':'dm_fulfill'
       ,'password':'fulfillment123QWE!@#DM'
    }  


def save_data(config, dataframe, table_name):
    """
    테이블에 dataframe객체를 저장한다. 테이블에 값이 있을 경우 뒤에 이어서 추가한다.
    중복값이 생길 수 있으니 주의!!
    """
    dataframe.write.jdbc(url=config.URL.value
    , table=table_name
    , mode='append'
    , properties=config.PROPS.value)

def overwrite_data(config, dataframe, table_name):
    """
    mode가 overwrite라 테이블에 덮어씌워버린다.
    사용시 주의할 것!!
    """
    dataframe.write.jdbc(url=config.URL.value
    , table=table_name
    , mode='overwrite'
    , properties=config.PROPS.value)

def find_data(config, table_name):
    """
    전달 받은 접속명, 테이블명에 접속하여 값을 반환해주는 함수 ex (DataWareHouse, 'LOC') --> 데이터웨어하우스라는 접속명의 테이블중 LOC라는 테이블 호출
    """
    return get_spark_session().read.jdbc(url=config.URL.value, table=table_name, properties=config.PROPS.value)