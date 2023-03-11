# 절대경로 인식
import sys, os
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))


from spark_infra.spark_session import get_spark_session
from spark_infra.util import cal_std_day
from spark_infra.jdbc import *
from pyspark.sql.functions import col,lit,round


class DailyOutTransformer:

    # 기준 날짜로 주기적으로 하지만 test에서는 정해놓고간다.
    TODAY = '2023-03-10'
    # TODAY = cal_std_day(-1)
    DF_SOOUT = find_data(DataWarehouse, 'SO_OUT')
    DEL_METHOD = find_data(DataWarehouse, 'DELIVERY_METHOD')
    DM_TABLES = []

    
    @classmethod
    def transform(cls):
        df_soout = cls.DF_SOOUT.where(col('SHIP_DATE')==cls.TODAY)
        df_join = df_soout.join(cls.DEL_METHOD, on='DM_KEY',how='left')
        # SO
        df_fin = df_join.groupby(['SHIP_DATE','DEL_MED']).count()
        df_fin = df_fin.select(
            col('SHIP_DATE'),
            col('DEL_MED'),
            col('count').cast('int').alias('QTY')
        )
        save_data(DataMart, df_fin, 'DAILY_SO_OUT')
