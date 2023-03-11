# INSERT TO WEB_FULFILL
# 절대경로 인식
import sys, os
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))


from spark_infra.spark_session import get_spark_session
from spark_infra.util import cal_std_day
from spark_infra.jdbc import *
from pyspark.sql.functions import col,lit

class TotalStockTransformer:

    # 기준 날짜로 주기적으로 하지만 test에서는 정해놓고간다.
    # TODAY = '2023-02-01'
    TODAY = cal_std_day(-1)
    sys_stock=find_data(DataWarehouse, 'SYSTEM_STOCK')
    products = find_data(DataWarehouse, 'PRODUCTS')
    
    @classmethod
    def transform(cls):

        sys_stock_df=cls.sys_stock.where(col('STD_DAY')==cls.TODAY)
        sys_group = sys_stock_df.groupBy(['ARTICLE_NUMBER','SUBINVENTORY']).count().collect()
        df_sys = get_spark_session().createDataFrame(sys_group)
        join_df = df_sys.join(cls.products, on='ARTICLE_NUMBER', how='left').collect()
        df_fin = get_spark_session().createDataFrame(join_df)
        df_fin = df_fin.select("*").withColumn('STD_DAY',lit(cls.TODAY))
        df_fin = df_fin.withColumnRenamed('count','QUANTITY')
        df_fin = df_fin.select("*").withColumn('STATE',lit('None'))
        df_fin = df_fin.select("*").withColumn('STATE_TIME',lit('None'))
        save_data(WebDB, df_fin, 'TOTAL_STOCK')
