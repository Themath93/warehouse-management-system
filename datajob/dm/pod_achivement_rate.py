# 절대경로 인식
import sys, os
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))

import json
import datetime as dt

from spark_infra.spark_session import get_spark_session
from spark_infra.util import cal_std_day
from spark_infra.jdbc import *
from pyspark.sql.functions import col,lit,round,avg

class PODAchivemnetRate:
    DF_SOOUT = find_data(DataWarehouse, 'SO_OUT')
    TODAY = cal_std_day(-1)
    @classmethod
    def transform(cls):
        df_soout = cls.DF_SOOUT.where(col('POD_DATE').isNotNull()).select(col('TIMELINE'))
        rdd2 =df_soout.rdd.map(
            lambda e: 
            (json.loads(e[0])['data'][0]['c'], json.loads(e[0])['data'][1]['c']))
        df_fin =rdd2.toDF(["POD_TIME","SHIP_TIME"])
        rdd3 = df_fin.rdd.map(
            lambda e:
            [(dt.datetime.strptime(e[1], '%Y-%m-%d %H:%M:%S') - dt.datetime.strptime(e[0], '%Y-%m-%d %H:%M:%S')).days*24*60 +
            (dt.datetime.strptime(e[1], '%Y-%m-%d %H:%M:%S') - dt.datetime.strptime(e[0], '%Y-%m-%d %H:%M:%S')).seconds//60]

        )
        df_fin = rdd3.toDF(["TAKE_MINUTES"])
        df_fin = df_fin.select(col('TAKE_MINUTES').cast('int'))
        rdd4 = df_fin.rdd.map(
            lambda e : [e[0]/60]

        )
        df_fin = rdd4.toDF(["TAKE_HOURS"])
        
        df_fin = df_fin.agg({ 
            'TAKE_HOURS':'avg'
        })
        df_fin = df_fin.select(col('avg(TAKE_HOURS)').cast('int').alias("TAKE_HOURS"))
        df_fin = df_fin.withColumn('STD_DAY',lit(cls.TODAY))
        save_data(DataMart, df_fin, 'POD_DELAY')
        
