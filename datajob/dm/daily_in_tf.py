# 절대경로 인식
import sys, os
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))


from spark_infra.spark_session import get_spark_session
from spark_infra.util import cal_std_day
from spark_infra.jdbc import *
from pyspark.sql.functions import col,lit,round


class DailyInTransformer:

    # 기준 날짜로 주기적으로 하지만 test에서는 정해놓고간다.
    # TODAY = '2022-10-13'
    TODAY = cal_std_day(-1)
    DF_SI=find_data(DataWarehouse, 'SHIPMENT_INFORMATION')
    DF_LC=find_data(DataWarehouse, 'LOCAL_LIST')

    DM_TABLES = ['DAILY_IN_SO','DAILY_IN_IR','DAILY_IN_BR','DAILY_IN_LC']

    
    @classmethod
    def transform(cls):
        df_lc = cls.DF_LC.where(col('ARRIVAL_DATE')==cls.TODAY)
        df_lc_tmp = df_lc.select(col('ARRIVAL_DATE'),col('LC_INDEX').astype('int'))
        df_lc_tmp = df_lc_tmp.groupBy(['ARRIVAL_DATE']).agg({ 
            'LC_INDEX':'count'
        })
        df_lc_tmp = df_lc_tmp.select(
            col('ARRIVAL_DATE').alias('STD_DAY'),
            col('count(LC_INDEX)').cast('int').alias('QTY'),
        )
        save_data(DataMart, df_lc_tmp, cls.DM_TABLES[3])

        df_si = cls.DF_SI
        df_si = df_si.where(col('ARRIVAL_DATE')==cls.TODAY)
        # SO
        df_si_so=df_si.filter(~df_si.ORDER_NM.contains('IR')). \
                        filter(~df_si.REMARK.contains('대리점')). \
                        filter(~df_si.SHIP_TO.contains('특송'))
        # IR
        df_si_ir=df_si.filter(df_si.ORDER_NM.contains('IR'))
        # BR
        df_si_br=df_si.filter(df_si.REMARK == '대리점')
        df_list = [df_si_so,df_si_ir,df_si_br]

        for idx, df_si_tmp in enumerate(df_list):

            if idx != 2 :
                df_si_tmp = df_si_tmp.select(col('ARRIVAL_DATE'),col('NM_OF_PACKAGE').astype('int'),col('ORDER_TOTAL').astype('float'))
                df_si_tmp = df_si_tmp.groupBy(['ARRIVAL_DATE']).agg({ 
                    'NM_OF_PACKAGE':'count',
                    'ORDER_TOTAL':'sum'
                })
                df_si_tmp = df_si_tmp.select(
                    col('ARRIVAL_DATE').alias('STD_DAY'),
                    col('count(NM_OF_PACKAGE)').cast('int').alias('QTY'),
                    round(col('sum(ORDER_TOTAL)'),2).cast('float').alias('AMOUNT')
                )
            # 대리점은 SHIP_TO 
            else:
                df_si_tmp = df_si_tmp.select(col('SHIP_TO'), col('ARRIVAL_DATE'),col('NM_OF_PACKAGE').astype('int'),col('ORDER_TOTAL').astype('float'))
                df_si_tmp = df_si_tmp.groupBy(['SHIP_TO','ARRIVAL_DATE']).agg({ 
                    'NM_OF_PACKAGE':'count',
                    'ORDER_TOTAL':'sum'
                })
                df_si_tmp = df_si_tmp.select(
                    col('SHIP_TO').alias("BRANCH_NAME"),
                    col('ARRIVAL_DATE').alias('STD_DAY'),
                    col('count(NM_OF_PACKAGE)').cast('int').alias('QTY'),
                    round(col('sum(ORDER_TOTAL)'),2).cast('float').alias('AMOUNT')
                )


            save_data(DataMart, df_si_tmp, cls.DM_TABLES[idx])