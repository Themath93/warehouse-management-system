import sys, os
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))

import cx_Oracle
import os
import pandas as pd



LOCATION = os.path.dirname(os.path.abspath(os.path.dirname(__file__))) + "\\instantclient_fulfill"



def DataWarehouse():
    try:
        os.environ["PATH"] = LOCATION + ";" + os.environ["PATH"]
        cx_Oracle.init_oracle_client(lib_dir=LOCATION)
        connection = cx_Oracle.connect(
            user='dw_fulfill', password='fulfillment123QWE!@#', dsn='fulfill_high'

        )
        cursor = connection.cursor()
        return cursor

    except:

        connection = cx_Oracle.connect(
            user='dw_fulfill', password='fulfillment123QWE!@#', dsn='fulfill_high'

        )
        cursor = connection.cursor()
        return cursor


def insert_data(cursor,DataFrame,table_name):
    """
    cursor객체와 DataFrame, table_name을 받아 해당 테이블에 insert 구문실행
    is_set_index 는 키값을 컬럼으로 진행하는 경우 get_insert_values(DataFrame.columns) 에서 values를 1개 적게 반환 한다.
    그래서 set_index가 되어있는 df의 경우는 get_insert_values로 여부를 넘겨줘야한다.
    """

    query = 'INSERT INTO '+ table_name + f' VALUES ({get_insert_values(DataFrame.columns)})'

    df_tmp = DataFrame.values.tolist()
    df_tmp_tuple =list(pd.Series(df_tmp).map(tuple))
    cursor.executemany(query,df_tmp_tuple,batcherrors= True)
    cursor.execute("commit")

# def select_data(cursor,table_name):




def get_insert_values(col_list):

    """
    cx_Oracle용 values 연속된 숫자 그룹 str만들때 사용 str 반환

    """
    
    list_len = len(col_list)

    value_list = []
    for i in range(1,list_len+1):
        value_list.append(":"+str(i))
    str_val = ','.join(value_list)
    
    return str_val