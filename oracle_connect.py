import cx_Oracle
import os
import pandas as pd



LOCATION = r"C:\Users\lms46\Desktop\fulfill\instantclient_fulfill"



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

    # LOCATION = r"C:\Users\lms46\Desktop\fulfill\instantclient_fulfill"
    # os.environ["PATH"] = LOCATION + ";" + os.environ["PATH"]
    # cx_Oracle.init_oracle_client(lib_dir=LOCATION)
    # connection = cx_Oracle.connect(
    # user='dw_fulfill', password='fulfillment123QWE!@#', dsn='fulfill_high')


# class DataMart():



# 
def insert_data(cursor,DataFrame,table_name):
    """
    cursor객체와 DataFrame, table_name을 받아 해당 테이블에 insert 구문실행
    
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