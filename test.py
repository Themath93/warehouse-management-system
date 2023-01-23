import numpy as np
import pandas as pd
import xlwings as xw
from pathlib import Path
import os
import sys
import cx_Oracle

def test():
    wb = xw.Book.caller()
    # LOCATION = os.getcwd()+"\instantclient_fulfill"

    wb.sheets['통합제어'].range('U9').value = get_df_from_db()

def get_df_from_db(table_name='all_tables'):
    """
    table_name 에는 해당하는 table_name만 넣으면 자동으로 pandas 데이터 프레임으로 변환해준다.
    """
    LOCATION = r"C:\Users\lms46\Desktop\fulfill\instantclient_fulfill"
    os.environ["PATH"] = LOCATION + ";" + os.environ["PATH"]
    cx_Oracle.init_oracle_client(lib_dir=LOCATION)
    connection = cx_Oracle.connect(
    user='dw_fulfill', password='fulfillment123QWE!@#', dsn='fulfill_high'

)
    cursor = connection.cursor()
    rows_or = []
    try:
        for row_data in cursor.execute("select * from " + table_name):
            rows_or.append(list(row_data))
        colum_num = len(list(row_data))
    except: 
        print("Error -> 테이블명이 정확하지 않습니다.")
        raise
    num = int(colum_num)
    col_name = []
    for i in range(num):
        col_name.append('column_'+str(i))

    df = pd.DataFrame(columns=col_name)
    for idx,row in enumerate(rows_or):
        df.loc[idx] = row
    if table_name == 'all_tables' :
        df = df[df['column_0'] == 'DW_FULFILL'].rename(columns={'column_1':'조회가능한_테이블명'})[['조회가능한_테이블명']]. \
            reset_index().drop(columns='index')
        return df
    else :    
        return df
