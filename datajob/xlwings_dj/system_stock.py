import sys, os
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))))

import os
import json
import pandas as pd
import win32com.client as cli
import datetime as dt
from xlwings_job.oracle_connect import insert_data,DataWarehouse
from xlwings_job.xl_utils import get_empty_row, create_db_timeline
import xlwings as xw

my_date_handler = lambda year, month, day, **kwargs: "%04i-%02i-%02i" % (year, month, day)
class SystemStock:
    """
    SYSTEM_STOCK DB CRUD 담당
    """
    WB_CY = xw.Book("cytiva.xlsm")
    DataWarehouse_DB = DataWarehouse()
    DB_COLS = ['ARTICLE_NUMBER', 'SUBINVENTORY', 'LOCATION', 'QUANTITY', 'IN_DATE', 'EXPIRY_DATE', 'CURRENCY', 'LOT_COST', 'LOT_COST_IN_USD', 'STD_DAY']
    @classmethod
    def put_data(self):
        std_day = str(dt.datetime.now()).split(' ')[0]
        answer = self.WB_CY.app.alert(std_day + " 기준으로 재고리스트 시트내용 업데이트후 yes를 눌러주세요.","STOCK UPDATE",buttons='yes_no_cancel')
        if answer != 'yes':
            self.WB_CY.app.alert("종료합니다.","Quit")
            return
        
        ws_stock = self.WB_CY.sheets['재고리스트']
        last_row = ws_stock.range('A1048576').end('up').row
        last_col = ws_stock.range('XFD1').end('left').column
        col_names = ws_stock.range((1,1),(1,last_col)).value
        for idx,name in enumerate(col_names):
            col_names[idx] = name.replace(' ','_')
        tb_rng_val = ws_stock.range((2,1),(last_row,last_col)).options(numbers=int, dates=my_date_handler).value
        df_stock = pd.DataFrame(tb_rng_val,columns=col_names)

        df_stock = df_stock[self.DB_COLS[:-1]]
        df_stock['STD_DAY']= std_day
        df_stock = df_stock.fillna('')
        df_stock.insert(0,'SS_KEY','')
        insert_data(self.DataWarehouse_DB,df_stock,'SYSTEM_STOCK')