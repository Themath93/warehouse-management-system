import os
import pandas as pd
import win32com.client as cli
from datetime import datetime
from xlwings_job.oracle_connect import insert_data,DataWarehouse
from xlwings_job.xl_utils import create_db_timeline
import xlwings as xw

class SOOut:
    """
    SO_OUT DB CRUD 담당
    """
    DataWarehouse_DB = DataWarehouse()
    WB_CY = xw.Book("cytiva_worker.xlsm").set_mock_caller()
    WB_CY = xw.Book.caller()
    def put_data(self,out_list):
        df_so = pd.DataFrame(out_list).T
        df_so[0][0]= None
        # df_so.to_csv(r"C:\Users\lms46\Desktop\fulfill\test.csv",encoding='UTF-8')
        
        insert_data(DataWarehouse(),df_so,'SO_OUT')

    
    @classmethod
    def update_pod_date(self,so_index,date):

        cur = self.DataWarehouse_DB
        query = f"UPDATE SO_OUT SET POD_DATE = '{date}' WHERE SO_INDEX = {so_index}"
        bring_tl_query = f"SELECT TIMELINE FROM SO_OUT WHERE SO_INDEX = {so_index}"
        try:
            json_tl = cur.execute(bring_tl_query).fetchone()[0]
        except:
            self.WB_CY.app.alert(f" SO_INDEX : {so_index} 해당 INDEX는 DB에 등록되지 않은 INDEX입니다. Data는 DataInput기능으로만 저장 가능합니다. 종료합니다.","Quit")
            return 
        timeline_query = f"UPDATE SO_OUT SET TIMELINE = '{create_db_timeline(json_tl)}' WHERE SO_INDEX = {so_index}"
        cur.execute(timeline_query)
        cur.execute(query)
        cur.execute("commit")