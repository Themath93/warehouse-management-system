import sys, os
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))))

import os
import pandas as pd
import win32com.client as cli
import datetime as dt
from xlwings_job.oracle_connect import insert_data,DataWarehouse
from xlwings_job.xl_utils import get_empty_row
import xlwings as xw

class ShipmentInformation:
    """
    Shipment_Information DB CRUD 담당
    """
    WB_CY = xw.Book.caller()
    WS_SI = WB_CY.sheets['Shipment information']
    DataWarehouse_DB = DataWarehouse()

    @classmethod
    def put_data(self,manager_mode = False):
        db_pass = 'themath93'

        """
        해당 모듈은 매우 신중히 사용 하여야 한다.
        비밀번호를 입력받아서 맞을 경우에만 사용한다.
        """

        input_data = self.WB_CY.app.api.InputBox("해당 매서드 진행을 위해 발급받은 'db_pass'를 입력해주세요.","DATABASE WARNING", Type=2)

        if input_data == db_pass:
            self.__create_db_form_and_insert()
        else :
            self.WB_CY.app.alert("'db_pass'가 맞지않아 매서드를 종료합니다.",'안내')
        return None

    @classmethod
    def __create_db_form_and_insert(self):
        self.WS_SI.api.AutoFilterMode = False   # 필터모드해제
        last_row = self.WS_SI.range("A1048576").end('up').row
        last_col = self.WS_SI.range("XFD9").end('left').column
        col_names_1 = self.WS_SI.range((9,1),(9,7)).options(numbers=int,dates=dt.date).value
        content_1 = self.WS_SI.range((10,1),(last_row,7)).options(numbers=int,dates=dt.date).value
        col_names_2 = self.WS_SI.range((9,8),(9,9)).options(numbers=int,dates=dt.date).value
        content_2 = self.WS_SI.range((10,8),(last_row,9)).options(numbers=int,dates=dt.date).value
        col_names_3 = self.WS_SI.range((9,10),(9,last_col)).options(numbers=int,dates=dt.date).value
        content_3 = self.WS_SI.range((10,10),(last_row,last_col)).options(numbers=int,dates=dt.date).value
        df_1 = pd.DataFrame([content_1],columns=col_names_1)
        df_2 = pd.DataFrame([content_2],columns=col_names_2)
        df_3 = pd.DataFrame([content_3],columns=col_names_3)
        df = pd.concat([df_1,df_2,df_3],axis=1)
        df = df.astype('string')
        df = df.astype({
                'SI_INDEX':'int'
            })
        df = df.fillna('None')
        insert_data(self.DataWarehouse_DB,df,'SHIPMENT_INFORMATION')

        # si업데이트
    def update_data(self,get_each_index_num,ship_date):
        """
        so out시 db의 ship_date수정 및 pod완료시 pod_date 수정이 필요하다.
        # ws_si 시트의 해당하는 db 값 수정
        """
        cur = self.DataWarehouse_DB
        idx_info=get_each_index_num
        for val in idx_info['idx_list']:
            si_query = f"UPDATE SHIPMENT_INFORMATION SET SHIP_DATE = '{ship_date}' WHERE SI_INDEX = {val}"
            cur.execute(si_query)
        cur.execute("commit")
    @classmethod
    def data_input(self):
        cur = self.DataWarehouse_DB
        last_idx_db = pd.DataFrame(cur.execute('select max(si_index) from shipment_information').fetchall())[0][0]
        last_row = self.WS_SI.range("B1048576").end('up').row
        last_col = self.WS_SI.range("XFD9").end("left").column
        col_names_1 = self.WS_SI.range((9,1),(9,7)).options(numbers=int,dates=dt.date).value
        content_1 = self.WS_SI.range((10,1),(last_row,7)).options(numbers=int,dates=dt.date).value
        col_names_2 = self.WS_SI.range((9,8),(9,9)).options(numbers=int,dates=dt.date).value
        content_2 = self.WS_SI.range((10,8),(last_row,9)).options(numbers=int,dates=dt.date).value
        col_names_3 = self.WS_SI.range((9,10),(9,last_col)).options(numbers=int,dates=dt.date).value
        content_3 = self.WS_SI.range((10,10),(last_row,last_col)).options(numbers=int,dates=dt.date).value
        df_1 = pd.DataFrame([content_1],columns=col_names_1)
        df_2 = pd.DataFrame([content_2],columns=col_names_2)
        df_3 = pd.DataFrame([content_3],columns=col_names_3)
        df = pd.concat([df_1,df_2,df_3],axis=1)
        df = df.astype('string')
        df = df.astype({
            'SI_INDEX':'int'
        })
        df = df.fillna('None')
        df['SI_INDEX'] = None
        df_len = len(df)
        df['SI_INDEX'] = [*range(last_idx_db+1,last_idx_db+df_len+1)]
        insert_data(self.DataWarehouse_DB,df,'SHIPMENT_INFORMATION')