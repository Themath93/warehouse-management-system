import sys, os
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))))

import pandas as pd
import datetime as dt
import win32com.client as cli
from datetime import datetime
from xlwings_job.oracle_connect import DataWarehouse,insert_data

import xlwings as xw
from xlwings_job.xl_utils import create_db_timeline, get_empty_row

class LocalList:
    """
    LocalList DB CRUD 담당
    """
    WB_CY = xw.Book.caller()
    WS_LC = WB_CY.sheets['로컬리스트']
    DataWarehouse_DB = DataWarehouse()

    @classmethod
    def put_data(self):
        
        db_pass = 'themath93'

        """
        해당 모듈은 매우 신중히 사용 하여야 한다.
        비밀번호를 입력받아서 맞을 경우에만 사용한다.
        """
        
        input_data = self.WB_CY.app.api.InputBox("해당 매서드 진행을 위해 발급받은 'db_pass'를 입력해주세요.","DATABASE WARNING", Type=2)
        
        if input_data == db_pass:
            last_row = get_empty_row(self.WS_LC,col=1)-1
            last_col = self.WS_LC.range("XFD9").end('left').column
            
            col_names = self.WS_LC.range((9,1),(9,last_col)).options(numbers=int).value
            content = self.WS_LC.range((10,1),(last_row,last_col)).options(numbers=int).value

            df = pd.DataFrame(content,columns=col_names)
            df.columns = col_names
            df = df.astype('string')
            df = df.astype({
                'LC_INDEX':'int',
                'QUANTITY':'int'
            })
            df['STATE'] = 'GOOD_WR'
            df['TIMELINE'] = create_db_timeline()
            df = df.fillna('')
            insert_data(DataWarehouse(),df,'LOCAL_LIST')
        else :
            self.WB_CY.app.alert("'db_pass'가 맞지않아 매서드를 종료합니다.",'안내')

        # lc업데이트
    @classmethod
    def update_shipdate(self,get_each_index_num,ship_date,status='SHIP_CONFIRM'):
        """
        so out시 db의 ship_date수정 및 pod완료시 pod_date 수정이 필요하다.
        # ws_lc 시트의 해당하는 db 값 수정
        """
        cur = DataWarehouse()
        if type(get_each_index_num) is list:
            idx_list = get_each_index_num
        elif type(get_each_index_num) is dict:
            idx_list=get_each_index_num['idx_list']
        else:
            idx_list = [get_each_index_num]

        for val in idx_list:
            # shipdate 
            query = f"UPDATE LOCAL_LIST SET SHIP_DATE = '{ship_date}' WHERE LC_INDEX = {val}"
            cur.execute(query)
            # timeline
            bring_tl_query = f"SELECT TIMELINE FROM LOCAL_LIST WHERE LC_INDEX = {val}"
            try:
                json_tl = cur.execute(bring_tl_query).fetchone()[0]
            except:
                self.WB_CY.app.alert(f" LC_INDEX : {val} 해당 INDEX는 DB에 등록되지 않은 INDEX입니다. Data는 DataInput기능으로만 저장 가능합니다. 종료합니다.","Quit")
                return 
            cur.execute(bring_tl_query).fetchone()[0]
            timeline_query = f"UPDATE LOCAL_LIST SET TIMELINE = '{create_db_timeline(json_tl)}' WHERE LC_INDEX = {val}"
            cur.execute(timeline_query)
            # state
            update_state_query = f"UPDATE LOCAL_LIST SET STATE = '{status}' WHERE LC_INDEX = {val}"
            cur.execute(update_state_query)
            
        cur.execute("commit")

        return None
    @classmethod
    def update_arrival_date(self):
        return None

    @classmethod
    def update_status(self):
        return None

    @classmethod
    def data_input(self):
        cur = self.DataWarehouse_DB
        last_idx_db = pd.DataFrame(cur.execute('select max(lc_index) from local_list').fetchall())[0][0]
        last_row = self.WS_LC.range("B1048576").end('up').row
        last_col = self.WS_LC.range("XFD9").end("left").column
        col_names = self.WS_LC.range((9,1),(9,last_col)).options(numbers=int,dates=dt.date).value
        content = self.WS_LC.range((10,1),(last_row,last_col)).options(numbers=int,dates=dt.date).value
        df = pd.DataFrame([content],columns=col_names)
        df = df.astype('string')
        df = df.astype({
            'LC_INDEX':'int'
        })
        df = df.fillna('None')
        df['LC_INDEX'] = None
        df_len = len(df)
        df['LC_INDEX'] = [*range(last_idx_db+1,last_idx_db+df_len+1)]
        insert_data(self.DataWarehouse_DB,df,'LOCAL_LIST')