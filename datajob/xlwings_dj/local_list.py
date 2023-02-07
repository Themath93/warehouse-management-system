import sys, os
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))))

import pandas as pd
import win32com.client as cli
from datetime import datetime
from xlwings_job.oracle_connect import DataWarehouse,insert_data

import xlwings as xw
from xlwings_job.xl_utils import get_empty_row

class LocalList:
    """
    LocalList DB CRUD 담당
    """
    WB_CY = xw.Book.caller()
    WS_LC = WB_CY.sheets['로컬리스트']

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
            df = df.fillna('None')
            # self.WB_CY.sheets['Temp_DB'].range("T4").value = str(df.dtypes)
            insert_data(DataWarehouse(),df,'LOCAL_LIST')
        else :
            self.WB_CY.app.alert("'db_pass'가 맞지않아 매서드를 종료합니다.",'안내')

        # lc업데이트
    @classmethod
    def update_data(self,get_each_index_num,ship_date):
        """
        so out시 db의 ship_date수정 및 pod완료시 pod_date 수정이 필요하다.
        # ws_lc 시트의 해당하는 db 값 수정
        """
        cur = DataWarehouse()
        idx_info=get_each_index_num
        for val in idx_info['idx_list']:
            query = f"UPDATE LOCAL_LIST SET SHIP_DATE = '{ship_date}' WHERE LC_INDEX = {val}"
            cur.execute(query)
        cur.execute("commit")

        return None