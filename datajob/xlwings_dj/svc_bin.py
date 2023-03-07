import sys, os
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))))


import pandas as pd
import win32com.client as cli
import datetime as dt
from xlwings_job.oracle_connect import insert_data,DataWarehouse
from xlwings_job.xl_utils import create_db_timeline, get_empty_row
import xlwings as xw

my_date_handler = lambda year, month, day, **kwargs: "%04i-%02i-%02i" % (year, month, day)
class SvcBin:
    """
    SVC_BIN DB CRUD 담당
    """
    WB_CY = xw.Book("cytiva_worker.xlsm").set_mock_caller()
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
            ws_bin = self.WB_CY.sheets['SVC_BIN'] 

            last_row = get_empty_row(ws_bin,col=1)-1
            last_col = ws_bin.range("XFD9").end('left').column
            
            col_names = ws_bin.range((9,1),(9,last_col)).options(numbers=int).value
            content = ws_bin.range((10,1),(last_row,last_col)).options(numbers=lambda e: str(int(e)), dates=my_date_handler).value

            df = pd.DataFrame(content,columns=col_names)
            df.columns = col_names
            df = df.astype('string')
            df['BIN_INDEX']=None
            df = df.fillna('')
                       
            insert_data(DataWarehouse(),df,'SVC_BIN')
        else :
            self.WB_CY.app.alert("'db_pass'가 맞지않아 매서드를 종료합니다.",'안내')
