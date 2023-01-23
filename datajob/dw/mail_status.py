import os
import pandas as pd
import win32com.client as cli
from datetime import datetime
from oracle_connect import insert_data,DataWarehouse



class MailStatus:


    def put_data(self,ml_status,bin_folder,req_type="SVC"):
        outlook = cli.Dispatch("Outlook.Application").GetNamespace("MAPI") # 아웃룩 
        inbox = outlook.GetDefaultFolder(6) # 받은편지함

        part_request = []
        for ms in inbox.Items:
            if req_type in ms.Subject:
                print(ms.Subject)
                part_request.append(ms)

        list_tmp = []
        for ms in part_request:
            tmp_ms=[]
            ### MAIL_STATUS
            # MS_INDEX(PK) 
            # 자동 증분
            tmp_ms.append(None)
            # ML_SUB(FK)
            tmp_ms.append(ms.Subject)
            # ML_STATUS
            tmp_ms.append(ml_status)
            # STD_DAY(분까지)
            tmp_ms.append(str(datetime.now()).split('.')[0])
            # ML_BIN
            tmp_ms.append(bin_folder)

            list_tmp.append(tmp_ms)
        df_ms= pd.DataFrame(list_tmp)


        insert_data(DataWarehouse(),df_ms,'MAIL_STATUS')