import os
import pandas as pd
import win32com.client as cli
from oracle_connect import insert_data,DataWarehouse


class MailDetail:


    def put_data(self,req_type="SVC"):
        outlook = cli.Dispatch("Outlook.Application").GetNamespace("MAPI") # 아웃룩 
        inbox = outlook.GetDefaultFolder(6) # 받은편지함

        part_request = []
        for ms in inbox.Items:
            if req_type in ms.Subject:
                print(ms.Subject)
                part_request.append(ms)

            list_tmp = []
        for ms in part_request:

            tmp_md = []
            ### MAIL_DETAIL
            # ML_SUB(PK)
            tmp_md.append(ms.Subject)
            # ML_TYPE_NM
            tmp_md.append(0)
            # STD_DAY
            tmp_md.append(str(ms.CreationTime).split('.')[0])
            # ML_BODY
            tmp_md.append(ms.Body.split("\r\n")[0])

            list_tmp.append(tmp_md)
        df_md= pd.DataFrame(list_tmp)

        insert_data(DataWarehouse(),df_md,'MAIL_DETAIL')