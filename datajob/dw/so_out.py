import os
import pandas as pd
import win32com.client as cli
from datetime import datetime
from oracle_connect import insert_data,DataWarehouse


class SOOut:
    """
    SO_OUT DB CRUD 담당
    """
    

    def put_data(self,out_list):
        df_so = pd.DataFrame(out_list).T
        df_so[0][0]= None
       
        # df_so.to_csv(r"C:\Users\lms46\Desktop\fulfill\test.csv",encoding='UTF-8')
        
        insert_data(DataWarehouse(),df_so,'SO_OUT')