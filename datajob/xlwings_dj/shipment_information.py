import os
import pandas as pd
import win32com.client as cli
from datetime import datetime
from xlwings_job.oracle_connect import insert_data,DataWarehouse


class ShipmentInformation:
    """
    Shipment_Information DB CRUD 담당
    """
    

    def put_data(self):

        return None

        # si업데이트
    def update_data(self,get_each_index_num,ship_date):
        """
        so out시 db의 ship_date수정 및 pod완료시 pod_date 수정이 필요하다.
        # ws_si 시트의 해당하는 db 값 수정
        """
        cur = DataWarehouse()
        idx_info=get_each_index_num
        for val in idx_info['idx_list']:
            si_query = f"UPDATE SHIPMENT_INFORMATION SET SHIP_DATE = '{ship_date}' WHERE SI_INDEX = {val}"
            cur.execute(si_query)
        cur.execute("commit")
