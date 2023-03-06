import sys, os
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))))

import os
import pandas as pd
import datetime as dt
from xlwings_job.oracle_connect import insert_data,DataWarehouse
from xlwings_job.xl_utils import get_empty_row, create_db_timeline
import xlwings as xw

class ServiceRequest:
    """
    SERVICE_REQUEST DB CRUD 담당
    """
    # STATUS = ['requested', 'pick/pack', 'dispathed', 'complete']
    WB_CY = xw.Book("cytiva.xlsm")
    WS_SI = WB_CY.sheets['Shipment information']
    DataWarehouse_DB = DataWarehouse()

    @classmethod
    def update_status(self,svc_key,up_time_content,status):
        cur = self.DataWarehouse_DB
        status_qry = f"UPDATE SERVICE_REQUEST SET STATE = '{status}' WHERE SVC_KEY = '{svc_key}'"
        cur.execute(status_qry)
        timeline_qry = f"UPDATE SERVICE_REQUEST SET TIMELINE = '{create_db_timeline(up_time_content)}' WHERE SVC_KEY = '{svc_key}'"
        cur.execute(timeline_qry)
        cur.execute('commit')