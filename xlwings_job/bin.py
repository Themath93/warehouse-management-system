## xl_wings 절대경로 추가
import sys, os

sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))

from xlwings_job.oracle_connect import DataWarehouse
from xlwings_job.xl_utils import bring_data_from_db, clear_form, get_each_index_num, sht_protect
import pandas as pd
import xlwings as xw
import json
import datetime as dt
import time
wb_cy = xw.Book('cytiva.xlsm')
my_date_handler = lambda year, month, day, **kwargs: "%04i-%02i-%02i" % (year, month, day)

def clear_form_bin():
    sel_sht = wb_cy.selection.sheet
    sel_sht.range("C2:C7").clear_contents()

    mode_cel = sel_sht.range("C2")
    mode_cel.value = 'CHANGE_BIN'

def change_mode_bin():
    sel_sht = wb_cy.selection.sheet
    mode_cel = sel_sht.range("C2")
    if mode_cel.value == "CHANGE_BIN":
        mode_cel.value = 'REGIST_BIN'
        wb_cy.app.alert("'REGIST_BIN' 모드에서는 파트번호,BIN 만 입력해도 등록이 가능합니다.")
    else : 
        mode_cel.value = 'CHANGE_BIN'
        wb_cy.app.alert("'CHANGE_BIN' 모드에서는 NEW_BIN에 들어가는 값이 새로운 BIN으로 등록되며 다른 값들은 적용되지 않습니다.")

def protect_sht_bin():
    sel_sht = wb_cy.selection.sheet
    is_protect = sel_sht.api.ProtectContents
    if is_protect == False :
        sel_sht.api.Protect(Password='themath93', DrawingObjects=True, Contents=True, Scenarios=True,
        UserInterfaceOnly=True, AllowFormattingCells=True, AllowFormattingColumns=True,
        AllowFormattingRows=True, AllowInsertingColumns=True, AllowInsertingRows=True,
        AllowInsertingHyperlinks=True, AllowDeletingColumns=True, AllowDeletingRows=True,
        AllowSorting=True, AllowFiltering=True, AllowUsingPivotTables=True)
    else:
        password = wb_cy.app.api.InputBox("시트보호 해제를 위한 비밀번호를 입력해주세요.","INPUT PASSWORD", Type=2)
        if password == 'themath93':
            sel_sht.api.Unprotect(Password=password)
        else : 
            wb_cy.app.alert("비밀번호가 맞지 않습니다.")
