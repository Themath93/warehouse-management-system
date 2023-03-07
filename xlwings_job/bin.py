## xl_wings 절대경로 추가
import sys, os

sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))

from xlwings_job.oracle_connect import DataWarehouse, insert_data
from xlwings_job.xl_utils import bring_data_from_db, clear_form, get_each_index_num, sht_protect
import pandas as pd
import xlwings as xw
import json
import datetime as dt
import string
wb_cy = xw.Book.caller()
my_date_handler = lambda year, month, day, **kwargs: "%04i-%02i-%02i" % (year, month, day)

def clear_form_bin():
    sel_sht = wb_cy.selection.sheet
    sel_sht.range("C2:C7").clear_contents()

    mode_cel = sel_sht.range("C2")
    mode_cel.value = 'CHANGE_BIN'
    sel_sht.range('C3:C6').api.Locked = True
    sel_sht.range('C3:C6').color = (221, 217, 196)

def change_mode_bin():
    sel_sht = wb_cy.selection.sheet
    mode_cel = sel_sht.range("C2")
    table_name_cel = sel_sht.range("D5")
    status_cel = sel_sht.range("H4")
    table_name_cel.value = 'SVC_BIN'
    status_cel.value = 'edit_mode'
    if mode_cel.value == "CHANGE_BIN":
        mode_cel.value = 'REGIST_BIN'
        wb_cy.app.alert("'REGIST_BIN' 모드에서는 ARTICLE_NUMBER, NEW_BIN 만 입력해도 등록이 가능합니다.")
        sel_sht.range('C3:C6').api.Locked = False
        sel_sht.range('C3:C6').color = None
        sel_sht.range("C3:C7").clear_contents()
    else : 
        mode_cel.value = 'CHANGE_BIN'
        wb_cy.app.alert("'CHANGE_BIN' 모드에서는 NEW_BIN에 들어가는 값이 새로운 BIN으로 등록되며 다른 값들은 적용되지 않습니다.")
        clear_form_bin()

def protect_sht_bin(password=None):
    sel_sht = wb_cy.selection.sheet
    is_protect = sel_sht.api.ProtectContents
    if is_protect == False :
        sel_sht.api.Protect(Password='themath93', DrawingObjects=True, Contents=True, Scenarios=True,
        UserInterfaceOnly=True, AllowFormattingCells=True, AllowFormattingColumns=True,
        AllowFormattingRows=True, AllowInsertingColumns=True, AllowInsertingRows=True,
        AllowInsertingHyperlinks=True, AllowDeletingColumns=True, AllowDeletingRows=True,
        AllowSorting=True, AllowFiltering=True, AllowUsingPivotTables=True)
    else:
        if password == None:
            password = wb_cy.app.api.InputBox("시트보호 해제를 위한 비밀번호를 입력해주세요.","INPUT PASSWORD", Type=2)
        if password == 'themath93':
            sel_sht.api.Unprotect(Password=password)
        else : 
            wb_cy.app.alert("비밀번호가 맞지 않습니다.")

def select_row_bin():
    sel_sht = wb_cy.selection.sheet
    sel_cells = sel_sht.range(wb_cy.selection.address)
    col_num = sel_cells.column
    bin_mode = sel_sht.range("C2").value
    col_alphabet =string.ascii_uppercase[col_num-1]
    
    if bin_mode != "CHANGE_BIN":
        wb_cy.app.alert("'CHANGE_BIN' 모드에서만 사용가능합니다.","Change Cell WARNING")
        return
    if type(sel_cells.value) is list :
        wb_cy.app.alert("하나의 셀만 선택 후 진행해주세요. 두 개 이상은 불가합니다.","Change Cell WARNING")
        return
    if sel_cells.row < 10:
        wb_cy.app.alert("파트 정보가있는 셀만 클릭해주세요.","Clicked Empty Cell")
        return
    off_set_rng = sel_cells.address.replace(col_alphabet,'A')
    bin_idx = sel_sht.range(off_set_rng).options(numbers=lambda e : str(int(e))).value
    df_bin_idx = pd.DataFrame(DataWarehouse().execute(f"select BIN_INDEX, ARTICLE_NUMBER, BIN, SUBINVENTORY from svc_bin where bin_index = '{bin_idx}'").fetchone())
    sel_sht.range("C3").options(transpose=True).value = list(df_bin_idx[0])

def change_bin():
    cur=DataWarehouse()
    sel_sht = wb_cy.selection.sheet
    bin_forms = sel_sht.range("C2:C7").options(numbers=lambda e : str(int(e))).value

    answer = wb_cy.app.alert("입력하신 NEW_BIN 값으로 BIN을 변경하시겠습니까","CONFIRM",buttons="yes_no_cancel")
    if answer != "yes":
        wb_cy.app.alert("종료합니다.","Quit")
        return
    if bin_forms[1] is None:
        wb_cy.app.alert("선택된 ROW가 없습니다.SelectRow버튼으로 선택 후 다시 시도해주세요.","Empty BIN_INDEX")
        return
    if bin_forms[-1] is None:
        wb_cy.app.alert("'NEW_BIN'에 값이 없습니다. 값을 입력 후 다시 시도해주세요.","Empty Cell")
        return
    

    cur.execute(f"UPDATE SVC_BIN SET BIN = '{bin_forms[-1]}' WHERE BIN_INDEX = '{bin_forms[1]}'")
    cur.execute("commit")
    clear_form_bin()
    bring_data_from_db()


def regist_bin():
    sel_sht = wb_cy.selection.sheet
    bin_mode = sel_sht.range("C2").value
    bin_forms = sel_sht.range("C2:C7").options(numbers=lambda e : str(int(e))).value
    last_col = sel_sht.range('XFD9').end('left').column
    col_names = sel_sht.range((9,1),(9,last_col)).value
    if bin_mode != "REGIST_BIN":
        wb_cy.app.alert("'REGIST_BIN' 모드에서만 사용가능합니다.","Regist Cell WARNING")
        return
    none_count = [bin_forms[2],bin_forms[-1]].count(None)
    if none_count > 0 :
        wb_cy.app.alert("'ARTICLE_NUMBER', 'NEW_BIN'에는 필수로 값이 있어야합니다.","Regist Cell WARNING")
        return

    is_db_data = not pd.DataFrame(DataWarehouse().execute(f"SELECT * FROM SVC_BIN WHERE ARTICLE_NUMBER = '{bin_forms[2]}'").fetchone()).empty
    if is_db_data is True:
        wb_cy.app.alert(f"입력하신 'ARTICLE_NUMBER' : {bin_forms[2]} 는 등록되어있습니다. 종료합니다.","Regist Cell WARNING")
        return

    df_regist = pd.DataFrame([bin_forms]).drop(columns=[0,3,5])
    df_regist[1]=None
    df_regist.insert(2,'BIN_OLD',None)
    df_regist.insert(2,'BIN',bin_forms[-1])
    df_regist.columns = col_names

    answer = wb_cy.app.alert("새로운 BIN을 등록하시겠습니까?","CONFIRM",buttons="yes_no_cancel")
    if answer != "yes":
        wb_cy.app.alert("종료합니다.","Quit")
        return

    insert_data(DataWarehouse(),df_regist,'SVC_BIN')
    sel_sht.range("C3:C7").clear_contents()
    bring_data_from_db()