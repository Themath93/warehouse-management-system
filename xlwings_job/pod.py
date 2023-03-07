## xl_wings 절대경로 추가
import sys, os




sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))
from datajob.xlwings_dj.so_out import SOOut
from datajob.xlwings_dj.local_list import LocalList
from datajob.xlwings_dj.shipment_information import ShipmentInformation
from xlwings_job.oracle_connect import DataWarehouse
from xlwings_job.xl_utils import bring_data_from_db, clear_form, get_each_index_num, sht_protect
import pandas as pd
import xlwings as xw
import json
import datetime as dt
import time
wb_cy = xw.Book("cytiva_worker.xlsm").set_mock_caller()
my_date_handler = lambda year, month, day, **kwargs: "%04i-%02i-%02i" % (year, month, day)





def bring_pod_sht_data():
    answer = wb_cy.app.alert("POD가 진행되지 않은 출고리스트를 가져오시겠습니까?","Data Bring Confirm",buttons='yes_no_cancel')
    if answer != 'yes':
        wb_cy.app.alert("종료합니다.","Quit")
        return
    wb_cy.app.screen_updating = False
    soout_col_names = ['SO_INDEX', 'SHT_ROW_IDX', 'PERSON_IN_CHARGE', 'SHIP_DATE', 'DM_KEY', 'POD_KEY', 'IS_LOCAL','POD_DATE', 'TIMELINE']
    sel_sht = wb_cy.selection.sheet
    idx_cel =  sel_sht.range("C2")
    idx_cel.clear_contents()
    idx_cel.api.Validation.Delete()

    last_row = sel_sht.range('A1048576').end('up').row
    if last_row < 10 :
        last_row = 10
    tb_rng = sel_sht.range((10,"A"),(last_row,"J"))
    tb_rng.clear_contents()

    df_soout = pd.DataFrame(DataWarehouse().execute("select * from so_out where pod_date is null").fetchall())
    df_soout.columns = soout_col_names
    df_soout = df_soout.sort_values('SO_INDEX')
    df_soout.set_index('SO_INDEX',inplace=True)
    df_soout['SO_DETAIL'] = None
    df_soout['LC_DETAIL'] = None


    for so_index in range(len(df_soout)):
        tmp_si_idx = df_soout['SHT_ROW_IDX'].iloc[so_index]
        tmp_lc_idx = df_soout['IS_LOCAL'].iloc[so_index]
        si_and_lc_list = [get_each_index_num(tmp_si_idx), get_each_index_num(tmp_lc_idx)]
        df_so = pd.DataFrame()
        df_lc = pd.DataFrame()
        for i,val in enumerate(si_and_lc_list):

            if val == 'local':
                None
            else:
                if type(val) is list:
                    idx_list = val
                elif type(val) is dict:
                    idx_list=val['idx_list']
                else:
                    idx_list = [val]
                for idx in idx_list:
                    if i == 0:
                        db_obj = DataWarehouse().execute(f"SELECT * FROM SHIPMENT_INFORMATION WHERE SI_INDEX = '{idx}'").fetchone()
                        df_db_obj = pd.DataFrame([db_obj])
                        df_so = pd.concat([df_so,df_db_obj])
                    else:
                        db_obj = DataWarehouse().execute(f"SELECT * FROM LOCAL_LIST WHERE LC_INDEX = '{idx}'").fetchone()
                        df_db_obj = pd.DataFrame([db_obj])
                        df_lc = pd.concat([df_lc,df_db_obj])
                    if not df_so.empty:
                        df_soout['SO_DETAIL'].iloc[so_index] = df_so
                    if not df_lc.empty:
                        df_soout['LC_DETAIL'].iloc[so_index] = df_lc

    df_soout['ORDER_NO'] = None
    df_soout['SHIP_TO'] = None
    for so_index in range(len(df_soout)):
        order_list = []
        shipto_list = []
        so_detail = df_soout['SO_DETAIL'].iloc[so_index]
        lc_detail = df_soout['LC_DETAIL'].iloc[so_index]
        if type(so_detail) != type(None):
            order_list = order_list + list(set(df_soout['SO_DETAIL'].iloc[so_index][6]))
            shipto_list = shipto_list + list(set(df_soout['SO_DETAIL'].iloc[so_index][9]))
        if type(lc_detail) != type(None):
            order_list = order_list + list(set(df_soout['LC_DETAIL'].iloc[so_index][5]))
            shipto_list = shipto_list + list(set(df_soout['LC_DETAIL'].iloc[so_index][8]))

        none_count = order_list.count(None)
        if none_count>0 :
            for i in range(none_count):
                none_idx = order_list.index(None)
                order_list.pop(none_idx)

        none_count_lc = shipto_list.count(None)
        if none_count_lc>0 :
            for i in range(none_count_lc):
                none_idx = shipto_list.index(None)
                shipto_list.pop(none_idx)    
        order_num = ", ".join(order_list)
        ship_to = ", ".join(shipto_list)
        df_soout['ORDER_NO'].iloc[so_index] = order_num
        df_soout['SHIP_TO'].iloc[so_index] = ship_to

    db_del_med = pd.DataFrame(DataWarehouse().execute("select * from DELIVERY_METHOD").fetchall())
    db_pod_med = pd.DataFrame(DataWarehouse().execute("select * from POD_METHOD").fetchall())
    for i in range(len(db_del_med)):
        df_soout['DM_KEY'].replace(db_del_med[0][i],db_del_med[1][i],inplace=True)
    for j in range(len(db_pod_med)):
        df_soout['POD_KEY'].replace(db_pod_med[0][j],db_pod_med[1][j],inplace=True)

    df_soout['TIMELINE'] = df_soout['TIMELINE'].map(lambda e: json.loads(e)['data'][-1]['c'])

    is_lc_list = df_soout[df_soout['IS_LOCAL'].str.contains('lc_')]['IS_LOCAL'].index
    for lc_idx in is_lc_list:
        df_soout['IS_LOCAL'].loc[lc_idx] = 'is_local'


    df_soout= df_soout[['POD_KEY','SHIP_DATE','ORDER_NO', 'SHIP_TO','PERSON_IN_CHARGE','IS_LOCAL', 'POD_DATE',  'DM_KEY',
            'TIMELINE']]

    sel_sht.range("A9").value = df_soout
    last_row = sel_sht.range('A1048576').end('up').row
    if last_row < 10 :
        last_row = 10
    so_idx_list = sel_sht.range("A10:A"+str(last_row)).options(numbers=lambda x: str(int(x))).value
    idx_list = ",".join(so_idx_list)
    idx_cel.api.Validation.Add(Type=3, Formula1=idx_list)
    
    tb_rng = sel_sht.range((9,"A"),(last_row,"J"))
    tb_rng.api.Borders.LineStyle = 1 

    wb_cy.app.screen_updating = True

def see_detail_pod_sht():
    
    sel_sht = wb_cy.selection.sheet
    call_index =  sel_sht.range("C2").options(numbers=lambda x: str(int(x))).value
    if call_index == None:
        wb_cy.app.alert("SO_INDEX를 DropDownList에서 하나 선택한 후 버튼을 눌러주세요. 종료합니다.",'Quit')
        return
    wb_pod = xw.Book(os.path.dirname(os.path.abspath(__file__)) +"\\pod_detail.xlsx")
    pod_sht_pf =wb_pod.sheets[0]
    
    soout_col_names = ['SO_INDEX', 'SHT_ROW_IDX', 'PERSON_IN_CHARGE', 'SHIP_DATE', 'DM_KEY', 'POD_KEY', 'IS_LOCAL','POD_DATE', 'TIMELINE']
    
    df_soout = pd.DataFrame(DataWarehouse().execute(f"select * from so_out where so_index = '{call_index}'").fetchall())
    df_soout.columns = soout_col_names
    df_soout = df_soout.sort_values('SO_INDEX')
    df_soout.set_index('SO_INDEX',inplace=True)
    df_soout['SO_DETAIL'] = None
    df_soout['LC_DETAIL'] = None


    for so_index in range(len(df_soout)):
        tmp_si_idx = df_soout['SHT_ROW_IDX'].iloc[so_index]
        tmp_lc_idx = df_soout['IS_LOCAL'].iloc[so_index]
        si_and_lc_list = [get_each_index_num(tmp_si_idx), get_each_index_num(tmp_lc_idx)]
        df_so = pd.DataFrame()
        df_lc = pd.DataFrame()
        for i,val in enumerate(si_and_lc_list):

            if val == 'local':
                None
            else:
                if type(val) is list:
                    idx_list = val
                elif type(val) is dict:
                    idx_list=val['idx_list']
                else:
                    idx_list = [val]
                for idx in idx_list:
                    if i == 0:
                        db_obj = DataWarehouse().execute(f"SELECT * FROM SHIPMENT_INFORMATION WHERE SI_INDEX = '{idx}'").fetchone()
                        df_db_obj = pd.DataFrame([db_obj])
                        df_so = pd.concat([df_so,df_db_obj])
                    else:
                        db_obj = DataWarehouse().execute(f"SELECT * FROM LOCAL_LIST WHERE LC_INDEX = '{idx}'").fetchone()
                        df_db_obj = pd.DataFrame([db_obj])
                        df_lc = pd.concat([df_lc,df_db_obj])
                    if not df_so.empty:
                        df_soout['SO_DETAIL'].iloc[so_index] = df_so
                    if not df_lc.empty:
                        df_soout['LC_DETAIL'].iloc[so_index] = df_lc
    df_details = df_soout[['SO_DETAIL','LC_DETAIL']].loc[int(call_index)]
    
    so_col_name= list(pd.DataFrame(DataWarehouse().execute("select * from cols where table_name = 'SHIPMENT_INFORMATION'").fetchall())[1])
    lc_col_name= list(pd.DataFrame(DataWarehouse().execute("select * from cols where table_name = 'LOCAL_LIST'").fetchall())[1])

    last_row = pod_sht_pf.range("A1048576").end('up').row
    pod_sht_pf.range((1,"A"),(last_row,"Z")).clear_contents()

    try : 
        so_detail = df_details['SO_DETAIL']
        so_detail.columns = so_col_name
        so_detail.set_index('SI_INDEX',inplace=True)
        pod_sht_pf.range('A1').value = so_detail
    except:
        None

    last_row = pod_sht_pf.range("A1048576").end('up').row + 1

    try :    
        lc_detail = df_details['LC_DETAIL']
        lc_detail.columns = lc_col_name
        lc_detail.set_index('LC_INDEX',inplace=True)
        pod_sht_pf.range('A'+str(last_row)).value = lc_detail
    except:
        None

def pod_done():
    sel_sht = wb_cy.selection.sheet
    date_cell = sel_sht.range("C3").options(dates=my_date_handler)
    date_answer = wb_cy.app.alert(date_cell.value + " 의 날짜로 POD진행하시겠습니까?","POD Confirm",buttons='yes_no_cancel')
    if date_answer != 'yes':
        wb_cy.app.alert("DATE를 입력하시고 다시 진행해주세요. 종료합니다..","Quit")
        return

    answer = wb_cy.app.alert("POD DONE을 진행하시겠습니까?","POD Confirm",buttons='yes_no_cancel')
    if answer != 'yes':
        wb_cy.app.alert("종료합니다.","Quit")
        return

    call_index =  sel_sht.range("C2").options(numbers=lambda x: str(int(x))).value
    if call_index == None:
        wb_cy.app.alert("SO_INDEX를 DropDownList에서 하나 선택한 후 버튼을 눌러주세요. 종료합니다.",'Quit')
        return

    wb_cy.app.screen_updating = False
    soout_col_names = ['SO_INDEX', 'SHT_ROW_IDX', 'PERSON_IN_CHARGE', 'SHIP_DATE', 'DM_KEY', 'POD_KEY', 'IS_LOCAL','POD_DATE', 'TIMELINE']

    df_soout = pd.DataFrame(DataWarehouse().execute(f"select * from so_out where so_index = '{call_index}'").fetchall())
    df_soout.columns = soout_col_names
    df_soout = df_soout.sort_values('SO_INDEX')
    df_soout.set_index('SO_INDEX',inplace=True)
    df_soout['SO_DETAIL'] = None
    df_soout['LC_DETAIL'] = None


    for so_index in range(len(df_soout)):
        tmp_si_idx = df_soout['SHT_ROW_IDX'].iloc[so_index]
        tmp_lc_idx = df_soout['IS_LOCAL'].iloc[so_index]
        si_and_lc_list = [get_each_index_num(tmp_si_idx), get_each_index_num(tmp_lc_idx)]
        df_so = pd.DataFrame()
        df_lc = pd.DataFrame()
        for i,val in enumerate(si_and_lc_list):

            if val == 'local':
                None
            else:
                if type(val) is list:
                    idx_list = val
                elif type(val) is dict:
                    idx_list=val['idx_list']
                else:
                    idx_list = [val]
                for idx in idx_list:
                    if i == 0:
                        db_obj = DataWarehouse().execute(f"SELECT * FROM SHIPMENT_INFORMATION WHERE SI_INDEX = '{idx}'").fetchone()
                        df_db_obj = pd.DataFrame([db_obj])
                        df_so = pd.concat([df_so,df_db_obj])
                    else:
                        db_obj = DataWarehouse().execute(f"SELECT * FROM LOCAL_LIST WHERE LC_INDEX = '{idx}'").fetchone()
                        df_db_obj = pd.DataFrame([db_obj])
                        df_lc = pd.concat([df_lc,df_db_obj])
                    if not df_so.empty:
                        df_soout['SO_DETAIL'].iloc[so_index] = df_so
                    if not df_lc.empty:
                        df_soout['LC_DETAIL'].iloc[so_index] = df_lc
    df_details = df_soout[['SO_DETAIL','LC_DETAIL']].loc[int(call_index)]
    try:
        so_idx_list = list(df_details['SO_DETAIL'][0])
        ShipmentInformation.update_pod_date(so_idx_list,date_cell.value)
        
        wb_cy.sheets['SHIPMENT_INFORMATION'].activate()
        sht_protect(False)
        time.sleep(2)
        bring_data_from_db()

        wb_cy.sheets['POD'].activate()
    except:
        None

    try:
        lc_idx_list = list(df_details['LC_DETAIL'][0])
        
        LocalList.update_pod_date(lc_idx_list,date_cell.value)

        wb_cy.sheets['LOCAL_LIST'].activate()
        
        sht_protect(False)
        time.sleep(2)
        bring_data_from_db()

        wb_cy.sheets['POD'].activate()
    except:
        None
    SOOut.update_pod_date(call_index,date_cell.value)
    wb_cy.app.display_alerts = False
    bring_pod_sht_data()
    wb_cy.app.screen_updating = True
    wb_cy.app.display_alerts = True


def bring_condition_data_pod():
    
    sel_sht = wb_cy.selection.sheet
    answer = wb_cy.app.alert("해당 조건으로 데이터를 가져오시겠습니까?","Confirm",buttons='yes_no_cancel')
    if answer != 'yes':
        wb_cy.app.alert("종료합니다.","Quit")
        return
    
    wb_cy.app.screen_updating = False
    
    #셀정보받기
    last_row = sel_sht.range('A1048576').end('up').row
    if last_row < 10 : last_row=10
    last_col = sel_sht.range('XFD9').end('left').column
    sel_sht.range((10,1),(last_row,last_col)).delete() # 기존 데이터 삭제
    col_names = sel_sht.range((9,1),(9,last_col)).value

    # 검색조건 받기
    conditions= []
    for i in range(2,8):
        conditions.append(sel_sht.range((i,"C")).options(numbers=lambda e : str(int(e)), dates=my_date_handler).value)
        j=i-2
        if conditions[j] is None:
            if j == 0:
                conditions[j] = 'ARRIVAL_DATE'
            else:
                conditions[j] = 'ALL'
                
    #조건값 가져오기
    base_query = "SELECT * FROM IR_ORDER "

    if conditions[1] == 'ALL': conditions[1] = '1999-01-01'
    if conditions[2] ==  'ALL': conditions[2] = '2199-12-31'

    query = base_query + f"WHERE {conditions[0]} BETWEEN '{conditions[1]}' AND '{conditions[2]}' "
    query = query if conditions[3] == 'ALL' else query + f"AND STATE = '{conditions[3]}' "
    query = query if conditions[4] ==  'ALL' else query + f"AND ARTICLE_NUMBER = '{conditions[4]}' "
    query = query if conditions[5] ==  'ALL' else query + f"AND SUBINVENTORY = '{conditions[5]}'"

    # 데이터프레임 형식에 맞게 조정
    df_ir = pd.DataFrame(DataWarehouse().execute(query).fetchall(),columns=col_names)
    df_ir['TIMELINE'] = df_ir['TIMELINE'].map(lambda e: json.loads(e)['data'][-1]['c'])
    df_ir = df_ir.sort_values('IR_INDEX')
    df_ir.set_index('IR_INDEX',inplace=True)
    
    # 데이터 엑셀로 전송
    sel_sht.range("A9").value = df_ir
    
    # 서식조정
    last_row = sel_sht.range("B1048576").end('up').row
    last_col = sel_sht.range("XFD9").end("left").column
    xl_content = sel_sht.range((10,1),(last_row,last_col))
    xl_content.api.Borders.LineStyle = 1 
    
    clear_form()

    wb_cy.app.screen_updating = True