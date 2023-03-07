## xl_wings 절대경로 추가
import sys, os



sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))


from datajob.xlwings_dj.shipment_information import ShipmentInformation
from xlwings_job.oracle_connect import DataWarehouse
## 출고
import datetime as dt
from xlwings_job.xl_utils import bring_data_from_db, clear_form, get_each_index_num, get_idx, get_out_info, get_row_list_to_string, get_xl_rng_for_ship_date, row_nm_check, sht_protect
import xlwings as xw
import pandas as pd
import html
import win32com.client as cli
my_date_handler = lambda year, month, day, **kwargs: "%04i-%02i-%02i" % (year, month, day)

wb_cy = xw.Book("cytiva_worker.xlsm").set_mock_caller()
wb_cy = xw.Book.caller()

def branch_ship_ready():

    cur = DataWarehouse()
    sel_sht = wb_cy.selection.sheet
    mode = sel_sht.range("H4")
    if mode.value != 'waiting_for_out':
        wb_cy.app.alert("'waiting_for_out' 모드에서만 사용 가능합니다.","Quit")
        return

    col_names = list(pd.DataFrame(cur.execute(f"select column_name from user_tab_columns where table_name = upper('SHIPMENT_INFORMATION')").fetchall())[0])

    val_rng = sel_sht.range("P3:P7")
    val_rng.clear_contents()

    idx_list = get_out_table_for_branch()
    idx_list_str = list(map(lambda e : str(e),idx_list))
    
    db_row_list =[]
    try:
        for idx in idx_list:
            qry = f"select * from SHIPMENT_INFORMATION where si_index = '{idx}'"
            db_row_list.append(cur.execute(qry).fetchone())
        df_br = pd.DataFrame(db_row_list,columns=col_names)
    except:
        wb_cy.app.alert("시트 내용이 손상됬거나 값이 있는 테이블이외의 셀을 선택하셨습니다. 정상적인 선택인 경우 BringDATA 버튼 클릭후 다시 시도해주세요.","Quit")
        return
    
    # 대리점 품목인지여부
    is_branch = len(df_br)==len(df_br[df_br['REMARK']=='대리점'])
    if is_branch != True:
        wb_cy.app.alert("대리점이 아닌 품목이 선택되었습니다. SHIP TO 컬럼의 대리점명이 오타일 경우 ChangeCell기능을 이용하여 정확한 이름으로 변경 및 REMARK컬럼은 '대리점'으로 변경 후 다시시도해주세요","Quit")
        return
    # 출고가능 여부
    is_good = len(df_br)==len(df_br[df_br['STATE'].str.contains("GOOD")])
    if is_good != True:
        wb_cy.app.alert("선택하신 품목의 STATE가 GOOD_WR 또는 GOOD_ESLE 상태에서만 진행 가능합니다. HOLDING일경우 검수완료를 진행 후 다시 시도해주세요.","Quit")
        return
    # ship_to 종류
    selected_site_count = len(set(df_br['SHIP_TO']))
    if selected_site_count != 1:
        wb_cy.app.alert("서로 다른 대리점들이 선택 되었거나 SHIP_TO에 오타가 있습니다. 확인해주세요","Quit")
        return
    branch_name = list(set(df_br['SHIP_TO']))[0]
    
    val_rng[0].value = ",".join(idx_list_str) # 출고행번호
    val_rng[2].value = branch_name # 대리점명

def branch_ship_confirm():
    wb_cy.app.screen_updating = False
    ship_answer = wb_cy.app.alert("Form의 출고를 진행 하시겠습니까?",'SHIP CONFIRM',buttons="yes_no_cancel")
    if ship_answer != "yes":
        wb_cy.app.alert("종료합니다.","Quit")
        return
    
    cur = DataWarehouse()
    sel_sht = wb_cy.selection.sheet
    mode = sel_sht.range("H4")
    val_rng = sel_sht.range("P3:P7")
    if mode.value != 'waiting_for_out':
        wb_cy.app.alert("'waiting_for_out' 모드에서만 사용 가능합니다.","Quit")
        return
    col_names = list(pd.DataFrame(cur.execute(f"select column_name from user_tab_columns where table_name = upper('SHIPMENT_INFORMATION')").fetchall())[0])
    input_values = val_rng.options(numbers=int, dates=my_date_handler).value
    del_method = input_values[3]
    
    if input_values.count(None) > 0 :
        wb_cy.app.alert("ONLY FOR BRANCH의 FORM에 빈 값이 있습니다. 입력 후 시도 해주세요!","Quit")
        return
    idx_list = input_values[0].split(',')
    db_row_list=[]
    try:
        for idx in idx_list:
            qry = f"select * from SHIPMENT_INFORMATION where si_index = '{idx}'"
            db_row_list.append(cur.execute(qry).fetchone())
        df_br = pd.DataFrame(db_row_list,columns=col_names)
    except:
        wb_cy.app.alert("시트 내용이 손상됬거나 값이 있는 테이블이외의 셀을 선택하셨습니다. 정상적인 선택인 경우 BringDATA 버튼 클릭후 다시 시도해주세요.","Quit")
        return
    parcel_list = list(set(df_br['PARCELS_NO']))
    
    if input_values[3] =='택배':
        while True:
            courier_num_list = input_delivery_invoice_number_for_branch(parcel_list)
            pacels_dict = dict(zip(parcel_list,courier_num_list))
            alert_str = str(pacels_dict).replace(",","\n").replace('{','').replace('}',"")
            answer = wb_cy.app.alert("당신의 입력 : \n"+alert_str+" \n 수정하시겠습니까?","CONFIRM",buttons="yes_no_cancel")
            if answer == 'no':
                del_method = pacels_dict
                break
            if answer == 'cancel':
                wb_cy.app.alert("종료합니다.","Quit")
                return
            
    # 1. update_shipdate 매서드사용으로 DB변화완료
    ShipmentInformation.update_shipdate(idx_list,input_values[1],del_method=del_method)
    
    # 2. sht_proctect() 매서드로 edit_mode진입
    sht_protect(False)
    bring_data_from_db()
    sht_protect(True)
    
    # 3. mail form작성
    try : 
        wb_pf = xw.Book('print_form.xlsx')
        ws_br_out= wb_pf.sheets['BRANCH_SHIPPING']
    except:
        print_form_dir = os.path.dirname(os.path.abspath(__file__)) +"\\print_form.xlsx"
        wb_pf =xw.Book(print_form_dir)
        ws_br_out= wb_pf.sheets['BRANCH_SHIPPING']
    carton_count = len(df_br[pd.notna(df_br['NM_OF_PACKAGE'])])
    branch_out_col_names = [
        'ARRIVAL_DATE',
        'AWB_NO',
        'TRIP_NO',
        'NM_OF_PACKAGE',
        'PARCELS_NO',
        'ORDER_NM',
        'SHIP_TO'
    ]
    df_br_out = df_br[branch_out_col_names]
    df_br_out['COURIER_INVOICE_NM'] = None

    last_row = ws_br_out.range("A1048576").end('up').row
    if last_row < 9 :
        last_row = 8
    req_rng = ws_br_out.range((9,1),(last_row,"H"))
    req_rng.api.Borders.LineStyle = -4142
    req_rng.clear_contents()

    ws_br_out.range("A7").value = f"총 {carton_count} CARTON 입니다."
    if input_values[3] == "택배":
        ws_br_out.range("D7").value = f"{input_values[1]}에 택배로 출고 완료 되었습니다."
        for k,y in pacels_dict.items():
            df_idx = df_br_out[df_br_out['PARCELS_NO']==k].index
            df_br_out['COURIER_INVOICE_NM'].loc[df_idx] = y
    else : 
        ws_br_out.range("D7").value = f"{input_values[1]}에 {input_values[3]}(으)로 배송 예정입니다."
        df_br_out['COURIER_INVOICE_NM'] = input_values[3]

    df_br_out.set_index("ARRIVAL_DATE",inplace=True)
    ws_br_out.range("A8").value = df_br_out

    last_row = ws_br_out.range("A1048576").end('up').row
    if last_row < 9 :
        last_row = 8
    req_rng = ws_br_out.range((8,1),(last_row,"H"))
    req_rng.autofit()
    req_rng.api.Borders.LineStyle = 1
    
    # 메일 작성 및 저장
    outlook_send=cli.Dispatch("Outlook.Application")
    branch_db = cur.execute('select * from branch').fetchall()
    branch_db_df = pd.DataFrame(branch_db)
    br_row_df =branch_db_df[branch_db_df[0]==input_values[2]]

    to_mail = br_row_df.iloc[0][2]
    cc_mail = br_row_df.iloc[0][2]

    # 메일 저장 작업
    file_name = 'BR_HTML_CONTENT.html'
    br_html = os.getcwd() + '\\'+file_name
    ws_br_out.to_html(file_name)
    html_body = html.unescape(open(br_html).read())
    mail_obj = outlook_send.CreateItem(0)
    mail_obj.To = to_mail
    mail_obj.CC = cc_mail
    mail_obj.Subject = input_values[2] + " " + input_values[1] + " 배송 정보"
    mail_obj.HTMLBody = html_body
    os.remove(br_html)
    mail_obj.Save()
    
    wb_pf.close()

    wb_cy.app.alert(f"{input_values[1]} 출고가 완료 되었습니다. 메일 임시보관함(Draft)를 확인해보세요!",'DONE')
    wb_cy.app.screen_updating = True


def get_out_table_for_branch(sheet_name=wb_cy.selection.sheet,index_row_number=9):
    """
    xlwings.main.Sheet를 인수로 입력, 해당시트의 index행번호 default = 9 (int)
    list return
    """
    out_row_nums = get_row_list_to_string(row_nm_check(wb_cy)['selection_row_nm'])
    last_col = sheet_name.range("XFD9").end('left').column
    idx_row_num = index_row_number
    col_names = sheet_name.range(sheet_name.range(int(idx_row_num),1),sheet_name.range(int(idx_row_num),last_col)).value
    
    for idx ,i in enumerate(col_names):
        
        if i == None:
            continue
        elif '_INDEX' in i :
            col_num = idx
    
    df_so = pd.DataFrame()
    
    # wb_cy.app.alert(out_row_nums)
    try :
        row_list = out_row_nums.replace(' ','').split(',')
    except:
        row_list = [out_row_nums]
    for row in row_list :
        
        # 연속된 행인 경우
        if '~' in str(row) :
            left_row =int(row.split('~')[0])
            right_row = int(row.split('~')[1])
            rng = sheet_name.range(sheet_name.range(left_row,1),sheet_name.range(right_row,last_col))
            df_so = pd.concat([df_so,pd.DataFrame(sheet_name.range(rng).options(numbers=int).value)]) 
        else :
            
            left_row =int(row)
            right_row = int(row)
            rng = sheet_name.range(sheet_name.range(left_row,1),sheet_name.range(right_row,last_col))
            df_so = pd.concat([df_so,pd.DataFrame(sheet_name.range(rng).options(numbers=int).value).T])

    
    return list(df_so[col_num])

def input_delivery_invoice_number_for_branch(parcel_list):
    wb_cy.app.alert("Parcel_NO에 맞는 송장번호를 입력해주세요.","INFO")
    courier_num_list = []
    for i in range(len(parcel_list)):
        courier_num_list.append(wb_cy.app.api.InputBox(f"'{parcel_list[i]}'에 맞는 송장번호를 입력해주세요.","Delivery invoice number INPUT",Type=2))
    return courier_num_list