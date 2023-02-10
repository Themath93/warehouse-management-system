## xl_wings 절대경로 추가
import sys, os

sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))




import json
import cx_Oracle
import pandas as pd
import xlwings as xw
import pandas as pd
import datetime as dt
import html
import win32com.client as cli

from xlwings_job.oracle_connect import DataWarehouse
from xlwings_job.xl_utils import bring_data_from_db, clear_form, get_each_index_num, get_idx, get_xl_rng_for_ship_date, row_nm_check, sht_protect

wb_cy = xw.Book('cytiva.xlsm')
my_date_handler = lambda year, month, day, **kwargs: "%04i-%02i-%02i" % (year, month, day)

def warehousing_ready(input_date=str):


    sel_sht = wb_cy.selection.sheet

    default_in_date = sel_sht.range("C4").options(dates=my_date_handler).value
    check_in_date_ans = wb_cy.app.alert("입고일 :"+default_in_date+" 가 맞습니까? ","IN DATE CHECK",buttons='yes_no_cancel')

    if check_in_date_ans == 'no':

            month_ans = wb_cy.app.api.InputBox("아니라면 입고되는 월을 적어주세요 **숫자만 적어주세요**' ","IN DATE CHECK",Type=1)
            if int(month_ans) < 10:
                month_ans = "0"+str(int(month_ans))
            elif int(month_ans) >= 10:
                month_ans = str(int(month_ans))
            else :
                wb_cy.app.alert("취소하셨습니다. 메서드를 종료합니다.","EXIT")
                return None
            
            day_ans = wb_cy.app.api.InputBox("입고되는 일수를 적어주세요 **숫자만 적어주세요**' ","IN DATE CHECK",Type=1)
            if int(day_ans) < 10:
                day_ans = "0"+str(int(day_ans))
            elif int(day_ans) >= 10:
                day_ans = str(int(day_ans))
            else :
                wb_cy.app.alert("취소하셨습니다. 메서드를 종료합니다.","EXIT")
                return None
            
            current_year_str = str(dt.datetime.today().year)    
            input_date =  current_year_str + "-" + month_ans  + "-" + day_ans

    elif check_in_date_ans == 'cancel':
        wb_cy.app.alert("취소하셨습니다. 메서드를 종료합니다.","EXIT")
        return None
    else :
        input_date = default_in_date

    # wating_for_out에서만 사용가능
    cel_status = sel_sht.range("H4")
    if cel_status.value != 'waiting_for_out':
        wb_cy.app.alert("'waiting_for_out' 상태에서만 해당 기능 사용이 가능합니다. 매서드를 종료합니다.","WAREHOUSING WARNING")
        return None

    # ARRIVAL_DATE에 값이 있으면 사용 불가
    arrv_col_add = get_xl_rng_for_ship_date(wb_cy.selection,ship_date_col_num="K")
    arrv_val_rng = sel_sht.range(arrv_col_add).options(numbers=int,dates=my_date_handler,empty=None, ndim=1)
    is_empty = [True for val in arrv_val_rng.value if type(val) is str]

    if len(is_empty) > 0 : # 값이 있다는 이야기
        wb_cy.app.alert("'ARRIVAL_DATE'컬럼에 이미 값이 있습니니다. 중북 입고는 불가합니다.. 매서드를 종료합니다.","WAREHOUSING WARNING")
        return None

    cel_in_row= sel_sht.range("C2")
    cel_in_row.value = ' '.join(row_nm_check()['selection_row_nm']).replace(',','~').replace(' ', ', ')
    idx_key = get_idx(sel_sht)


    # DB 반영 및 sel_sht 내용 반영 arrvial_date, status ==> HOLDING
    status = 'HOLDING'


    # DB 업데이트
    __update_db_content(input_date, sel_sht, idx_key, status)


    # 검수지출력
    warehousing_inspection(input_date)

    # 대리점 입고예정메일 작성 
    branch_receiving(input_date)

    # form정리
    clear_form()

    # edit_mode로 이동
    sht_protect()

    # DB내용 excel에반영
    bring_data_from_db()

    # edit_mode 종료
    sht_protect()

def __update_db_content(input_date, sel_sht, idx_key, status):


    sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))))
    from datajob.xlwings_dj.shipment_information import ShipmentInformation
    from datajob.xlwings_dj.local_list import LocalList

    tb_dict = {'LOCAL_LIST':LocalList, 'SHIPMENT_INFORMATION':ShipmentInformation}
    list_for_DB = get_each_index_num(idx_key)['idx_list']
    db_tb_name = sel_sht.range("D5").value
    db_obj = tb_dict[db_tb_name]
    db_obj.update_arrival_date(list_for_DB,input_date)
    db_obj.update_status(list_for_DB,status)

    # 대리점리스트 가져오기 FROM DB
    branch_list = DataWarehouse().execute('select branch_name from branch').fetchall()
    branch_list = list(map(lambda e: str(e).replace("('","").replace("',)","")
                        ,branch_list))

    ## SHIPMENT_INFORMATION 컬럼명다가져오기
    query = """
        select column_name
        from   user_tab_columns
        where table_name = 'SHIPMENT_INFORMATION'
    """
    df_col = pd.DataFrame(DataWarehouse().execute(query).fetchall()).T
    df_col[19] = "UP_TIME"
    df_col = df_col.drop(0,axis=1)
    si_col_list = list(df_col.loc[0])

    # input_date가 적혀있는 db정보 가져오기 데이터프레임 타입으로
    df_in = pd.DataFrame(DataWarehouse().execute(f"select * from SHIPMENT_INFORMATION where ARRIVAL_DATE = '{input_date}'"),columns=si_col_list)


    # get_svc_or_branch_idx_list
    svc_idx_list = []
    branch_idx_list = []
    for i in range(len(df_in)):
        order_nm = df_in['ORDER_NM'].iloc[i]
        ship_to = df_in['SHIP_TO'].iloc[i]
        if 'IR-SM' in order_nm:
            svc_idx_list.append(df_in['SI_INDEX'].iloc[i])
        else:
            for br_name in branch_list:
                if br_name in ship_to:
                    branch_idx_list.append(df_in['SI_INDEX'].iloc[i])
    # 서비스 품목 ship_date, pod_date 수정 ==> "SVC"
    if len(svc_idx_list) > 0 :
        ShipmentInformation.update_shipdate(svc_idx_list,input_date)
        ShipmentInformation.update_pod_date(svc_idx_list,"SVC")
        
    # 대리점 품목 Remark업데이트
    if len(branch_idx_list) > 0 :
        ShipmentInformation.update_remark(branch_idx_list,'대리점')



def warehousing_inspection(input_date=str,print_form_dir = "C:\\Users\\lms46\\Desktop\\fulfill\\xlwings_job\\print_form.xlsx"):
    """
    입고품목 검수지 출력
    """
    cur = DataWarehouse()
    
    wb_pf = xw.Book(print_form_dir)
    ws_wi = wb_pf.sheets['WAREHOUSING_INSPECTION']
    division_list = ['대리점','SVC', 'SO','특송']
    # 입고날짜 임시로정함
    # tmp_in_date = '2022-10-06'
    # input_date = tmp_in_date
    
    # 대리점리스트 가져오기 FROM DB
    branch_list = cur.execute('select branch_name from branch').fetchall()
    branch_list = list(map(lambda e: str(e).replace("('","").replace("',)","")
                           ,branch_list))

    # print_form에서 사용할 컬럼명
    col_list = ['AWB_NO','ORDER_NM','SHIP_TO','PARCELS_NO','CHECK','COMMENT','DIVISION']
    join_col = ', '.join(col_list[:-3]) # query용 join
    
    # DB불러오기
    df_in_1 = pd.DataFrame(
        cur.execute(f"select {join_col} from SHIPMENT_INFORMATION where ARRIVAL_DATE = '{input_date}'").fetchall(),
        columns=col_list[:-3]
    )
    df_in_2= pd.DataFrame(columns=col_list[4:])
    df_in = pd.concat([df_in_1,df_in_2],axis=1).fillna("")
    
    # fill division col Division나누기
    for i in range(len(df_in)):
        col_divi = df_in['DIVISION']
        if 'IR' in df_in.iloc[i]['ORDER_NM'] :
            col_divi[i] = division_list[1]
        elif df_in.iloc[i]['ORDER_NM'] == '':
            col_divi[i] = division_list[3]
        elif df_in.iloc[i]['SHIP_TO'] in branch_list:
            col_divi[i] = division_list[0]
        else : 
            col_divi[i] = division_list[2]
    df_in.sort_values(['DIVISION', 'SHIP_TO'],ascending=[True,True],inplace=True)        
    df_in.set_index('DIVISION',inplace=True)
    
    # xs_wi 내용삭제 후 입고품목 입력
    last_row = ws_wi.range("A1048576").end('up').row
    if last_row == 3 :
        last_row = 4
    ws_wi.range((4,1),(last_row,"G")).clear()
    ws_wi.range("A3").value = df_in
    ws_wi.range("C2").value = str(len(df_in)) + " carton" # 화물개수
    ws_wi.range("E2").value = input_date # 입고일자
    
    
    # 서식정리
    last_row = ws_wi.range("A1048576").end('up').row
    ws_wi.range((4,1),(last_row,"G")).api.Borders.LineStyle = 1
    ws_wi.range((4,1),(last_row,"G")).api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
    
    # Print 개발모드에서는 PrintPreivew
    ws_wi.range((1,1),(last_row,"G")).api.PrintPreview()
    # ws_wi.range((1,1),(last_row,"G")).api.Print()



def branch_receiving(input_date=str,print_form_dir = "C:\\Users\\lms46\\Desktop\\fulfill\\xlwings_job\\print_form.xlsx"):
    """
    대리점 입고예정메일
    Return값으로 json 타입의 각대리점별 CI리스트를 반환한다.

    """
    cur = DataWarehouse()
    wb_pf = xw.Book(print_form_dir)
    ws_br_in = wb_pf.sheets['BRANCH_RECEIVING']
    worker = '홍길동'
    outlook_send=cli.Dispatch("Outlook.Application")

    # DB에서 값가져오기
    # tmp_in_date = '2022-10-06'
    # input_date = tmp_in_date

    col_list = ['ARRIVAL_DATE', 'AWB_NO', 'TRIP_NO', 'NM_OF_PACKAGE', 'PARCELS_NO', 'ORDER_NM', 'SHIP_TO']
    join_col = ', '.join(col_list)

    branch_list = cur.execute('select branch_name from branch').fetchall()
    branch_list = list(map(lambda e: str(e).replace("('","").replace("',)","")
                           ,branch_list))

    df_in = pd.DataFrame(
        cur.execute(f"select {join_col} from SHIPMENT_INFORMATION where ARRIVAL_DATE = '{input_date}'").fetchall(),
        columns=col_list
    )
    df_db_br = pd.DataFrame(cur.execute('select * from branch').fetchall())
    
    # 입고된 브랜치리스트만 추리기
    df_list = []
    for br_name in branch_list:
        df_list.append(df_in[df_in['SHIP_TO']==br_name])
    df_br = pd.concat(df_list).reset_index(drop=True).replace("None")
    in_br_list = list(df_br['SHIP_TO'].sort_values().drop_duplicates())
    
    ci_dict_list = []
    # 메일저장 및 form 제작 작업
    for br_name in in_br_list:
        
        df_tmp = df_in[df_in['SHIP_TO']==br_name].reset_index(drop=True).replace("None",pd.NA).set_index('ARRIVAL_DATE')
        to_mail = df_db_br[df_db_br[0]==br_name][2].iloc[0]
        cc_mail = df_db_br[df_db_br[0]==br_name][2].iloc[0]
        
        ### xl form 채우기###
        
        nm_pakage = df_tmp['NM_OF_PACKAGE']
        last_row =  ws_br_in.range("A1048576").end('up').row
        
        ci_list = {"branch":br_name,"ci_list":list(df_tmp['TRIP_NO'].drop_duplicates())}
        ci_dict_list.append(ci_list)
        if last_row < 9: 
            last_row = 9
        ws_br_in.range((9,1),(last_row,"G")).clear()
        ws_br_in.range("A8").value = df_tmp
        last_row =  ws_br_in.range("A1048576").end('up').row
        content = ws_br_in.range((9,1),(last_row,"G"))
        content.api.Borders.LineStyle = 1
        content.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        content.font.size = 11
        ws_br_in.range("A2").value = f'DHL {worker} 입니다.'
        ws_br_in.range("A4").value = f'{input_date[5:].replace("-","월 ")+"일"} 통관되어 {input_date[5:].replace("-","월 ")+"일"}입고 예정인 리스트 입니다.'
        ws_br_in.range("A7").value = f'총 {nm_pakage.count()} Carton 입니다.'
        ### xl form 채우기###
        
        
        # 메일 저장 작업
        file_name = 'BR_HTML_CONTENT.html'
        br_html = os.getcwd() + '\\'+file_name
        ws_br_in.to_html(file_name)
        html_body = html.unescape(open(br_html).read())
        mail_obj = outlook_send.CreateItem(0)
        mail_obj.To = to_mail
        mail_obj.CC = cc_mail
        mail_obj.Subject = br_name + " " + input_date + " 입고예정 품목"
        mail_obj.HTMLBody = html_body
        os.remove(br_html)
        mail_obj.Save()

        # 메일보내기 기능은 CI file을 첨부가 가능할 때 진행행
        # mail_obj.Send()
    return json.dumps(ci_dict_list,ensure_ascii=False)