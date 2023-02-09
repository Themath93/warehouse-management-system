import cx_Oracle
import os
import pandas as pd
import xlwings as xw
import pandas as pd
import datetime as dt
import html
import win32com.client as cli
from xlwings_job.oracle_connect import DataWarehouse

wb_cy = xw.Book('cytiva.xlsm')

def warehousing_inspection(input_date=str):
    """
    입고품목 검수지 출력
    """
    cur = DataWarehouse()
    print_form_dir = "C:\\Users\\lms46\\Desktop\\fulfill\\xlwings_job\\print_form.xlsx"
    wb_pf = xw.Book(print_form_dir)
    ws_wi = wb_pf.sheets['WAREHOUSING_INSPECTION']
    
    # 입고날짜 임시로정함
    tmp_in_date = '2022-10-06'
    input_date = tmp_in_date
    print_form_dir = "C:\\Users\\lms46\\Desktop\\fulfill\\xlwings_job\\print_form.xlsx"
    wb_pf = xw.Book(print_form_dir)
    division_list = ['대리점','SVC', 'SO','특송']
    
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



def branch_receiving(input_date=str):
    """
    대리점 입고예정메일
    """
    cur = DataWarehouse()
    print_form_dir = "C:\\Users\\lms46\\Desktop\\fulfill\\xlwings_job\\print_form.xlsx"
    wb_pf = xw.Book(print_form_dir)
    ws_br_in = wb_pf.sheets['BRANCH_RECEIVING']
    worker = '홍길동'
    outlook_send=cli.Dispatch("Outlook.Application")

    # DB에서 값가져오기
    tmp_in_date = '2022-10-06'
    input_date = tmp_in_date

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