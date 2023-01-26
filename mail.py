from datetime import datetime
from oracle_connect import DataWarehouse
from datajob.dw.mail_detail import MailDetail
from datajob.dw.mail_status import MailStatus
from xl_utils import get_current_time, get_empty_row

import xlwings as xw
import pandas as pd
import win32com.client as cli
import json




wb_cy = xw.Book.caller()


class Email():
    SHEET_NAMES =  ['Temp_DB', 'Shipment information', '인수증', 
    '대리점송장', '대리점 출고대기', '로컬리스트', 'In-Transit part report', '기타리스트',
     '출고리스트', 'Cytiva Inventory BIN','통합제어']

    STATUS = ['waiting_for_out', 'ship_is_ready', '_is_empty','local_out_row_input_required']

    ML_STATUS= ['REQUESTED','PROCESSING', 'SHIPPED']

    ML_FOLDERS = ['inbox','1_Requests', '2_Processing', '3_ShipConfirmed']

    SORT_REQUEST = ['SVC']

    WS_DB = wb_cy.sheets[SHEET_NAMES[0]]
    WS_SI = wb_cy.sheets[SHEET_NAMES[1]]
    WS_POD = wb_cy.sheets[SHEET_NAMES[2]]
    WS_BR = wb_cy.sheets[SHEET_NAMES[3]]
    WS_LC = wb_cy.sheets[SHEET_NAMES[5]]
    WS_SVMX = wb_cy.sheets[SHEET_NAMES[6]]
    WS_OTHER = wb_cy.sheets[SHEET_NAMES[7]]
    WS_MAIN = wb_cy.sheets[SHEET_NAMES[-1]]

    @classmethod
    def connect_email_with_shape(self):
        
        outlook = cli.Dispatch("Outlook.Application").GetNamespace("MAPI") # 아웃룩
        inbox = outlook.GetDefaultFolder(6) # 받은편지함
        msg = inbox.Items #메세지 정보
        
        #시간 계산
        self.WS_MAIN.range('K7').value = datetime.today().strftime("%y년%m월%d일 %H시%M분")

        part_request = []
        service_request = self.SORT_REQUEST[0]
        for ms in msg:
            if service_request in ms.Subject:

                part_request.append(ms)

        # mail_detail 테이블에 인서트
        MailDetail.put_data(self)
        # mail_stauts 테이블에 인서트
        MailStatus.put_data(self,self.ML_STATUS[0],self.ML_FOLDERS[1])


        # 박스생성하여 메일의 서브젝트 이름 설정하기기
        for idx, ms in enumerate(part_request) :
            cel_left = self.WS_MAIN.range('J'+str(get_empty_row(self.WS_MAIN,'J'))).left
            cel_top = self.WS_MAIN.range('J'+str(get_empty_row(self.WS_MAIN,'J'))).top
            cel_width = self.WS_MAIN.range('J'+str(get_empty_row(self.WS_MAIN,'J'))).width
            cel_height = self.WS_MAIN.range('J'+str(get_empty_row(self.WS_MAIN,'J'))).height
            self.WS_MAIN.api.Shapes.AddShape(125, cel_left, cel_top, cel_width, cel_height)
            self.WS_MAIN.shapes[-1].name = ms.Subject
            self.WS_MAIN.shapes[-1].text = '메일열기'
            self.WS_MAIN.shapes[-1].api.TextFrame.HorizontalAlignment = 2
            self.WS_MAIN.shapes[-1].api.TextFrame.VerticalAlignment = 2

            # print_form 연결
            cel_left_pf = self.WS_MAIN.range('R'+str(get_empty_row(self.WS_MAIN,'R'))).left
            cel_top_pf = self.WS_MAIN.range('R'+str(get_empty_row(self.WS_MAIN,'R'))).top
            cel_width_pf = self.WS_MAIN.range('R'+str(get_empty_row(self.WS_MAIN,'R'))).width
            cel_height_pf = self.WS_MAIN.range('R'+str(get_empty_row(self.WS_MAIN,'R'))).height
            self.WS_MAIN.api.Shapes.AddShape(125, cel_left_pf, cel_top_pf, cel_width_pf, cel_height_pf)
            self.WS_MAIN.shapes[-1].name = ms.Subject+'_prfm'
            self.WS_MAIN.shapes[-1].text = '프린트하기'
            self.WS_MAIN.shapes[-1].api.TextFrame.HorizontalAlignment = 2
            self.WS_MAIN.shapes[-1].api.TextFrame.VerticalAlignment = 2
            
            ms_body = (ms.Body)
            split_list_sub =  ms.Subject.split('_')
            req_type = split_list_sub[0]
            body = ms_body[:ms_body.rfind('}')+1]
            json_body = json.loads(body)
            dict_data = json_body['data']

            # 메일 내용 채우기
            # index는 메일제목
            self.WS_MAIN.range('J'+str(get_empty_row(self.WS_MAIN,'J'))).value = ms.Subject
            # print_form용 번호도 추가
            self.WS_MAIN.range('R'+str(get_empty_row(self.WS_MAIN,'R'))).value = idx +1
            # 요청타입 K
            self.WS_MAIN.range('K'+str(get_empty_row(self.WS_MAIN,'K'))).value = req_type
            # 출고요청일 L
            self.WS_MAIN.range('L'+str(get_empty_row(self.WS_MAIN,'L'))).value = json_body['meta']['std_day']
            # 배송요청일 M
            self.WS_MAIN.range('M'+str(get_empty_row(self.WS_MAIN,'M'))).value = dict_data['req_day']+' '+dict_data['req_time']
            # 담당자 N
            self.WS_MAIN.range('N'+str(get_empty_row(self.WS_MAIN,'N'))).value = dict_data['fe_initial']
            # 긴급여부 O
            self.WS_MAIN.range('O'+str(get_empty_row(self.WS_MAIN,'O'))).value = dict_data['is_urgent']

        

            # 출고까지남은시간 P
            left_time = self.WS_MAIN.range('M'+str(get_empty_row(self.WS_MAIN,'M')-1)).value - datetime.today()
            # 시간
            left_hour = round(left_time.total_seconds()/60 //60)
            # 분
            left_min = round(left_time.total_seconds()/60 %60)
            self.WS_MAIN.range('P'+str(get_empty_row(self.WS_MAIN,'P'))).value = str(left_hour) + '시간 ' + str(left_min)+'분 남음'
           
            # 해당 배송건 현재상태 Q
            self.WS_MAIN.range('Q'+str(get_empty_row(self.WS_MAIN,'Q'))).value = self.ML_STATUS[0]
       
        # 시간계산


        # 도형리스트
        shape_list = list(self.WS_MAIN.shapes)
        df_shape_list = pd.DataFrame(shape_list)
        # 서비스만 추출하기

        shape_fe = []

        for shp in shape_list:
            if (service_request in shp.name) and ('_prfm' not in shp.name) :
                shape_fe.append(shp)
                shp.api.OnAction = 'connect_email'
            elif '_prfm' in shp.name :
                shp.api.OnAction = 'print_form'

        
        # move_mail로 실제메일 이동
        move_mail('inbox','1_Requests')


    @classmethod
    def get_email(self):
        shape_list = list(self.self.WS_MAIN.shapes)
        df_shape_list = pd.DataFrame(shape_list)
        # 서비스만 추출하기
        shape_fe = []

        for shp in shape_list:
            if self.SORT_REQUEST[0] in shp.name :
                shape_fe.append(shp)
                shp.api.OnAction = 'ClickedShapeName'

    @classmethod
    def display_mail(self,shape_name=None):
        
        folder_bin = get_mail_status(shape_name)[4][0]
        ms_list = self.get_email_obj(folder_bin)
        for ms in ms_list:
            if shape_name in ms.Subject:
                ms.Display()

    @classmethod
    def get_email_obj(self, folder_name = 'inbox'):

        outlook = cli.Dispatch("Outlook.Application").GetNamespace("MAPI") # 아웃룩

        if folder_name == 'inbox':

            argu_folder = outlook.GetDefaultFolder(6) # 받은편지함
        else:
            argu_folder = outlook.GetDefaultFolder(6).Parent.Folders(folder_name) # 1번폴더


        msg = argu_folder.Items #메세지 정보
        part_request = []
        service_request = 'SVC'
        for ms in msg:
            if service_request in ms.Subject:
                part_request.append(ms)
        # 서비스만 추출하기

        return part_request



# 통합제어에서 프린트 클릭하면 폼채워서 프린트하기
def print_form(shape_name=None,print_form_dir = "C:\\Users\\lms46\\Desktop\\fulfill\\print_form.xlsx",print_met=None):
    """
    메일제목, print_form.xlsx 절대경로 
    """
    wb_cy = xw.Book('cytiva.xlsm')
    ws_main= wb_cy.sheets['통합제어']
    shape_name = shape_name.replace('_prfm','')

    # ws_main = wb_cy.sheets['통합제어']
    ml_status = ['REQUESTED','PROCESSING', 'SHIPPED']
    ml_folders = ['inbox','1_Requests', '2_Processing', '3_ShipConfirmed']



    ## print_form이 켜져있지 않을 경우
    try :
        wb_pf = xw.Book('print_form.xlsx')
    except:
        wb_pf = xw.Book(print_form_dir)
    ws_svc = wb_pf.sheets['SVC']
    bin_loc = wb_cy.sheets['Cytiva Inventory BIN']
    
    
    
    folder_bin = get_mail_status(shape_name)[4][0]
    ms_list = Email.get_email_obj(folder_bin)
    
    part_request = []
    for ms in ms_list:
        if shape_name in ms.Subject:
            part_request.append(ms)
            
    
    if 'SVC' in part_request[0].Subject:
    
        #ws_svc에 폼입력하는 모듈
        __forming_datas(ws_svc, bin_loc, part_request,shape_name)
        
        # 서비스요청사항 프린트하기
        ws_svc.range("A1:H43").api.PrintPreview()
        
        # 만약 프린트만 요구하는 것이라면 여기서 끝..

        # 해당 index의 df을 DB로부터 읽고, 통합제어 시트의에서 해당 요청사항의 row번호를 찾는다.
        current_df = get_mail_status(shape_name)
        last_row = get_empty_row(col='J')
        idx_list = ws_main.range((12,"J"),(last_row-1,'J')).options(numbers=int).value
        idx_nm = idx_list.index(shape_name)
        # SHIP_CONFIRM "S" 버튼생성

        # ship_confirm 기능을 수행하는 버튼을 생성한다.
        sh_conf= ws_main.range(12+idx_nm,19)
        cel_left_sf = sh_conf.left
        cel_top_sf = sh_conf.top
        cel_width_sf = sh_conf.width
        cel_height_sf = sh_conf.height

        ws_main.api.Shapes.AddShape(125, cel_left_sf, cel_top_sf, cel_width_sf, cel_height_sf)

        ws_main.shapes[-1].name = shape_name+'_shcf'
        # 생성된 ship_confrim 버튼 객체
        obj_shcf = ws_main.shapes[shape_name+'_shcf']
        obj_shcf.text = '발송하기'
        obj_shcf.api.TextFrame.HorizontalAlignment = 2
        obj_shcf.api.TextFrame.VerticalAlignment = 2
        obj_shcf.api.OnAction = 'ship_confirm_mail'


        # 해당 메일을 1_Requests 에서 2_Processing로 이동한다.
        move_mail(ml_folders[1],ml_folders[2],ml_index=shape_name)
        ml_status_to = ml_status[1]
        bin_folder_to = ml_folders[2]
        
        # 현재시간
        # index 는 None이어야 자동 증분
        current_df[0][0] = None
        # 메일 상태 db 업데이트
        # 상태 업데이트
        current_df[2][0] = ml_status_to
        # 업데이트시간
        current_df[3][0] = get_current_time()
        # 아웃룩상 폴더
        current_df[4][0] = bin_folder_to
        MailStatus.update_status(df_ms=current_df)


        # excel상의 status 바꾸기 ml_status[1]
        ws_main.range((12+idx_nm,'Q')).value =ml_status[1]


def ship_confirm_mail(shape_name =None):
    wb_cy = xw.Book('cytiva.xlsm')
    ws_main= wb_cy.sheets['통합제어']
    # shape_name = shape_name.replace('_prfm','')
    md_index= shape_name.replace('_shcf','')

    ml_status = ['REQUESTED','PROCESSING', 'SHIPPED']
    ml_folders = ['inbox','1_Requests', '2_Processing', '3_ShipConfirmed']

        # 해당 index의 df을 DB로부터 읽고, 통합제어 시트의에서 해당 요청사항의 row번호를 찾는다.
    current_df = get_mail_status(md_index)
    last_row = get_empty_row(col='J')
    idx_list = ws_main.range((12,"J"),(last_row-1,'J')).options(numbers=int).value
    idx_nm = idx_list.index(md_index)
    xl_state_rng = ws_main.range((12+idx_nm,'Q'))

    if (current_df[2][0] == ml_status[1]) and (xl_state_rng.options(numbers=int).value== ml_status[1]) :
        #excel stauts 변경
        xl_state_rng.value = ml_status[2]

        # DB status 변경
        list_c_df = list(current_df.loc[0])
        list_c_df[0] = None
        # 상태 업데이트
        list_c_df[2] = ml_status[2]
        # 업데이트시간
        list_c_df[3] = get_current_time()
        # 아웃룩상 폴더
        list_c_df[4] = ml_folders[3]
        df_ms = pd.DataFrame(list_c_df).T
        MailStatus.update_status(df_ms=df_ms)

        obj_shcf = ws_main.shapes[shape_name]
        obj_shcf.text = '발송완료'

        move_mail(ml_folders[2],ml_folders[3],ml_index=md_index)

    else : 
        wb_cy.app.alert('STATUS가 '+ml_status[1]+' 인 상황에서만 ShipConfirm이 가능합니다. 진행을 취소합니다.',
        'Warning',mode='critical')
 

def __forming_datas(ws_svc, bin_loc, part_request,shape_name=None):
    body_exam = part_request[0].Body
    body = body_exam[:body_exam.rfind('}')+1]
    json_body = json.loads(body)

    json_body.keys()

    form_data= json_body['data']

    df_parts = pd.DataFrame(form_data['parts'])
        ## BIN 컬럼채우기
        
    bin_last_row = bin_loc.range('A100000').end('up').row
    bin_last_col = bin_loc.range('AAA'+str(bin_last_row)).end('left').column
    df_bin= pd.DataFrame(bin_loc.range((2,1),(bin_last_row,bin_last_col)).options(numbers=int).value)
    dict_bin = dict(zip(df_bin[0].astype(str),df_bin[1]))
    bin_list = []
    for part_no in df_parts[df_parts.columns[0]]:
        bin_list.append(dict_bin[part_no])
    df_parts['BIN'] = bin_list
        ## BIN 컬럼채우기

        # Index

    ws_svc.range('B21:H40').clear_contents()


        # 요청출고의 인덱스번호로결정 1 로가정
    ws_svc.range('E2').value = shape_name
        # Printed Day E3
    ws_svc.range('E3').value = get_current_time()
        # Delievery Type E5
    is_return = form_data['is_return']
    if is_return == 0 :
        is_return = None
    else :
        is_return = '왕복'
    ws_svc.range('E5').value = is_return
    ws_svc.range('E3').value = get_current_time()
        # Recipient B6
    ws_svc.range('B6').value = form_data['recipient']
        # Contact E6
    ws_svc.range('E6').value = form_data['contact']
        # Request day A10
    ws_svc.range('A10').value  = form_data['req_day']
        # request hour A11
    ws_svc.range('A11').value  = form_data['req_time']
        # Address D10
    ws_svc.range('D10').value = form_data['address']
        # Deilevery Instructions A14
    ws_svc.range('A14').value = form_data['del_instruction']
        # parts_info A20
    ws_svc.range('A20').value = df_parts
        # Depart From B41
    ws_svc.range('B41').value = '서울시 강서구 하늘길 247 3층 C구역'
        # TEL B42
    ws_svc.range('B42').value = '02-2660-3767'
        # Phone B43
    ws_svc.range('B43').value = '담당자 폰번호'
        # URGENT E41
    is_urgent = form_data['is_urgent']
    if is_urgent == 0 :
        is_urgent = None       
    else :
        is_urgent = '긴급' 
    ws_svc.range('E41').value = is_urgent
        








# 메일 폴더간 이동 모듈 
def move_mail(from_fd,to_fd,req_type="SVC",ml_index=None):
    """
    메일 폴더간 이동 모듈///
    """
    outlook = cli.Dispatch("Outlook.Application").GetNamespace("MAPI") # 아웃룩
    inbox = outlook.GetDefaultFolder(6) # 받은편지함
    request_folder = outlook.GetDefaultFolder(6).Parent.Folders('1_Requests') # 1번폴더
    process_folder = outlook.GetDefaultFolder(6).Parent.Folders('2_Processing') # 2번폴더
    shipped_folder = outlook.GetDefaultFolder(6).Parent.Folders('3_ShipConfirmed') # 3번폴더

    
    ml_folders = ['inbox','1_Requests', '2_Processing', '3_ShipConfirmed']

    fd_dict = {

        ml_folders[0]:inbox, ml_folders[1]:request_folder, ml_folders[2]:process_folder,ml_folders[3]:shipped_folder
    }


    from_fd = fd_dict[from_fd]
    to_fd = fd_dict[to_fd]


    part_request = []
    if ml_index == None:
        for ms in from_fd.Items:
            if req_type in ms.Subject:
                part_request.append(ms)
    else :
        for ms in from_fd.Items:
            if ml_index in ms.Subject:
                part_request.append(ms)
        
    # from_folder내용이 없을 경우
    if len(part_request) == 0:
        return None
    else :
        for ms in part_request:
            if req_type in ms.Subject:
                ms.Move(to_fd)




# 메일제목을 매개변수로 받아 현재 상태를 pd.Dataframe형태로 반환 한다.
# 현재는 한 개의 메일만 받을 수 있음
def get_mail_status(ml_sub):
    """
    메일제목을 argument로 받으면 현재상태의 해당 메일의 현재 ML_BIN, STATUS를 반환 
    """
    ml_sub = ml_sub.replace('_prfm','')
    ml_sub = ml_sub.replace('_shcf','')
    cur = DataWarehouse()
    query = 'select * from MAIL_STATUS where ML_SUB = (:name1)'
    db_obj = cur.execute(query, name1= ml_sub)
    df_status = pd.DataFrame(db_obj.fetchall())
    # 최신 날짜로 업데이트된 부분만 가져오기 MAIL_STATUS의 detail 부분은 제거
    # UP_TIME 컬럼 datetime 객체로 변환 후 내림차순으로 최신 업데이트 내역을 위에서 볼 수 있도록함
    df_status[3] = pd.to_datetime(df_status[3], format='%Y-%m-%d %H:%M:%S', errors='raise')
    df_status = df_status.sort_values(3,ascending=False).iloc[[0]].reset_index(drop=True)
    return df_status