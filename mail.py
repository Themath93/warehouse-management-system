from datetime import datetime
from xl_utils import get_empty_row
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
            req_type = ms.Subject.split('_')[0]
            body = ms_body[:ms_body.rfind('}')+1]
            json_body = json.loads(body)
            dict_data = json_body['data']

            # 메일 내용 채우기
            # index는 1부터 시작
            self.WS_MAIN.range('J'+str(get_empty_row(self.WS_MAIN,'J'))).value = idx + 1
            # print_form용 번호도 추가
            self.WS_MAIN.range('R'+str(get_empty_row(self.WS_MAIN,'R'))).value = idx + 1
            # 요청타입 K
            self.WS_MAIN.range('K'+str(get_empty_row(self.WS_MAIN,'K'))).value = req_type = ms.Subject.split('_')[0]
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

        ms_list = self.get_email_obj()

        for ms in ms_list:
            if shape_name in ms.Subject:
                ms.Display()

    def get_email_obj():

        outlook = cli.Dispatch("Outlook.Application").GetNamespace("MAPI") # 아웃룩
        inbox = outlook.GetDefaultFolder(6) # 받은편지함
        msg = inbox.Items #메세지 정보

        part_request = []
        service_request = 'SVC'
        for ms in msg:
            if service_request in ms.Subject:
                part_request.append(ms)
        # 서비스만 추출하기

        return part_request



# 통합제어에서 프린트 클릭하면 폼채워서 프린트하기
def print_form(subject=None):
    """
    """
    wb_cy = xw.Book('cytiva.xlsm')
    wb_pf = xw.Book('print_form.xlsx')
    ws_svc = wb_pf.sheets['SVC']
    bin_loc = wb_cy.sheets['Cytiva Inventory BIN']



    subject = subject.replace('_prfm','')

    
    outlook = cli.Dispatch("Outlook.Application").GetNamespace("MAPI") # 아웃룩
    inbox = outlook.GetDefaultFolder(6) # 받은편지함
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    part_request = []
    for ms in inbox.Items:
        if subject in ms.Subject:
            part_request.append(ms)
            
    
    if 'SVC' in part_request[0].Subject:
    

        __forming_datas(ws_svc, bin_loc, now, part_request)

        ws_svc.range("A1:H43").api.PrintPreview()





def __forming_datas(ws_svc, bin_loc, now, part_request):
    body_exam = part_request[0].Body
    body = body_exam[:body_exam.rfind('}')+1]
    body
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
    ws_svc.range('E2').value = 1
        # Printed Day E3
    ws_svc.range('E3').value = now
        # Delievery Type E5
    is_return = form_data['is_return']
    if is_return == 0 :
        is_return = None
    else :
        is_return = '왕복'
    ws_svc.range('E5').value = is_return
    ws_svc.range('E3').value = now
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
        


            
        
        