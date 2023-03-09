## xl_wings 절대경로 추가
import sys, os
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))


from datajob.xlwings_dj.shipment_information import ShipmentInformation
from datajob.xlwings_dj.local_list import LocalList
from datajob.xlwings_dj.so_out import SOOut

## 출고
from datetime import datetime
from xlwings_job.oracle_connect import DataWarehouse
from xlwings_job.xl_utils import bring_data_from_db, clear_form, get_each_index_num, get_idx, get_out_info, get_out_table, get_row_list_to_string, get_xl_rng_for_ship_date, row_nm_check, sht_protect
import xlwings as xw
import pandas as pd


wb_cy = xw.Book("cytiva_worker.xlsm").set_mock_caller()
wb_cy = xw.Book.caller()

class ShipReady():
    SHEET_NAMES =  ['Temp_DB', 'SHIPMENT_INFORMATION', 'POD', 
    'LOCAL_LIST', 'IR_ORDER','SVC_BIN','MAIN']

    STATUS = ['waiting_for_out', 'ship_is_ready', '_is_empty','local_out_row_input_required', 'edit_mode']

    PRODUCT_STATUS = ['HOLDING']

    SEL_SHT = wb_cy.selection.sheet

    STATUS_CELL = SEL_SHT.range("H4")

    WS_DB = wb_cy.sheets[SHEET_NAMES[0]]
    WS_SI = wb_cy.sheets[SHEET_NAMES[1]]
    WS_POD = wb_cy.sheets[SHEET_NAMES[2]]
    WS_LC = wb_cy.sheets[SHEET_NAMES[4]]
    WS_SVMX = wb_cy.sheets[SHEET_NAMES[5]]
    WS_MAIN = wb_cy.sheets[SHEET_NAMES[-1]]

    @classmethod
    def ship_ready(self):
        # self.WS_DB.range("U2") ==> ws_si 시트를 위한 임시 값 저장소
        table_name = self.WS_SI.range("D5").value
        cur=DataWarehouse()
        row_cel = self.WS_SI.range("C2")
        out_rows_str=get_row_list_to_string(row_nm_check(wb_cy)['selection_row_nm'])
        # waiting_for_out 에서만 사용가능!
        if self.STATUS[2] in self.STATUS_CELL.value:

            self.check_local_empty()
        elif self.STATUS_CELL.value != self.STATUS[0] :

            wb_cy.app.alert(self.STATUS_CELL.value+"에서는 출고기능 사용이 불가합니다.",'INFO')
            return
        
        # 이미 출고가능은 확인된 상태에서는 더이상 출고번호를 입력받을 필요는 없음
 

        elif self.STATUS_CELL.value == self.STATUS[0] :
            # STATUS CHECK
            # HOLDING, BRANCH, DMG, WRONG
            out_row_list=get_out_table(direct_call=True)
            db_row_list = []
            for si_index in out_row_list:
                db_row_list.append(cur.execute(f"SELECT SI_INDEX,ARRIVAL_DATE,SHIP_DATE,STATE,REMARK FROM {table_name} WHERE SI_INDEX = '{si_index}'").fetchone())
            df_sel =  pd.DataFrame(db_row_list,columns=['SI_INDEX','ARRIVAL_DATE','SHIP_DATE','STATE','REMARK'])
            
            check_list = []
            check_list.append(len(df_sel[df_sel['STATE'] == "HOLDING"]) > 0)
            check_list.append(len(df_sel[df_sel['REMARK'] == "대리점"]) > 0)
            check_list.append(df_sel['SHIP_DATE'].count(None) != 0)
            check_available = check_list.count(True)

            if check_available > 0 :
                wb_cy.app.alert("""선택하신 품목은 출고가 불가합니다. 아래내용을 참고해주세요. 
                                \n 1. STATE가 HOLDING 인경우 ** (검수가 안된 품목일 수 있습니다. 해당 품목에 대한 검수를 마쳐주세요!)
                                \n 2. 대리점 품목인 경우 **대리점용 출고시스템을 따로사용해주세요 (ONLY FOR BRANCH)
                                \n 3. 이미 출고가 된 품목인 경우""",
                                "WARNING")
                return
            
            is_arrived = len(df_sel[~pd.notna(df_sel['ARRIVAL_DATE'])]) > 0
            is_damaged = len(df_sel[df_sel['STATE'].str.contains("DMG")]) > 0
            is_wrong = len(df_sel[df_sel['STATE'].str.contains("WR_")]) > 0
            if is_arrived == True:
                ready_answer = wb_cy.app.alert("아직 입고가 되지 않은 품목이 포함되어 있습니다. 계속진행 하시겟습니까?","Ship Ready Confirm",buttons='yes_no_cancel')
                if ready_answer != 'yes': #출고 안하겠다는말
                    wb_cy.app.alert("Ship Ready를 종료합니다.","Ship Ready Confirm")
                    return
                
            if is_damaged == True:
                ready_answer = wb_cy.app.alert("검수 중 Damamged 품목이 포함되어 있습니다. 계속진행 하시겟습니까?","Ship Ready Confirm",buttons='yes_no_cancel')
                if ready_answer != 'yes': #출고 안하겠다는말
                    wb_cy.app.alert("Ship Ready를 종료합니다.","Ship Ready Confirm")
                    return
                
            if is_wrong == True:
                ready_answer = wb_cy.app.alert("검수 중 Wrongshipped 품목이 포함되어 있습니다 계속진행 하시겟습니까?","Ship Ready Confirm",buttons='yes_no_cancel')
                if ready_answer != 'yes': #출고 안하겠다는말
                    wb_cy.app.alert("Ship Ready를 종료합니다.","Ship Ready Confirm")
                    return            
            row_cel.value = out_rows_str
            self.check_local_empty()



    
    @classmethod
    def only_local(self):
        """
        SO출고시 Local품목만 따로 출고하는경우
        """

        

        #edit_mode인 경우에는 출고 자체금지!
        if self.STATUS_CELL.value != self.STATUS[4] :
            
            self.WS_SI.range("C2").value = 'only_local'
           
            self.check_local_empty()
            if self.WS_SI.range("H4").value == self.STATUS[3] :
                
                local_check(self.WS_SI,self.WS_LC)
        
        else : 
            wb_cy.app.alert(self.STATUS_CELL.value+"에서는 출고기능 사용이 불가합니다.",'안내')

    @classmethod
    def check_local_empty(self):
        si_sht_list = self.WS_SI.range('C2:C6').value
        lc_out_row = self.WS_SI.range('C7').value
            
        si_sht_list_name = self.WS_SI.range('B2:B6').value
            
        none_list = []
        must_fill = []
        for index, val in enumerate(si_sht_list):
            if val == None:
                none_list.append(index)
                    
            # 필수 입력분이 누락되었을 경우
        if len(none_list) > 0 :
            for val in none_list:
                must_fill.append(si_sht_list_name[val])

            self.WS_SI.range("H4").value = (', '.join(must_fill)) + "_is_empty"
            self.WS_SI.range("H4").color = (255,0,0)
            ## must_fill == [] 일 경우 출고를 시작 할 수 있게된다.


        #로컬만 출고인경우 only_local
        elif si_sht_list[0] == 'only_local' :

            # 로컬출고행이 이미 정해져 있는 경우
            if lc_out_row == None:
                # self.WS_SI.range("H4").value = self.STATUS[3]
                # self.WS_SI.range("H4").color = (255,255,0)
                self.WS_LC.activate()
                wb_cy.app.alert('로컬출고행을 여기서 정한후 ConfirmLocal버튼을 눌러주세요',
                                'Local Check')
                clear_form(self.WS_LC)

            else :
                self.WS_SI.range("H4").value = self.STATUS[1]
                self.WS_SI.range("H4").color = (0,255,255)

                


            # Local출고 체크
        else:
            self.WS_SI.range("H4").value = self.STATUS[3]
            self.WS_SI.range("H4").color = (255,255,0)
            local_check(self.WS_SI,self.WS_LC)




    @classmethod
    def local_ready(self):
        # self.WS_DB.range("U2") ==> ws_si 시트를 위한 임시 값 저장소
        tmp_lc_address = self.WS_DB.range("U3")
        local_row_nums = get_row_list_to_string(row_nm_check(wb_cy)['selection_row_nm'])

        ## yes == 1, no ==0
        confirm_local = wb_cy.app.alert("행 번호 : " + local_row_nums + " 출고가 맞습니까?","Local Check",buttons='yes_no_cancel')



        if confirm_local == 'yes' :
            # local_only
            # 로컬에서 먼저 로컬출고행을 고른다음 ws_si에서 이후 출고진행
            if self.WS_LC.range("C7").value == None :
                self.WS_LC.range("C2").value = local_row_nums
                self.WS_LC.range("H4").color = (0,255,255)
                self.WS_LC.range("H4").value = self.STATUS[1]
                self.WS_SI.activate()
                self.WS_SI.range("C7").value = local_row_nums
                self.WS_SI.range("C2").value = 'only_local'
                tmp_lc_address.value = get_xl_rng_for_ship_date(ship_date_col_num="J")
                self.check_local_empty()

            else : 
                print("로컬 출고 확인")
                self.WS_LC.range("C2").value = local_row_nums
                tmp_lc_address.value = get_xl_rng_for_ship_date(ship_date_col_num="J")
                self.WS_LC.range("H4").color = (0,255,255)
                self.WS_LC.range("H4").value = self.STATUS[1]
                self.WS_SI.activate()
                self.WS_SI.range("C7").value = local_row_nums
                self.WS_SI.range("H4").color = (0,255,255)
                self.WS_SI.range("H4").value = self.STATUS[1]

        elif confirm_local == 'no' :
            self.WS_LC.range("C2").clear_contents()
            self.WS_LC.range("H4").value = self.STATUS[3]
            print("로컬 출고 행 번호가 아님")

        elif confirm_local == 'cancel' :
            print("취소")
    







# ShipConform 기능
class ShipConfirm():
    """
    모든 품목에 대하여 출고 확정 기능을 담당
    """
    SHEET_NAMES =  ['Temp_DB', 'SHIPMENT_INFORMATION', 'POD', 
    'LOCAL_LIST', 'IR_ORDER','SVC_BIN','MAIN']


    STATUS = ['waiting_for_out', 'ship_is_ready', '_is_empty','local_out_row_input_required', 'edit_mode']

    WS_DB = wb_cy.sheets[SHEET_NAMES[0]]
    WS_SI = wb_cy.sheets[SHEET_NAMES[1]]
    WS_POD = wb_cy.sheets[SHEET_NAMES[2]]
    WS_LC = wb_cy.sheets[SHEET_NAMES[4]]
    WS_SVMX = wb_cy.sheets[SHEET_NAMES[5]]
    WS_MAIN = wb_cy.sheets[SHEET_NAMES[-1]]



    @classmethod
    def ship_confirm(self):
        cur = DataWarehouse()
        wb_cy.app.screen_updating = False
        # SO 품목 출고 담당

        # so,lc 임시 주소저장함 확인
        tmp_si_address = self.WS_DB.range("U2")
        tmp_lc_address = self.WS_DB.range("U3")

        # 선택한 시트 객체
        sel_sht = wb_cy.selection.sheet
        del_method = sel_sht.range("C5").value
        is_no_local = sel_sht.range("C7").value == 'no_local'
        if (is_no_local == True) & (del_method == '택배'):
            col_names = list(pd.DataFrame(cur.execute(f"select column_name from user_tab_columns where table_name = upper('SHIPMENT_INFORMATION')").fetchall())[0])
            si_idx_list = get_out_table()
            db_row_list = []
            for idx in si_idx_list:
                qry = f"SELECT * FROM SHIPMENT_INFORMATION WHERE SI_INDEX = '{idx}'"
                db_row_list.append(cur.execute(qry).fetchone())
            df_si = pd.DataFrame(db_row_list,columns=col_names)
            parcel_list = list(set(df_si['PARCELS_NO']))
            while True:
                courier_num_list = input_delivery_invoice_number_for_out_bound(parcel_list)
                pacels_dict = dict(zip(parcel_list,courier_num_list))
                alert_str = str(pacels_dict).replace(",","\n").replace('{','').replace('}',"")
                answer = wb_cy.app.alert("당신의 입력 : \n"+alert_str+" \n 수정하시겠습니까?","CONFIRM",buttons="yes_no_cancel")
                if answer == 'no':
                    del_method = pacels_dict
                    break
                if answer == 'cancel':
                    wb_cy.app.alert("종료합니다.","Quit")
                    return
        elif (is_no_local == False) & (del_method == '택배') :
            wb_cy.app.alert("택배배송은 로컬품목없이 SO건만 출고시 사용 가능합니다. 배송방법을 다시선택해주세요.","Quit")
            return




        status_cel = sel_sht.range("H4")


        ## ship_is_ready 상태에서만 ship_confirm기능 사용가능
        if status_cel.value == self.STATUS[1] :
            tmp_idx = str(self.WS_DB.range("C500000").end('up').row + 1)
            tmp_list = get_out_info(sel_sht)
            ship_date = tmp_list[3]
            # si 시트 출고 시 local품목이 있을 경우
            if (sel_sht.name == self.SHEET_NAMES[1]) and (tmp_list[-3] != 'no_local') :
                self.__create_soout_index_insert_data_to_db(sel_sht, status_cel, tmp_idx)

                local_idx = self.WS_DB.range("I3").value
                # ws_si 및 ws_lc 내용  db update

                
                if tmp_list.count('only_local') == 0 : #only_local일 경우에는 shipment update는필요없음
                    ShipmentInformation.update_shipdate(get_each_index_num=get_each_index_num(tmp_list[1]),ship_date=ship_date)
                
                LocalList.update_shipdate(get_each_index_num=get_each_index_num(local_idx),ship_date=ship_date)
                
                tmp_si_address.clear_contents()
                tmp_lc_address.clear_contents()
                clear_form(self.WS_LC)
                clear_form()
                
                self.WS_LC.select()
                bring_data_from_db(in_method=True)
                self.WS_SI.select()

            else :
                # 로컬품목없이 ws_si 품목만 출고
                self.__create_soout_index_insert_data_to_db(sel_sht, status_cel, tmp_idx)
                
                
                # xl_column date update ==> si_sht
                sel_sht.range(tmp_si_address.value).value = ship_date
                # ws_si 내용 db update
                ShipmentInformation.update_shipdate(get_each_index_num=get_each_index_num(tmp_list[1]),ship_date=ship_date,del_method=del_method)
                tmp_si_address.clear_contents()
                clear_form(self.WS_LC)
                clear_form()


            bring_data_from_db(in_method=True)
            wb_cy.app.screen_updating = True
        else  : 
            print("무응답")
            wb_cy.app.screen_updating = True








        

    @classmethod
    def __create_soout_index_insert_data_to_db(self, sel_sht, status_cel, tmp_idx):
        tmp_list = get_out_info(sel_sht)
        ## local 출고 품목이 있다면 local출고 index key화
        if (sel_sht.name == self.SHEET_NAMES[1]) and (tmp_list[-3] != 'no_local') :
            tmp_list[-3] = str(get_idx(self.WS_LC))
        
        self.WS_DB.range("C"+ tmp_idx).value = tmp_list
        SOOut.put_data(self,tmp_list)
        # 성공적으로 tmep_db에 저장이 된상태라면 출고양식들을 지워줄 필요가있다.




        







# utils for out_bound it self

def local_check(WS_SI,WS_LC):
    
    ans_local = wb_cy.app.alert("Local 출고 품목이 있습니까?","Local Check",buttons='yes_no_cancel')


    if ans_local == 'yes' :
        print("로컬품목 있음")
        WS_SI.range("H4").color = (255,255,0)
        WS_SI.range("H4").value = 'local_out_row_input_required'
        WS_SI.range("C7").value = "need_row_nums"
        ##여기서 로컬페이지 actvate 
        local_act_ob(WS_SI,WS_LC)
        
    elif ans_local == 'no' :
        print("로컬품목 없음")
        WS_SI.range("C7").value = "no_local"
        WS_SI.range("H4").color = (0,255,255)
        WS_SI.range("H4").value = 'ship_is_ready'
        
    elif ans_local == 'cancel' :
        print("출고 취소")
        WS_SI.range("C2:C7").clear_contents()
        WS_SI.range("C4").value = "=TODAY()+1"
        WS_SI.range("H4").color = (255,255,255)
        WS_SI.range("H4").value = "waiting_for_out"
        ## clear_form() 매서드호출 하여 초기화

def local_act_ob(WS_SI,WS_LC) : 

    WS_LC.range('C3').options(transpose=True).value = WS_SI.range("C3:C6").value
    WS_LC.range('C7').options(transpose=True).value = WS_SI.range("C2").value
    
    WS_LC.activate()

    WS_LC.range("H4").value = "local_out_row_input_required"
    WS_LC.range("H4").color = (0,255,255)

def input_delivery_invoice_number_for_out_bound(parcel_list):
    wb_cy.app.alert("Parcel_NO에 맞는 송장번호를 입력해주세요.","INFO")
    courier_num_list = []
    for i in range(len(parcel_list)):
        courier_num_list.append(wb_cy.app.api.InputBox(f"'{parcel_list[i]}'에 맞는 송장번호를 입력해주세요.","Delivery invoice number INPUT",Type=2))
    return courier_num_list