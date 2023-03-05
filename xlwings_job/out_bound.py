## xl_wings 절대경로 추가
import sys, os

sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))


from datajob.xlwings_dj.shipment_information import ShipmentInformation
from datajob.xlwings_dj.local_list import LocalList

## 출고
from datetime import datetime
from datajob.xlwings_dj.so_out import SOOut
from xlwings_job.xl_utils import bring_data_from_db, clear_form, get_each_index_num, get_idx, get_out_info, get_row_list_to_string, get_xl_rng_for_ship_date, row_nm_check, sht_protect
import xlwings as xw
import pandas as pd


wb_cy = xw.Book.caller()
# wb_cy = xw.Book('cytiva.xlsm')

class ShipReady():
    SHEET_NAMES =  ['Temp_DB', 'Shipment information', '인수증', 
    '대리점송장', '대리점 출고대기', '로컬리스트', 'IR_SVC', '기타리스트',
     '출고리스트', 'BIN']

    STATUS = ['waiting_for_out', 'ship_is_ready', '_is_empty','local_out_row_input_required', 'edit_mode']

    PRODUCT_STATUS = ['HOLDING']

    SEL_SHT = wb_cy.selection.sheet

    STATUS_CELL = SEL_SHT.range("H4")

    WS_DB = wb_cy.sheets[SHEET_NAMES[0]]
    WS_SI = wb_cy.sheets[SHEET_NAMES[1]]
    WS_POD = wb_cy.sheets[SHEET_NAMES[2]]
    WS_LC = wb_cy.sheets[SHEET_NAMES[5]]
    WS_SVMX = wb_cy.sheets[SHEET_NAMES[6]]
    WS_OTHER = wb_cy.sheets[SHEET_NAMES[7]]

    @classmethod
    def ship_ready(self):
        # self.WS_DB.range("U2") ==> ws_si 시트를 위한 임시 값 저장소
        tmp_si_address = self.WS_DB.range("U2")
        #edit_mode인 경우에는 출고 자체금지!
        if self.STATUS_CELL.value != self.STATUS[4] :
            # ship_ready기능을 사용하기 위해선 기본적으로 선택한 out_rows의 ship_date 컬럼이 None 값
            # 그리고 DB상의 ship_date 컬럼에서도 'None' 값이어야만한다.
            out_rows=get_row_list_to_string(row_nm_check(wb_cy)['selection_row_nm'])

            # status HOLDING 확인
            out_rows_xl_status = get_xl_rng_for_ship_date(ship_date_col_num="R")
            sel_rng_status = self.WS_SI.range(out_rows_xl_status).options(ndim=1)

            hold_count = sel_rng_status.value.count("HOLDING")

            # 선택한 행중 어느 하나라도 Status가 HOLDING 상태임
            if hold_count > 0 : 
                wb_cy.app.alert("해당 품목의 STATUS가 'HOLDING'이기 때문에 진행이 불가합니다. 매서드를 종료합니다.","Ship Ready WARNING")
                return None

            # # arrval_date 입고안된 파트 출고인지 확인
            # out_rows_xl_ad = get_xl_rng_for_ship_date(ship_date_col_num="K")
            # sel_rng_ad = self.WS_SI.range(out_rows_xl_ad).options(ndim=1)
            # wb_cy.app.alert(str(sel_rng_ad))
            # # None의개수가 1개이상이면 입고안된 품목에대해 ship_ready를 요청한 것
            # cnt_none_ad = sel_rng_ad.value.count(None)

            # if cnt_none_ad > 0 :
            #     ready_answer = wb_cy.app.alert("아직 입고가 되지 않은 파트가 신청 되었습니다. 계속진행 하시겟습니까?","Ship Ready Confirm",buttons='yes_no_cancel')
            #     if ready_answer != 'yes': #출고 안하겠다는말
            #         wb_cy.app.alert("Ship Ready를 종료합니다.","Ship Ready Confirm")
            #         return None
            
            

            ### ws_si 시트상의 ship_date에 값이 없는지 확인 
            out_rows_xl = get_xl_rng_for_ship_date(ship_date_col_num="L")
            # ndim=1 => value가 하나일때도 list로 value 반환
            sel_rng = self.WS_SI.range(out_rows_xl).options(ndim=1)

            # try : 
            cnt_none = sel_rng.value.count(None)
            len_rng =len(sel_rng.value)
            # except :
            #     cnt_none = sel_rng.value
            #     len_rng = None

            if (cnt_none == len_rng) :

                # 출고행 번호에 값이 없을 경우
                if self.WS_SI.range("C2").value == None:

                    self.WS_SI.range("C2").value = out_rows
                    
                    tmp_si_address.value = get_xl_rng_for_ship_date(ship_date_col_num="L")
                                # arrval_date 입고안된 파트 출고인지 확인
                    out_rows_xl_ad = get_xl_rng_for_ship_date(ship_date_col_num="K")
                    sel_rng_ad = self.WS_SI.range(out_rows_xl_ad).options(ndim=1)
                    # None의개수가 1개이상이면 입고안된 품목에대해 ship_ready를 요청한 것
                    cnt_none_ad = sel_rng_ad.value.count(None)

                    if cnt_none_ad > 0 :
                        ready_answer = wb_cy.app.alert("아직 입고가 되지 않은 파트가 신청 되었습니다. 계속진행 하시겟습니까?","Ship Ready Confirm",buttons='yes_no_cancel')
                        if ready_answer != 'yes': #출고 안하겠다는말
                            wb_cy.app.alert("Ship Ready를 종료합니다.","Ship Ready Confirm")
                            return None
                    
                    self.check_local_empty()
                else :
                # 출고행 번호에 값이 있을 경우

                    self.check_local_empty()
            #ship_date에 날짜 또는 값이 있을 경우
            else:
                wb_cy.app.alert("SHIP_DATE 컬럼에 값이 있습니다. 중복출고는 불가합니다." , "Mutiple Out WARNING")
        else :
            wb_cy.app.alert(self.STATUS_CELL.value+"에서는 출고기능 사용이 불가합니다.",'안내')
    
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
    SHEET_NAMES =  ['Temp_DB', 'Shipment information', '인수증', 
    '대리점송장', '대리점 출고대기', '로컬리스트', 'IR_SVC', '기타리스트',
     '출고리스트', 'BIN']

    STATUS = ['waiting_for_out', 'ship_is_ready', '_is_empty','local_out_row_input_required', 'edit_mode']

    WS_DB = wb_cy.sheets[SHEET_NAMES[0]]
    WS_SI = wb_cy.sheets[SHEET_NAMES[1]]
    WS_POD = wb_cy.sheets[SHEET_NAMES[2]]
    WS_LC = wb_cy.sheets[SHEET_NAMES[5]]
    WS_SVMX = wb_cy.sheets[SHEET_NAMES[6]]
    WS_OTHER = wb_cy.sheets[SHEET_NAMES[7]]



    @classmethod
    def ship_confirm(self):
        wb_cy.app.screen_updating = False
        # SO 품목 출고 담당

        # so,lc 임시 주소저장함 확인
        tmp_si_address = self.WS_DB.range("U2")
        tmp_lc_address = self.WS_DB.range("U3")

        # 선택한 시트 객체
        sel_rng = wb_cy.selection
        sel_sht = sel_rng.sheet

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
                sht_protect(False)
                bring_data_from_db()
                sht_protect()
                self.WS_SI.select()

            else :
                # 로컬품목없이 ws_si 품목만 출고
                self.__create_soout_index_insert_data_to_db(sel_sht, status_cel, tmp_idx)
                
                
                # xl_column date update ==> si_sht
                sel_sht.range(tmp_si_address.value).value = ship_date
                # ws_si 내용 db update
                ShipmentInformation.update_shipdate(get_each_index_num=get_each_index_num(tmp_list[1]),ship_date=ship_date)
                tmp_si_address.clear_contents()
                clear_form(self.WS_LC)
                clear_form()


            sht_protect(False)
            bring_data_from_db()
            sht_protect()
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