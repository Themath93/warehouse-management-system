## 출고
from datetime import datetime
from datajob.dw.so_out import SOOut
from xl_utils import clear_form, get_idx, get_out_info, get_row_list_to_string, row_nm_check
import xlwings as xw
import pandas as pd


wb_cy = xw.Book.caller()
# wb_cy = xw.Book('cytiva.xlsm')

class ShipReady():
    SHEET_NAMES =  ['Temp_DB', 'Shipment information', '인수증', 
    '대리점송장', '대리점 출고대기', '로컬리스트', 'In-Transit part report', '기타리스트',
     '출고리스트', 'Cytiva Inventory BIN']

    STATUS = ['waiting_for_out', 'ship_is_ready', '_is_empty','local_out_row_input_required', 'edit_mode']

    ACT_SEL = wb_cy.selection.sheet

    STATUS_COL = ACT_SEL.range('XFD4').end('left').column

    STATUS_CELL = ACT_SEL.range(4,STATUS_COL)

    WS_DB = wb_cy.sheets[SHEET_NAMES[0]]
    WS_SI = wb_cy.sheets[SHEET_NAMES[1]]
    WS_POD = wb_cy.sheets[SHEET_NAMES[2]]
    WS_BR = wb_cy.sheets[SHEET_NAMES[3]]
    WS_LC = wb_cy.sheets[SHEET_NAMES[5]]
    WS_SVMX = wb_cy.sheets[SHEET_NAMES[6]]
    WS_OTHER = wb_cy.sheets[SHEET_NAMES[7]]

    @classmethod
    def ship_ready(self):
        
        #edit_mode인 경우에는 출고 자체금지!
        if self.STATUS_CELL.value != self.STATUS[4] :

            out_rows=get_row_list_to_string(row_nm_check(wb_cy)['selection_row_nm'])

            # 출고행 번호에 값이 있을 경우
            if self.WS_SI.range("C2").value == None:

                self.WS_SI.range("C2").value = out_rows
                
                self.check_local_empty()
            else :
            # 출고행 번호에 값이 없을 경우

                self.check_local_empty()
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

        local_row_nums = get_row_list_to_string(row_nm_check(wb_cy)['selection_row_nm'])

        ## yes == 1, no ==0
        confirm_local = wb_cy.app.alert("행 번호 : " + local_row_nums + " 출고가 맞습니까?","Local Check",buttons='yes_no_cancel')



        if confirm_local == 'yes' :
            # local_only
            # 로컬에서 먼저 로컬출고행을 고른다음 ws_si에서 이후 출고진행
            if self.WS_LC.range("C2").value == None :
                self.WS_LC.range("C2").value = local_row_nums
                self.WS_LC.range("E4").color = (0,255,255)
                self.WS_LC.range("E4").value = self.STATUS[1]
                self.WS_SI.activate()
                self.WS_SI.range("C7").value = local_row_nums
                self.WS_SI.range("C2").value = 'only_local'
                self.check_local_empty()

            else : 
                print("로컬 출고 확인")
                self.WS_LC.range("C2").value = local_row_nums
                self.WS_LC.range("E4").color = (0,255,255)
                self.WS_LC.range("E4").value = self.STATUS[1]
                self.WS_SI.activate()
                self.WS_SI.range("C7").value = local_row_nums
                self.WS_SI.range("H4").color = (0,255,255)
                self.WS_SI.range("H4").value = self.STATUS[1]

        elif confirm_local == 'no' :
            self.WS_LC.range("C2").clear_contents()
            self.WS_LC.range("E4").value = self.STATUS[3]
            print("로컬 출고 행 번호가 아님")

        elif confirm_local == 'cancel' :
            print("취소")
    







# ShipConform 기능
class ShipConfirm():
    """
    모든 품목에 대하여 출고 확정 기능을 담당
    """
    SHEET_NAMES =  ['Temp_DB', 'Shipment information', '인수증', 
    '대리점송장', '대리점 출고대기', '로컬리스트', 'In-Transit part report', '기타리스트',
     '출고리스트', 'Cytiva Inventory BIN']

    STATUS = ['waiting_for_out', 'ship_is_ready', '_is_empty','local_out_row_input_required', 'edit_mode']

    WS_DB = wb_cy.sheets[SHEET_NAMES[0]]
    WS_SI = wb_cy.sheets[SHEET_NAMES[1]]
    WS_POD = wb_cy.sheets[SHEET_NAMES[2]]
    WS_BR = wb_cy.sheets[SHEET_NAMES[3]]
    WS_LC = wb_cy.sheets[SHEET_NAMES[5]]
    WS_SVMX = wb_cy.sheets[SHEET_NAMES[6]]
    WS_OTHER = wb_cy.sheets[SHEET_NAMES[7]]

    # @classmethod
    # def local_ship(self):
    #     """
    #     local만출고시 사용
    #     """



    @classmethod
    def ship_confirm(self):
        # SO 품목 출고 담당

        # 선택한 시트 객체
        sel_sht=wb_cy.selection.sheet
        status_cel = sel_sht.range("AAA4").end('left')


        ## ship_is_ready 상태에서만 ship_confirm기능 사용가능
        if status_cel.value == self.STATUS[1] :

            tmp_idx = str(self.WS_DB.range("C500000").end('up').row + 1)
            tmp_list = get_out_info(sel_sht)

            # si 시트 출고 시 local품목이 있을 경우
            if sel_sht.name == self.SHEET_NAMES[2] \
                    and tmp_list[-2] != 'no_local' :
                    
                tmp_list[-2] = 'lc_' + str(get_idx(self.WS_LC))
                self.__creatd_tmp_db_row(sel_sht, status_cel, tmp_idx)
                # 로컬도 같이 출고 했으니 로컬항목도 지울 것
            # 로컬만 출고일 경우
            # elif sel_sht.name == self.SHEET_NAMES[2] \
            #         and sel_sht.range("C2") == 'only_local' :
            #     tmp_list[-1] = 'lc_' + str(get_idx(self.WS_LC))
            #     self.__creatd_tmp_db_row(sel_sht, status_cel, tmp_idx)
                clear_form(self.WS_LC)
            else :
                self.__creatd_tmp_db_row(sel_sht, status_cel, tmp_idx)
        else  : 
            print("무응답")









        

    @classmethod
    def __creatd_tmp_db_row(self, sel_sht, status_cel, tmp_idx):
        self.WS_DB.range("C"+ tmp_idx).value = get_out_info(sel_sht)
        SOOut.put_data(self,get_out_info(sel_sht))
        # 성공적으로 tmep_db에 저장이 된상태라면 출고양식들을 지워줄 필요가있다.
        sel_sht.range("C2:C7").clear_contents()
        sel_sht.range("C4").value = '=TODAY()+1'
        status_cel.color = None
        status_cel.value = self.STATUS[0]



        







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

    WS_LC.range("E4").value = "local_out_row_input_required"
    WS_LC.range("E4").color = (0,255,255)