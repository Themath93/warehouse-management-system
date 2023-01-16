from datetime import datetime
import pandas as pd
from dicts_cy import return_dict
import xlwings as xw


wb_cy = xw.Book.caller()
# wb_cy = xw.Book('cytiva.xlsm')


## 선택한 행, 시트 이름 딕셔너리로 반환
def row_nm_check(xw_book_name=xw.books[0]):
    """
    Return Dict, get activated sheet's name(str), get selected row's number(list)
    """
    ## SpecialCells(12) 셀이 한개만 클릭됬을 경우에는 제대로 작동 불가

    # 셀한개만 클릭할경우 $의 개수가 2개이고 2개이상일경우 $의 개수는 4개 이다.
    sel_cell = xw_book_name

    count_dollar = sel_cell.selection.address.count("$")
    
    # 셀한개만 클릭할경우
    if count_dollar == 2 :
        
        sel_rng = [sel_cell.selection.address]
    else : 
    
        sel_rng = xw_book_name.selection.api.SpecialCells(12).Address.split(",")
    
    range_list = []
    
    for rng in sel_rng:

        if ":" in rng:
            num_0 = rng.split(":")[0].split("$")[-1]
            num_1 = rng.split(":")[1].split("$")[-1]
            #범위선택이 한 컬럼이 아닌 여러 컬럼을 동시에 선택된 경우
            if num_0 == num_1 :
                range_list.append(num_0)
            else : 
                range_list.append(str(num_0) + ',' + str(num_1))
        else :
            range_list.append(rng.split("$")[-1])
            
    dict_row_num_sheet_name = {"sheet_name":xw_book_name.selection.sheet.name,'selection_row_nm':range_list}
    
    
    return dict_row_num_sheet_name


## 출고행 번호 리스트 -> 문자열로 변환
def get_row_list_to_string(seleted_row_list) :
    return ' '.join(seleted_row_list).replace(',','~').replace(' ', ', ')


## 출고 form 초기화 및 출고 대기상태로 변경
def clear_form(sheet = wb_cy.selection.sheet):
    current_sheet = sheet
    
    #통합제어 시트일 경우
    if current_sheet.name == '통합제어' :

        main_clear('SVC')
    else :

        
        status_cel = current_sheet.range("AAA4").end('left')

        current_sheet.range("C2:C7").clear_contents()
        status_cel.color = None
        status_cel.value = "waiting_for_out"
        current_sheet.range("C4").value = "=TODAY()+1"


def main_clear(type=None):

    ws_main =wb_cy.sheets['통합제어']

    if type != None:
        
        shapes = ws_main.shapes
        last_row = get_empty_row(ws_main,'J')
        for shp in shapes :
            if type in shp.name :
                shp.delete()

        ws_main.range("J12:R"+str(last_row)).clear_contents()
    else :
        pass



def get_idx(sheet_name):
    """
    xlwings.main.Sheet를 인수로 입력
    """
    # 연속된 숫자 표현

    # 'c'는 연속된 숫자, 'd'는 분리필요

    idx_list = get_out_table(sheet_name)

    idx_cal = [idx_list[0]]
    tmp = idx_list[0]
    for val in idx_list:
        if tmp == val:
            continue
        elif tmp == val - 1 :
            idx_cal.append('c')
            tmp = val
        else :
            idx_cal.append('d')
            idx_cal.append(val) 
            tmp = val
    
    idx_cal = list(map(str, idx_cal))
    d_count = idx_cal.count('d')

    idx_done = ''.join(idx_cal).split('d')
    idx_list_fin = []

    
    fin_idx = return_dict(1)[sheet_name.name]
    
    for idx in range(d_count+1):
        val = idx_done[idx]
        if 'c' in val :
            c_count = val.count('c')
            val = val.replace('c','')
            idx_list_fin.append(val + 'C' +str(int(val) + c_count))
        else :
            idx_list_fin.append(val)

    fin_idx = fin_idx + '_'+'A'.join(idx_list_fin)

    return fin_idx


#############################################################################
def get_out_table(sheet_name,index_row_number=9):
    """
    xlwings.main.Sheet를 인수로 입력, 해당시트의 index행번호 default = 9 (int)
    """
    out_row_nums = sheet_name.range("C2").options(numbers=int).value
    col_count = sheet_name.range("XFD9").end('left').column
    idx_row_num = index_row_number
    col_names = sheet_name.range(sheet_name.range(int(idx_row_num),1),sheet_name.range(int(idx_row_num),col_count)).value
    
    for idx ,i in enumerate(col_names):
        
        if i == None:
            continue
        elif '_index' in i :
            col_num = idx
    
    df_so = pd.DataFrame()
    
    try :
        row_list = out_row_nums.replace(' ','').split(',')
    except:
        row_list = [out_row_nums]
    
    for row in row_list :

        # 연속된 행인 경우
        if '~' in str(row) :
            left_row =int(row.split('~')[0])
            right_row = int(row.split('~')[1])
            rng = sheet_name.range(sheet_name.range(left_row,1),sheet_name.range(right_row,col_count))
            df_so = pd.concat([df_so,pd.DataFrame(sheet_name.range(rng).options(numbers=int).value)]) 
        else :
            left_row =int(row)
            right_row = int(row)
            rng = sheet_name.range(sheet_name.range(left_row,1),sheet_name.range(right_row,col_count))
            df_so = pd.concat([df_so,pd.DataFrame(sheet_name.range(rng).options(numbers=int).value).T])

    
    return list(df_so[col_num])



# 
def get_out_info(sheet_name):
    

    info_list = sheet_name.range("C3:C7").value
    info_list[1]=info_list[1].date().isoformat()
    today= datetime.today().date().isoformat().replace('-','')
    tmp_idx = str(wb_cy.sheets['temp_db'].range("C500000").end('up').row -1 )
    out_idx = today + '_' + tmp_idx
    
    if sheet_name.range("C2").value == 'only_local':
        
        info_list.insert(0,'only_local')
    else :
        info_list.insert(0,get_idx(sheet_name))

    info_list.insert(0,out_idx)

    return info_list


## 시트보호 잠금 및 해제 매서드

def sht_protect(mode=True):
    """
    True 이면 시트보호모드, False 이면 시트보호해제
    """
    wb = xw.Book.caller()
    act_sht=wb_cy.selection.sheet
    status_col = act_sht.range("XFD4").end('left').column
    status_cel = act_sht.range(4,status_col)
    password = 'themath93'

    if mode == True:

        if status_cel.value != 'edit_mode' :
            wb_cy.save()
            act_sht.api.Unprotect(Password = password)

            # status창 변경
            status_cel.value = 'edit_mode'

        else : 
            clear_form()
            protect_sht(act_sht,password)

            # status창 변경



    elif mode == False:
        wb_cy.save()
        act_sht.api.Unprotect(Password='themath93')

        # status창 변경
        status_cel.value = 'edit_mode'

def protect_sht(act_sht,password):
    act_sht.api.Protect(Password=password, DrawingObjects=True, Contents=True, Scenarios=True,
            UserInterfaceOnly=True, AllowFormattingCells=True, AllowFormattingColumns=True,
            AllowFormattingRows=True, AllowInsertingColumns=True, AllowInsertingRows=True,
            AllowInsertingHyperlinks=True, AllowDeletingColumns=True, AllowDeletingRows=True,
            AllowSorting=True, AllowFiltering=True, AllowUsingPivotTables=True)
    



def get_empty_row(sheet=wb_cy.selection.sheet,col=1):
    """
    특정컬럼의 마지막 값의 행번호 구하기
    """
    sel_sht = sheet
    col_num = col
    if type(col) == int :
        row_start_nm = sel_sht.range(1048576,col_num).end('up').row + 1 
    elif type(col) == str :
        row_start_nm = sel_sht.range(col+str(1048576)).end('up').row + 1 
    return row_start_nm