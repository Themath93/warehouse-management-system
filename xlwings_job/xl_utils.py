## xl_wings 절대경로 추가
import sys, os



sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))


from dicts_cy import return_dict
from oracle_connect import DataWarehouse


import xlwings as xw
import pandas as pd
import datetime as dt
import json
from barcode import Code128
from barcode.writer import ImageWriter

wb_cy = xw.Book("cytiva_worker.xlsm").set_mock_caller()
wb_cy = xw.Book.caller()

## 선택한 행, 시트 이름 딕셔너리로 반환
def row_nm_check(xw_book_name=wb_cy):
    """
    Return Dict, get activated sheet's name(str), get selected row's number(list)

    each_list는 연속된 row번호들을 전부 계산하여 리스트안에 전부 각각 위치하도록 한다.
    each_list의 모든 값은 int
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

    fin_list = []
    for rng_val in range_list:
        if ',' in rng_val:
            tmp_split = rng_val.split(',')
            cnts = int(tmp_split[1])-int(tmp_split[0])
            tmp_list = []
            for i in range(cnts+1):
                tmp_list.append(int(tmp_split[0])+i)
                
            fin_list = fin_list + tmp_list
        else :
            fin_list.append(int(rng_val))
            
    dict_row_num_sheet_name = {"sheet_name":xw_book_name.selection.sheet.name,'selection_row_nm':range_list,'each_list':fin_list}
    
    
    return dict_row_num_sheet_name


## 출고행 번호 리스트 -> 문자열로 변환
def get_row_list_to_string(seleted_row_list) :
    return ' '.join(seleted_row_list).replace(',','~').replace(' ', ', ')


## 출고 form 초기화 및 출고 대기상태로 변경
def clear_form(sel_sht = wb_cy.selection.sheet):
    
    
    protect_sht_pass = 'themath93'
    wb_cy.app.screen_updating = False
    #통합제어 시트일 경우
    if sel_sht.name == '통합제어' :

        main_clear('SVC')
    else :
        if sel_sht.name == 'IR_ORDER' :
            # DATE_TYPE
            str_dt = 'ALL,ARRIVAL_DATE,SHIP_DATE'
            rng_dt = sel_sht.range((2,"C"))
            rng_dt.api.Validation.Delete()
            rng_dt.api.Validation.Add(Type=3, Formula1=str_dt)
            # STATE
            str_state = "ALL," + ",".join(pd.DataFrame(DataWarehouse().execute("select * from inspection_code").fetchall())[0])
            rng_state = sel_sht.range((5,"C"))
            rng_state.api.Validation.Delete()
            rng_state.api.Validation.Add(Type=3, Formula1=str_state)
        try:
            __xl_clear_values(sel_sht)
        except :
            sel_sht.api.Unprotect(Password='themath93')
            protect_sht(sel_sht,protect_sht_pass)
            __xl_clear_values(sel_sht)
            if sel_sht.name == 'SHIPMENT_INFORMATION':
                # si_index컬럼 기준으로 오름차순 정렬 -> so_out시 excel 내용 업데이트에 반드시필요
                last_row = sel_sht.range("A1048576").end('up').row
                sel_sht.range((9,'A'),(last_row,'R')).api.Sort(Key1=sel_sht.range((9,'A')).api, Order1=1, Header=1, Orientation=1)
        
    wb_cy.app.screen_updating = True
    
def __xl_clear_values(sel_sht):
    ws_db = wb_cy.sheets['Temp_DB']
    tmp_last_row = get_empty_row(ws_db,"T")-1
    k = ws_db.range((2,"T"),(tmp_last_row,"T")).value
    # tmp_dict = dict(zip(k,v))
    # wb_cy.app.alert(k[0])
    for sht_name in k:
        # wb_cy.app.alert(sht_name)
        if sel_sht.name == sht_name :
            k_idx = k.index(sel_sht.name)
            # wb_cy.app.alert(str(k_idx))
            k_row = 2+k_idx
            ws_db.range((k_row,"U")).clear_contents()

    status_cel = sel_sht.range("H4")
    ws_db.range("U2:U3").clear_contents()
    sel_sht.range("C2:C7").clear_contents()
    sel_sht.range('V3:V4').clear_contents()
    sel_sht.range('P3:P7').clear_contents()
    status_cel.color = None
    status_cel.value = "waiting_for_out"
    sel_sht.range("C4").value = "=TODAY()+1"
    

def main_clear(type=None):
    """
    통합제어 시트 from클리어 매서드
    """
    ws_main =wb_cy.sheets['통합제어']

    if type != None:
        
        shapes = ws_main.shapes
        last_row = get_empty_row(ws_main,'J')
        for shp in shapes :
            if type in shp.name :
                shp.delete()

        ws_main.range("J12:S"+str(last_row)).clear_contents()
    else :
        pass



def get_idx(sheet_name):
    """
    xlwings.main.Sheet를 인수로 입력

    get_each_index_num 모듈의 반대로 반환함
    example : get_idx_str='si_13384A14171A14243C14244'  ==>
    {'out_sht_id':'si','idx_list':[13384, 14171, 14243, 14244]}

    [13384, 14171, 14243, 14244] -> '13384A14171A14243C14244'
    
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
def get_out_table(sheet_name=wb_cy.selection.sheet,index_row_number=9,direct_call=None):
    """
    xlwings.main.Sheet를 인수로 입력, 해당시트의 index행번호 default = 9 (int)
    list형태의 출고하는 시트의 index들을 반환한다.
    """
    if direct_call == True:
        out_row_nums = get_row_list_to_string(row_nm_check(wb_cy)['selection_row_nm'])
    else:
        out_row_nums = sheet_name.range("C2").options(numbers=int).value
    last_col = sheet_name.range("XFD9").end('left').column
    idx_row_num = index_row_number
    col_names = sheet_name.range(sheet_name.range(int(idx_row_num),1),sheet_name.range(int(idx_row_num),last_col)).value
    
    for idx ,i in enumerate(col_names):
        
        if i == None:
            continue
        elif '_INDEX' in i :
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
            rng = sheet_name.range(sheet_name.range(left_row,1),sheet_name.range(right_row,last_col))
            df_so = pd.concat([df_so,pd.DataFrame(sheet_name.range(rng).options(numbers=int).value)]) 
        else :
            left_row =int(row)
            right_row = int(row)
            rng = sheet_name.range(sheet_name.range(left_row,1),sheet_name.range(right_row,last_col))
            df_so = pd.concat([df_so,pd.DataFrame(sheet_name.range(rng).options(numbers=int).value).T])

    
    return list(df_so[col_num])



# 
def get_out_info(sheet_name):

    #2 배송방법 3 인수증방식
    info_list = sheet_name.range("C3:C7").options(numbers=int).value
    info_list[1]=str(info_list[1].date().isoformat())
    # 배송방법, 인수증방식은 DB에서 해당 내용으로 키값을 받아 DB에저장 -> byte사용이적어 용량에 유리
    info_list[2] = get_tb_idx('DELIVERY_METHOD',info_list[2])
    info_list[3] = get_tb_idx('POD_METHOD',info_list[3])

    #POD_DATE 컬럼필요
    info_list.append(None)
    timeline= create_db_timeline()
    info_list.append(timeline)
    # tmp_idx = str(wb_cy.sheets['temp_db'].range("C500000").end('up').row -1 )
    out_idx = None
    
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
    edit_mode 실행 매서드
    """
    # wb = xw.Book.caller()
    sel_sht=wb_cy.selection.sheet
    status_cel = sel_sht.range("H4")
    password = 'themath93'

    if mode == True:

        if status_cel.value != 'edit_mode' :
            wb_cy.save()
            sel_sht.api.Unprotect(Password = password)

            # status창 변경
            status_cel.value = 'edit_mode'

        else : 
            clear_form()
            protect_sht(sel_sht,password)
            # 필터 on 기능 edit_mode 해제후 필터링 할 때 다시 edit_mode로 돌아가 필터기능 켜기 방지
            last_col = sel_sht.range("XFD9").end('left').column
            filter_range = sel_sht.range((9,1),(9,last_col))
            filter_range.api.AutoFilter(Field=1)
           



    elif mode == False:
        wb_cy.save()
        sel_sht.api.Unprotect(Password='themath93')

        # status창 변경
        status_cel.value = 'edit_mode'

def protect_sht(sel_sht,password):
    sel_sht.api.Protect(Password=password, DrawingObjects=True, Contents=True, Scenarios=True,
            UserInterfaceOnly=True, AllowFormattingCells=True, AllowFormattingColumns=True,
            AllowFormattingRows=True, AllowInsertingColumns=True, AllowInsertingRows=True,
            AllowInsertingHyperlinks=True, AllowDeletingColumns=True, AllowDeletingRows=True,
            AllowSorting=True, AllowFiltering=True, AllowUsingPivotTables=True)
    



def get_empty_row(sheet=wb_cy.selection.sheet,col=1):
    """
    특정컬럼의 값이 있는 마지막 행 + 1을 반환
    """
    sel_sht = sheet
    col_num = col
    if type(col) == int :
        row_start_nm = sel_sht.range(1048576,col_num).end('up').row + 1 
    elif type(col) == str :
        row_start_nm = sel_sht.range(col+str(1048576)).end('up').row + 1 
    return row_start_nm


def get_current_time():
    """
    현재시간 년,월,일 시,분,초 반환
    """
    now = str(dt.datetime.now()).split('.')[0]

    return now



def save_barcode_loc(index=str):
    """
    받은 index(str)을 바코드 이미지로 만들어 저장
    """
    file_name = index+".jpeg"
    render_options = {
                    "module_width": 0.05,
                    "module_height": 9.5,
                    "write_text": True,
                    "module_width": 0.25,
                    "quiet_zone": 0.1,
                }

    barcode=Code128(index,writer=ImageWriter()).render(render_options)
    barcode.save(file_name)
    pic = '\\'+file_name
    pic = os.getcwd()+pic
    return pic


def get_tb_idx(tb_name=str, content=str):
    """
    DW의 table이름을, content에는 테이블의 content를 입력하면 tb상 key값을 반환한다.
    """
    cur = DataWarehouse()
    dic_dm = dict(cur.execute(f'select * from {tb_name}').fetchall())
    dic_dm = dict(zip(list(dic_dm.values()),list(dic_dm.keys())))
    return dic_dm[content]


def get_each_index_num(get_idx_str):
    """
    DW테이블 SO_OUT상의 si_index 및 is_local 컬럼값을 넣으면 해당 row의 고유 키값을 dict형태로 반환

    example : get_idx_str='si_13384A14171A14243C14244'  ==>
    {'out_sht_id':'si','idx_list':[13384, 14171, 14243, 14244]}
    """
    del_sht_id = get_idx_str.split('_')[0]
    get_idx_str = get_idx_str.split('_')[1]
    count_A = get_idx_str.count("A")
    count_C = get_idx_str.count("C")
    if count_A == 0 and count_C == 0:
        return get_idx_str
    procs_1 = get_idx_str.split("A")
    procs_1
    A_list = []
    C_list = []
    for val in procs_1:
        if val.count("C") > 0 :
            C_list.append(val)
        else :
            A_list.append(int(val))
    
    for c_val in C_list:      
        tmp_c = c_val.split('C')
        tmp_diff = int(tmp_c[1])-int(tmp_c[0])
        fin_C_list = []
        for i, val in enumerate(range(tmp_diff+1)):
            fin_C_list.append(int(tmp_c[0]) + i)
        A_list = A_list+fin_C_list
            
    return {'out_sht_id':del_sht_id,'idx_list':A_list}


    

def get_xl_rng_for_ship_date(xl_selection = wb_cy.selection,  ship_date_col_num=str):
    """
    sheet.range(ship_date_col_xl_rng_list)
    ship_date_col_num는 해당 시트의 ship_date컬럼의 알파벳 입력하면됨
    range를 위한 str를 반환 객체를 반환하는 것은 아님
    """

    count_dollar = xl_selection.api.Address.count("$")
    count_dollar
    # 셀한개만 클릭할경우
    if count_dollar == 2 :

        rng_str = xl_selection.api.Address
    else : 

        rng_str = xl_selection.api.SpecialCells(12).Address

    rng_str_1 = rng_str.replace('$','')
    comma_spt = rng_str_1.split(',')

    for idx, rng in enumerate(comma_spt):
        if ':' in rng:
            colon_idx = rng.index(':')
            alpha_0 = rng[0]
            alpha_1 = rng[colon_idx+1]
            comma_spt[idx] = comma_spt[idx].replace(alpha_0,ship_date_col_num)
            comma_spt[idx] = comma_spt[idx].replace(alpha_1,ship_date_col_num)
        else:
            alpha_0 = rng[0]
            comma_spt[idx] = comma_spt[idx].replace(alpha_0,ship_date_col_num)
    rng_fin = ','.join(comma_spt)
    return rng_fin



##### change_cell 모듈 ######################### cell한개의 내용 변경########
def select_cell():
    sel_cells = wb_cy.selection
    sel_sht = wb_cy.selection.sheet
    # 선택한 셀의 row번호가 10미만이면 종료 ==> table값은 row가 10부터 시작이기 때문
    if wb_cy.selection.row < 10 :
        wb_cy.app.alert("선택한 셀은 바꿀 수 없습니다. 매서드를 종료합니다.","Change Cell WARNING")
        return None
    address_cell = sel_sht.range("E3")
    from_cell = sel_sht.range("E4")
    # 선택한셀의 value가 list type 이면 두 개이상의 셀을 선택 했다는 것 ==> 종료
    if type(sel_cells.value) is list :
        wb_cy.app.alert("하나의 셀만 선택 후 진행해주세요. 두 개 이상은 불가합니다.","Change Cell WARNING")
        return None
    
    address_cell.value = str(sel_cells.address)
    from_cell.value = sel_cells.value

def change_cell():
    sel_sht = wb_cy.selection.sheet
    address_cell = sel_sht.range("E3")
    change_cell_list = sel_sht.range("E3:E4")
    tb_name = sel_sht.range("D5").value
    idx_col_name = sel_sht.range("A9").value
    cur = DataWarehouse()
    
    # 셀주소가 빈값이면 중지한다.
    if address_cell.value == None:
        wb_cy.app.alert("바꿀 셀이 없습니다 매서드를 종료합니다","Change Cell WARNING")
        return None
    
    xl_from_cell = sel_sht.range(address_cell.value)
    to_cell = wb_cy.app.api.InputBox("바꿀 내용을 입력 해주세요", "Change Cell Input", Type=2)
    # to_cell == False면 입력을 취소 했다는 뜻이므로 바꿀 뜻이 없는 것으로 간주하고 주소와 바뀔 값들의 form을 지운다.
    if to_cell == False :
        wb_cy.app.alert("취소를 선택하셨습니다. 셀 변경을 취소합니다.","Change Cell WARNING")
        change_cell_list.clear_contents()
        return None
    
    # DB UPDATE 진행
    row_num = sel_sht.range(address_cell.value).row
    col_mum = sel_sht.range(address_cell.value).column
    idx_num = sel_sht.range(row_num,1).options(numbers=int).value
    col_name = sel_sht.range(9,col_mum).options(numbers=int).value
    query = f"UPDATE {tb_name} SET {col_name} = '{to_cell}' WHERE {idx_col_name} = {idx_num}"
    cur.execute(query)
    cur.execute("commit")
    
    # DB UPDATE 완료 후 xl_cell 내용 변경
    xl_from_cell.value = to_cell
    
    # 모든게 완료 되면 change_cell_list 내용 모두삭제
    change_cell_list.clear_contents()

    # 변경 성공 메시지
    wb_cy.app.alert("셀 내용 변경이 완료되었습니다.","Change Cell Done")
##### change_cell 모듈 ######################### cell한개의 내용 변경########


# data_insert 모드 및 data_input() 매서드로 실행
def data_insert():
    import sys, os
    sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))))
    from datajob.xlwings_dj.shipment_information import ShipmentInformation
    from datajob.xlwings_dj.local_list import LocalList
    from datajob.xlwings_dj.ir_order import IROrder
    """
    입력받은 데이터를 맞는 db에 입력한다.
    """
    tb_dict = {'LOCAL_LIST':LocalList, 'SHIPMENT_INFORMATION':ShipmentInformation,'IR_ORDER':IROrder}
    tmp_idx = [*range(1,1000)]
    sel_sht = wb_cy.selection.sheet
    status = sel_sht.range("H4")
    db_table_name = sel_sht.range("D5").value
    cur= DataWarehouse()
    if status.value == 'edit_mode':  # edit_mode에서만 data_input_mode() 사용 가능
        last_row = sel_sht.range("A1048576").end('up').row
        if last_row < 10:
            last_row = 10
        last_col = sel_sht.range("XFD9").end("left").column
        xl_table = sel_sht.range((10,1),(last_row,last_col))
        xl_table.clear_contents()
        status.value = 'data_insert_mode'

        sel_sht.range((10,1),(1008,last_col)).api.Borders.LineStyle = 1 

        sel_sht.range("A10").select() # 데이터 진입모드시 입력할 첫째 행으로 포인터 이동 필수..
        sel_sht.range("A10").options(transpose=True).value = tmp_idx # index컬럼에 값이 없다면 매서드 진행 불가로 index에는 자동으로 값을 넣어주자 최대 999개
        wb_cy.app.alert("data_insert_mode모드로 진입합니다. 데이터 입력 후 버튼을 다시한번 눌러주세요.","DATA Input Mode")
        return
    elif status.value == 'data_insert_mode':  # data_insert_mode 상태면 다시 data_input기능을 마칠지 물어본다.
        input_comfirm = wb_cy.app.alert(" B 컬럼은 빈칸이 없어야 정상작동 합니다. 입력한 DATA를 Confirm하시겠습니까?",
                    "Input Confirm",buttons="yes_no_cancel")
        if input_comfirm == 'yes': # data input 시작
            # 입력한 값이 있는지는 확인해봐야함
            last_row = sel_sht.range("B1048576").end('up').row
            last_col = sel_sht.range("XFD9").end("left").column
            if last_row > 9: # 10번 행에 값이 있다는 뜻.. data_input실행가능

                # 참고할 컬럼행에 값이 띄엄띄엄 있는것은 아닌지 확인하기
                none_count = sel_sht.range((9,"B"),(last_row,"B")).value.count(None)
                if none_count > 0 :
                    wb_cy.app.alert(f"'B' 컬럼은 반드시 값이 있어야합니다. \n 비어있는 행이 {none_count} 개 확인 됩니다. \n 수정 후 다시 시도해주세요!","Input WARNING")
                    return     

                tb_dict[db_table_name].data_input()
                # ShipmentInformation.data_input() # data db업로드 완료

                # db에서 해당 테이블 모든 데이터 불러와서 xl_si_sht로 데이터 전송 및 서식 맞추기
                __bring_tb_from_db_and_formatting_xltb(sel_sht, cur, last_col) 


                # edit_mode로 돌아가기
                sht_protect()
                wb_cy.app.alert("DATA INPUT이 완료 되었습니다! (DB_반영완료).","Input COMPLETE")
            else :
                wb_cy.app.alert("Input할 값이 없습니다. \n 'B'컬럼에는 반드시 값이 있어야합니다.","Input WARNING")
                return


        else : #False #취소이므로 매서드 중단
            wb_cy.app.alert("DATA Input을 중단합니다.","Input Cancel")
            return None
        
    else : 
        wb_cy.app.alert("해당 기능은 'edit_mode'상태에서만 사용가능합니다.","WARNING")
        return None

def __bring_tb_from_db_and_formatting_xltb(sel_sht, cur, last_col):
    table_name = sel_sht.range("D5").value
    sht_idx_col_name = sel_sht.range("A9").value
    sel_sht.api.AutoFilterMode = False # Filter가 걸려져있으면 데이터복사가 원활하게 되지 않는다. 필터해제 코드
    df_si = pd.DataFrame(cur.execute(f'select * from {table_name}').fetchall())
    col_names = sel_sht.range((9,1),(9,last_col)).value
    df_si.columns=col_names
    df_si = df_si.sort_values(sht_idx_col_name)
    df_si.set_index(sht_idx_col_name,inplace=True)
    df_si = df_si.replace('None','')
    # if table_name == 'SHIPMENT_INFORMATION':
    #     df_si['TIMELINE'] = df_si['TIMELINE'].map(lambda e: json.loads(e)['data'][-1]['c'])
    df_si['TIMELINE'] = df_si['TIMELINE'].map(lambda e: json.loads(e)['data'][-1]['c'])
    sel_sht.range("A9").value = df_si
    last_row = sel_sht.range("B1048576").end('up').row
    last_col = sel_sht.range("XFD9").end("left").column
    xl_content = sel_sht.range((10,1),(last_row,last_col))
    xl_content.api.Borders.LineStyle = 1 
    xl_content.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft


# db 테이블 데이터를 xl로 시트로 보여준다.
def bring_data_from_db(in_method=False):
    # sheet에 filter가 걸려져 있을경우 낭패를 본다. 반드시 필터해제후 진행
    """
    해당 시트에 DB전체 table 데이터 불러오기
    """
    wb_cy = xw.Book.caller()
    sel_sht = wb_cy.selection.sheet
    status = sel_sht.range("H4").value
    if in_method == True:
        status = 'edit_mode'
    if status == 'edit_mode': # edit_mode에서만 지원

        sel_sht.api.AutoFilterMode = False # Filter가 걸려져있으면 데이터복사가 원활하게 되지 않는다. 필터해제 코드
        cur=DataWarehouse()
        db_table_name = sel_sht.range("D5").value
        idx_col_name = sel_sht.range("A9").value
        
        last_col = sel_sht.range("XFD9").end("left").column

        df = pd.DataFrame(cur.execute(f'select * from {db_table_name}').fetchall())
        col_names = sel_sht.range((9,1),(9,last_col)).value
        df.columns=col_names
        if col_names.count('TIMELINE') > 0 :
            df['TIMELINE'] = df['TIMELINE'].map(lambda e: json.loads(e)['data'][-1]['c'])
        df = df.sort_values(idx_col_name)
        df.set_index(idx_col_name,inplace=True)
        df = df.replace('None','')
        sel_sht.range("A9").value = df
        last_row = sel_sht.range("B1048576").end('up').row
        xl_content = sel_sht.range((10,1),(last_row,last_col))
        xl_content.font.size = 11
        xl_content.api.Borders.LineStyle = 1 
        xl_content.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft

    else:
        wb_cy.app.alert("데이터 가져오기는 edit_mode에서만 사용 가능합니다.","Bring Data WARNING")


def create_db_timeline(up_time_content=None):
    """
    DB Table의 col중 UP_TIME 같이 업데이트가 되는 테이블들은 해당 매서드로 업데이트 
    UPDATE 회수는 68번 제한 oracle varchar2()는 4000이 한계
    json.loads(obj)로 이용가능한 str return
    """
    import datetime as dt
    import json

    # create 시 처음 json데이터 생번
    data=[]
    cols=['a','b','c']
    if up_time_content is None:
        rows=[]
        now = str(dt.datetime.now()).split('.')[0]
        rows.append('create')
        rows.append(str(0))
        rows.append(now)
        tmp = dict(zip(cols,rows))
        data.append(tmp)
        res = {
            'meta':{
                'desc':'data_update_timeline',

                'cols':{
                    'a':'db_method',
                    'b':'count',
                    'c':'timeline',

                }
            },
            'data':data
        }
        return json.dumps(res,ensure_ascii=False)
    # 데이터 생성이후 update시마다 계속 업데이트
    else : 
        load_json = json.loads(up_time_content)['data']
        update_count=len(load_json)
        if len(load_json) == 68:
            return up_time_content
        rows=[]
        now = str(dt.datetime.now()).split('.')[0]
        rows.append('update')
        rows.append(str(update_count))
        rows.append(now)
        tmp = dict(zip(cols,rows))
        load_json.append(tmp)
        res = {
            'meta':{
                'desc':'data_update_timeline',

                'cols':{
                    'a':'db_method',
                    'b':'count',
                    'c':'timeline',

                }
            },
            'data':load_json
        }
        
        return json.dumps(res,ensure_ascii=False)
    

def recall_col_names():
    """
    작업 중 실수로 컬럼명을 바꿀 경우 원활한 작업이 불가할 수 있어 해당 기능을 사용하여 컬럼이름을 복구한다.
    """
    sel_sht = wb_cy.selection.sheet
    table_name_cel = sel_sht.range("D5").value
    col_names_from_db = list(pd.DataFrame(DataWarehouse().execute(f"select column_name from user_tab_columns where table_name = upper('{table_name_cel}')").fetchall())[0])
    sel_sht.range("A9").value = col_names_from_db


def move_sht(sht_name):
    if '_sheetBtn' not in sht_name:
        return
    sht_name = sht_name.replace("_sheetBtn","")
    sht_names = wb_cy.sheet_names
    if sht_name not in sht_names:
        return
    selected_sht = wb_cy.sheets[sht_name]
    selected_sht.api.Activate()

def create_manual_print_form():
    wb_cy = xw.Book("cytiva_worker.xlsm").set_mock_caller()
    wb_cy = xw.Book.caller()
    answer = wb_cy.app.alert("print_form을 여시겠습니까?","CONFIRM",buttons="yes_no_cancel")
    if answer != "yes":
        wb_cy.app.alert("종료합니다.","Quit")
        return
    # 프린트폼 불러오기
    book_name = 'print_form.xlsx'
    form_book_dir = os.path.join(os.path.expanduser('~'),'Desktop') + "\\fulfill\\xlwings_job\\"+book_name

    wb_form = xw.Book(form_book_dir)
    # 새로운 워크북 생성해서 시트 붙여넣기
    new_wb = xw.Book()
    
    form_sheets = wb_form.sheets

    # worker 엑셀 시트에 from 시트값 주기
    for form_sht in reversed(form_sheets):
        # Copy to second Book requires to use before or after
        form_sht.copy(after=new_wb.sheets["Sheet1"])
    wb_form = xw.Book(form_book_dir)
    new_wb.sheets["Sheet1"].delete()
    wb_form.close()
    new_wb.app.alert("매뉴얼로 작업이 필요할 시 사용 해주세요.","INFO")

def create_manual_tool():
    wb_cy = xw.Book("cytiva_worker.xlsm").set_mock_caller()
    wb_cy = xw.Book.caller()
    answer = wb_cy.app.alert("TOOL LIST를 여시겠습니까?","CONFIRM",buttons="yes_no_cancel")
    if answer != "yes":
        wb_cy.app.alert("종료합니다.","Quit")
        return
    # 프린트폼 불러오기
    book_name = 'svc_tool.xlsm'
    form_book_dir = os.path.join(os.path.expanduser('~'),'Desktop') + "\\fulfill\\xlwings_job\\"+book_name

    wb_form = xw.Book(form_book_dir)
    # 새로운 워크북 생성해서 시트 붙여넣기
    new_wb = xw.Book()
    
    form_sheets = wb_form.sheets

    # worker 엑셀 시트에 from 시트값 주기
    for form_sht in reversed(form_sheets):
        # Copy to second Book requires to use before or after
        form_sht.copy(after=new_wb.sheets["Sheet1"])
    wb_form = xw.Book(form_book_dir)
    new_wb.sheets["Sheet1"].delete()
    wb_form.close()
    new_wb.app.alert("매뉴얼로 작업이 필요할 시 사용 해주세요.","INFO")


def call_system_stock():
    sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))))
    from datajob.xlwings_dj.system_stock import SystemStock
    SystemStock.put_data()

def open_user_guide():
    ppt_name = "cytiva_user_guide.pptx"
    guide_dir = os.path.join(os.path.expanduser('~'),'Desktop') + "\\fulfill\\xlwings_job\\"+ppt_name
    os.startfile(guide_dir)
