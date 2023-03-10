## xl_wings 절대경로 추가
import sys, os


sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))

import xlwings as xw
import pandas as pd
from xlwings_job.oracle_connect import DataWarehouse

wb_caller = xw.Book("cytiva_worker.xlsm").set_mock_caller()
wb_worker = xw.Book.caller()

## 수동으로 바꿀 것 새로운 버전 배포시 반드시 업데이트!
current_ver= float(1.133)

def version_check():
    os.chdir(os.path.join(os.path.expanduser('~'),'Desktop') + "\\fulfill\\")
    my_version = float(os.popen('git log -1 --pretty=%B').read().replace("\n\n","").split(" ")[-1])
    
    need_updated = my_version < current_ver
    if need_updated == True:
        answer = wb_worker.app.alert(f"New version has been released. Do you want to progress the update? \
                            \n 새로운 버전이 릴리즈 되었습니다. 업데이트를 진행 하시겠습니까? \
                            \n Current Version : {str(my_version)} \
                            \n New Version : {str(current_ver)}","UPDATE",buttons='yes_no_cancel')
        if answer != "yes":
            return
        else:
            wb_worker.app.alert("업데이트를 진행합니다.","UPDATE")
            os.system('git pull origin warehouse')
            # os.system('git fetch --all')
            # os.system('git reset --hard origin/warehouse')
            wb_worker.app.alert("업데이트가 완료 되어있습니다. \
                                \n 자세한 내용은 PATCH NOTE를 확인해주세요.","DONE")
            return




def begin_work():
    # wb_worker = xw.Book.caller()
    answer = wb_worker.app.alert("FULFILLMENT FORM 을  가져오시겠습니까? \n\n yes를 선택하신 경우 기존에 DB로 전송되지 않은 시트 값들은 \n\n 전부 사라집니다.","CONFIRM",buttons="yes_no_cancel")
    if answer != "yes":
        wb_worker.app.alert("종료합니다.","QUIT")
        return
    while True:
        worker_name = wb_worker.app.api.InputBox("귀하의 성함을 입력해주세요. Automail시 서명으로 사용됩니다.","NAME INPUT",Type=2)
        confirm_name = wb_worker.app.alert(f"귀하의 성함이 '{worker_name}'이 맞습니까? \n 'no' 클릭시 다시 입력가능합니다.","CONFIRM",buttons="yes_no_cancel")
        if confirm_name == "yes":
            break
        elif confirm_name == "no":
            pass
        else : 
            wb_worker.app.alert("종료합니다.","QUIT")
            return

    form_book_dir = os.path.join(os.path.expanduser('~'),'Desktop') + "\\fulfill\\xlwings_job\\cytiva.xlsm"
    
    wb_worker.app.screen_updating = False
    wr_sheets = wb_worker.sheets
    if len(wr_sheets) > 1:
        for sht in wr_sheets:
            if sht.name != "BEGIN":
                sht.delete()
    
    # 폼북열기
    wb_form = xw.Book(form_book_dir)
    form_sheets = wb_form.sheets
   
    # worker 엑셀 시트에 from 시트값 주기
    for form_sht in reversed(form_sheets):
        # Copy to second Book requires to use before or after
        form_sht.copy(after=wb_worker.sheets["BEGIN"])

    wb_form = xw.Book(form_book_dir)
    wb_form.close()
    wb_worker.app.screen_updating = False

    rp_word = 'fulfill\\xlwings_job\\cytiva.xlsm!'
    rp_word_2 = os.path.join(os.path.expanduser('~'),'Desktop') + "\\xlwings_job\\cytiva.xlsm!"
    rp_word_3 = "[1]!"

    for sht in wr_sheets:
        shapes = sht.shapes
        for idx,shp in enumerate(shapes):
            each_macro_name = shp.api.OnAction
            if rp_word in each_macro_name:
                print("첫번째")
                shapes[idx].api.OnAction = each_macro_name.replace(rp_word,"")
            elif rp_word_2 in each_macro_name:
                print("두번째")
                shapes[idx].api.OnAction = each_macro_name.replace(rp_word_2,"")
            elif rp_word_3 in each_macro_name:
                print("세번째")
                shapes[idx].api.OnAction = each_macro_name.replace(rp_word_3,"")

    
    ## validationlist 만들기
    cur = DataWarehouse()
    # validation_str
    req_type_str = 'ALL,SVC,PO,DEMO,RETURN'
    status_str = 'ALL,requested,pick/pack,dispatched,completed'
    del_met_str = ','.join(list(pd.DataFrame(cur.execute("SELECT * FROM DELIVERY_METHOD").fetchall())[1]))
    pod_met_str = ','.join(list(pd.DataFrame(cur.execute("SELECT * FROM POD_METHOD").fetchall())[1]))
    # SHEET OBJ
    sht_main = wb_worker.sheets['MAIN']
    sht_si = wb_worker.sheets['SHIPMENT_INFORMATION']
    sht_lc = wb_worker.sheets['LOCAL_LIST']

    # MAIN
    req_type_main = sht_main.range("M9")
    status_main = sht_main.range("O7")
    # SHIPMENT_INFORMATION
    del_met_si = sht_si.range("C5")
    pod_met_si = sht_si.range("C6")
    del_met_br_si = sht_si.range("P6")
    pod_met_si_br_si = sht_si.range("P7")
    # LOCAL_LIST
    del_met_lc = sht_lc.range("C5")
    pod_met_lc = sht_lc.range("C6")

    crate_validation(req_type_main,req_type_str)
    crate_validation(status_main,status_str)

    del_met_rngs = [del_met_si, del_met_br_si, del_met_lc]
    pod_met_rngs = [pod_met_si, pod_met_si_br_si, pod_met_lc]

    for rng in del_met_rngs:
        crate_validation(rng,del_met_str)
    for rng in pod_met_rngs:
        crate_validation(rng,pod_met_str)


    sht_main.range("D2").value = worker_name
    wb_worker.app.alert(f"Hello, {worker_name}",'WELCOME!!')

    

def crate_validation(req_type_rng=xw.Range,formula=str):
    req_type_rng.api.Validation.Delete()
    req_type_rng.api.Validation.Add(Type=3, Formula1=formula)
