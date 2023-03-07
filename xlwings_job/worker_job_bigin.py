## xl_wings 절대경로 추가
import sys, os
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))

import xlwings as xw

wb_caller = xw.Book.caller()

def begin_work():
    form_book_dir = os.path.join(os.path.expanduser('~'),'Desktop') + "\\fulfill\\xlwings_job\\cytiva.xlsm"
    wb_worker = xw.Book.caller()
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
    wb_worker.app.alert("FORM을 불러왔습니다. 작업 종료 시 저장은 필요 하지 않습니다.")