## xl_wings 절대경로 추가
import sys, os
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))

# os.path.dirname(os.path.abspath(os.path.dirname(__file__)))
# 위 명령어는 아래 의 경로를 말한다.
# c:\users\"username"\desktop\fulfill



import xlwings as xw

wb_caller = xw.Book.caller()

def test():
    wb_caller.app.alert(os.path.join(os.path.expanduser('~'),'Desktop') + "\\cytiva_worker.xlsm")
    wb_caller.app.alert(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))
    return