import sys, os
import xlwings as xw
# sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))

wb_pid = xw.apps.keys()[0]

print(xw.apps[wb_pid].books[0].sheets[0].name)
print(wb_pid)
print(xw.serve)
@xw.sub
def main():
    """Writes the name of the Workbook into Range("A1") of Sheet 1"""
    wb = xw.Book.caller()
    wb.sheets[0].range('A1').value = wb.name

if __name__ == '__main__':
    xw.serve()


    