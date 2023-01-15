import numpy as np
import xlwings as xw
from pathlib import Path
import os

def test():
    wb = xw.Book.caller()
    wb.sheets[0].range('A1').value = str(Path.cwd())

