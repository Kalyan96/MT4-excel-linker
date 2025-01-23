import xlwings as xw
from xlwings import Book

wbtest = xw.Book('myproject.xlsm')

ws1 = wbtest.sheets['sheet1']

ws1.range('A1').value = "hi there"
