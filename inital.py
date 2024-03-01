import xlwt
from xlwt import Workbook
from openpyxl import load_workbook
from openpyxl import Workbook
import os.path
from os import path


wb = Workbook()

#Checks if doc1.xlsx exists, creates it if it doesn't
if(not path.exists('doc1.xlsx')):
    wb.save('doc1.xlsx')


sheet = wb.active
sheet['A1'] = 'zdr'
sheet['B1'] = '3'

wb.save('doc1.xlsx')


