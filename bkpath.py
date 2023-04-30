import os
import datetime
import shutil
import pandas as pd
import pyminizip
import xlwings as xw
Pw=os.environ.get('ZIP_PW').replace('dummy','')
today=datetime.date.today()
num=1
print('backup start')
TargetFolder=r'C:\\Users\\user\\PATH'
ZipFile=TargetFolder+".zip"
DstFile=r'D:\\01c_path\PATH_bak{:%y%m%d}.zip'.format(today,num)
wb=xw.Book(r"C:\Users\user\git\excel_vba\test02.xlsm")
macro=wb.macro('Zip_Pas')
macro(TargetFolder, Pw)
print('compress (password zip) finished')
shutil.move(ZipFile, DstFile)
print(os.listdir('D:\\01c_path'))   
print('compress (password zip) finished')
