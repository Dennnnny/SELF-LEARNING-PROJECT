# -*- coding: utf-8 -*-
"""
Created on Fri Jul 20 22:44:41 2018

@author: Administrator
"""
import os
import time
#import requests
from openpyxl import load_workbook
# encoding: utf-8
# 使用pandas處理table資料
import pandas

date=time.strftime("%M%S%H%d%m%y")
date_for_file=time.strftime("%Y%m%d%H")


path=r"匯率\\"+time.strftime("%Y%m")+"\\"

if not os.path.isdir(r"匯率\\"):
	os.mkdir(r"匯率\\")

if not os.path.isdir(path):
	os.mkdir(path)
    
#file_name=("FX_"+time.strftime("%Y")+time.strftime("%m")+time.strftime("%d")+".xlsx")

#台銀匯率網址
dfs = pandas.read_html("http://rate.bot.com.tw/xrt?Lang=zh-TW")
#取dsf的list 資料
currency = dfs[0]
#只取前五欄
currency = currency.ix[:,0:5]
#重新命名欄位名稱 u-utf
currency.columns = [u'幣別',u'現金匯率-本行買入',u'現金匯率-本行賣出',u'現金匯率-本行買入',u'現金匯率-本行賣出']
#幣別值有重複字 利用正規式取英文代號
currency[u'幣別'] = currency[u'幣別'].str.extract('\((\w+)\)')
#將結果輸出到excel
currency.to_excel("網頁原檔.xlsx")

wb = load_workbook("網頁原檔.xlsx")
sheet = wb.active

sheet['B22'].value= "U/N"
sheet['B23'].value= "H/N"
sheet['B24'].value= "B/N"
sheet['B25'].value= "A/N"
sheet['B26'].value= "S/N"
sheet['B27'].value= "J/N"
sheet['B28'].value= "E/N"
sheet['B29'].value= "F/N"

sheet['C22'].value = float(sheet['E2'].value)/2 + float(sheet['F2'].value)/2
sheet['C23'].value = float(sheet['E3'].value)/2 + float(sheet['F3'].value)/2
sheet['C24'].value = float(sheet['E4'].value)/2 + float(sheet['F4'].value)/2
sheet['C25'].value = float(sheet['E5'].value)/2 + float(sheet['F5'].value)/2
sheet['C26'].value = float(sheet['E7'].value)/2 + float(sheet['F7'].value)/2
sheet['C27'].value = float(sheet['E9'].value)/2 + float(sheet['F9'].value)/2
sheet['C28'].value = float(sheet['E16'].value)/2 + float(sheet['F16'].value)/2
sheet['C29'].value = float(sheet['E8'].value)/2 + float(sheet['F8'].value)/2

sheet['D22'].value= "N/U"
sheet['D23'].value= "H/U"
sheet['D24'].value= "B/U"
sheet['D25'].value= "A/U"
sheet['D26'].value= "S/U"
sheet['D27'].value= "J/U"
sheet['D28'].value= "E/U"
sheet['D29'].value= "F/U"
sheet['D30'].value= "N/J"
sheet['D31'].value= "U/H"

sheet['E22'].value= round(int(1)/float(sheet['C22'].value),6)
sheet['E23'].value= round(float(sheet['C23'].value)/float(sheet['C22'].value),6)
sheet['E24'].value= round(float(sheet['C24'].value)/float(sheet['C22'].value),6)
sheet['E25'].value= round(float(sheet['C25'].value)/float(sheet['C22'].value),6)
sheet['E26'].value= round(float(sheet['C26'].value)/float(sheet['C22'].value),6)
sheet['E27'].value= round(float(sheet['C27'].value)/float(sheet['C22'].value),6)
sheet['E28'].value= round(float(sheet['C28'].value)/float(sheet['C22'].value),6)
sheet['E29'].value= round(float(sheet['C29'].value)/float(sheet['C22'].value),6)
sheet['E30'].value= round(int(1)/float(sheet['C27'].value),6)
sheet['E31'].value= round(float(sheet['C22'].value)/float(sheet['C23'].value),6)

wb.save(path+"FX_"+time.strftime("%Y")+time.strftime("%m")+time.strftime("%d")+".xlsx")

