## Written for classficate the image data and rewrite Excel data
## Contact by kesm03122@gmail.com or Kakaotalk ID: Leecarry16

import os
from openpyxl import Workbook, load_workbook
import openpyxl
import pandas
import hashlib
import sys

dataframeXlsx = pandas.read_excel("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/이미지 분류 진행과정.xlsx", sheet_name="Sheet1")

print(dataframeXlsx)
print()

## From https://swlock.blogspot.com/2021/10/python-count-files.html
def get_files_count(folder_path):
	dirListing = os.listdir(folder_path)
	return len(dirListing)
	


print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/0105"))
print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/0610"))
print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/1115"))
print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/1620"))
print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/2125"))
print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/2630"))
print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/3135"))
print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/3640"))
print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/4145"))
print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/4650"))
print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/5155"))
print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/5660"))
print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/6165"))
print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/6670"))
print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/7175"))
print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/7680"))
print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/8185"))
print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/8690"))
print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/9195"))
print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/9600"))

rb = load_workbook("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/이미지 분류 진행과정.xlsx")

ws = rb['Sheet1']
wb = openpyxl.load_workbook("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/이미지 분류 진행과정.xlsx")

Sheet1 = wb.active 
Sheet1['B2'] = int(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/0105"))
Sheet1['B3'] = int(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/0610"))
Sheet1['B4'] = int(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/1115"))
Sheet1['B5'] = int(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/1620"))
Sheet1['B6'] = int(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/2125"))
Sheet1['B7'] = int(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/2630"))
Sheet1['B8'] = int(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/3135"))
Sheet1['B9'] = int(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/3640"))
Sheet1['B10'] = int(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/4145"))
Sheet1['B11'] = int(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/4650"))
Sheet1['B12'] = int(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/5155"))
Sheet1['B13'] = int(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/5660"))
Sheet1['B14'] = int(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/6165"))
Sheet1['B15'] = int(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/6670"))
Sheet1['B16'] = int(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/7175"))
Sheet1['B17'] = int(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/7680"))
Sheet1['B18'] = int(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/8185"))
Sheet1['B19'] = int(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/9195"))
Sheet1['B21'] = int(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/9600"))

wb.save("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/이미지 분류 진행과정.xlsx")



