## Written for classficate the image data and rewrite Excel data
## Contact by kesm03122@gmail.com or Kakaotalk ID: Leecarry16

from importlib.resources import path
import os
from openpyxl import Workbook, load_workbook
import openpyxl
import pandas as pd
import hashlib
import sys
import shutil
import re

dataframeXlsx = pd.read_excel("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/이미지 분류 진행과정.xlsx", sheet_name="Sheet1")

'''print(dataframeXlsx)
print()'''

## From https://swlock.blogspot.com/2021/10/python-count-files.html
def get_files_count(folder_path):
	dirListing = os.listdir(folder_path)
	return len(dirListing)
	
#shutil.move("C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/rename/20220510155254,.JPG", "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/1620")

'''print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/0105"))
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
print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/9600"))'''
#print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/nan"))


im_path = "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/rename"
file_list = os.listdir(im_path)

pm_exc = pd.read_excel("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10_20220522.xlsx")

print(type(pm_exc.iat[6,2]))
print(type(file_list))
print(float('nan'))

num = 0
test_list = []
for file in file_list:
	name = file.split(',')[0]
	name_y = name[0:4]
	name_m = name[4:6]
	name_d = name[6:8]
	name_h = name[8:10]
	name_ymd = name_y + "." + name_m + '.' + name_d + '.'
	for row_num in range(2, 33):
		if name_ymd in str(pm_exc.iat[row_num, 0]):
			pm_date = pm_exc.iat[row_num, 1]
			for col_num in range(1, 25):
				if name_h in re.sub(r'[^0-9]', '', str(pm_exc.iat[1, col_num])):
					pm_hour = pm_exc.iat[2, col_num]
					test_list.append(pm_exc.iat[row_num, col_num])
					
					if 1 <= pm_exc.iat[row_num, col_num] <= 5:
						shutil.move("C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/rename/{}".format(file_list[num]), "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/0105")
					elif 6 <= pm_exc.iat[row_num, col_num] <= 10:
						shutil.move("C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/rename/{}".format(file_list[num]), "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/0610")
					elif 11 <= pm_exc.iat[row_num, col_num] <= 15:
						shutil.move("C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/rename/{}".format(file_list[num]), "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/1115")
					elif 16 <= pm_exc.iat[row_num, col_num] <= 20:
						shutil.move("C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/rename/{}".format(file_list[num]), "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/1620")
					elif 21 <= pm_exc.iat[row_num, col_num] <= 25:
						shutil.move("C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/rename/{}".format(file_list[num]), "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/2125")
					elif 26 <= pm_exc.iat[row_num, col_num] <= 30:
						shutil.move("C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/rename/{}".format(file_list[num]), "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/2630")
					elif 31 <= pm_exc.iat[row_num, col_num] <= 35:
						shutil.move("C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/rename/{}".format(file_list[num]), "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/3135")
					elif 36 <= pm_exc.iat[row_num, col_num] <= 40:
						shutil.move("C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/rename/{}".format(file_list[num]), "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/3640")
					elif 41 <= pm_exc.iat[row_num, col_num] <= 45:
						shutil.move("C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/rename/{}".format(file_list[num]), "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/4145")
					elif 46 <= pm_exc.iat[row_num, col_num] <= 50:
						shutil.move("C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/rename/{}".format(file_list[num]), "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/4650")
					elif 51 <= pm_exc.iat[row_num, col_num] <= 55:
						shutil.move("C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/rename/{}".format(file_list[num]), "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/5155")
					elif 56 <= pm_exc.iat[row_num, col_num] <= 60:
						shutil.move("C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/rename/{}".format(file_list[num]), "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/5660")
					elif 61 <= pm_exc.iat[row_num, col_num] <= 65:
						shutil.move("C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/rename/{}".format(file_list[num]), "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/6165")
					elif 66 <= pm_exc.iat[row_num, col_num] <= 70:
						shutil.move("C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/rename/{}".format(file_list[num]), "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/6670")
					elif 71 <= pm_exc.iat[row_num, col_num] <= 75:
						shutil.move("C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/rename/{}".format(file_list[num]), "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/7175")
					elif 76 <= pm_exc.iat[row_num, col_num] <= 80:
						shutil.move("C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/rename/{}".format(file_list[num]), "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/7680")
					elif 81 <= pm_exc.iat[row_num, col_num] <= 85:
						shutil.move("C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/rename/{}".format(file_list[num]), "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/8185")
					elif 86 <= pm_exc.iat[row_num, col_num] <= 90:
						shutil.move("C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/rename/{}".format(file_list[num]), "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/8690")
					elif 91 <= pm_exc.iat[row_num, col_num] <= 95:
						shutil.move("C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/rename/{}".format(file_list[num]), "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/9195")
					elif 96 <= pm_exc.iat[row_num, col_num] <= 100:
						shutil.move("C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/rename/{}".format(file_list[num]), "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/9600")
					else:
						shutil.move("C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/rename/{}".format(file_list[num]), "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/nan")
						
					num = num + 1

print(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/nan"))
print(test_list)
print(len(test_list))



#print(pm_exc)


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
Sheet1['B22'] = int(get_files_count("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/nan"))

wb.save("C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/이미지 분류 진행과정.xlsx")



