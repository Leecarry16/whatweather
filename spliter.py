## Written for split files to test_set and train_set

import random
import os
import math
import shutil

## Class
pm_dir = ["0105", "0610", "1115", "1620", "2125", "2630", "3135", "3640", "4145", "4650", "5155", "5660", "6165", "6670", "7175", "7680", "8185", "8690", "9195", "9600", "error"]

## Original Directory
pm_path = "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/PM10 Class/{}"

## TestSet and TrainSet Directory
test_path = "C:/Users/user/Desktop/School/AI_Build_Project_test/Term Projecct/Program/Code/dataset/test/{}"
train_path = "C:/Users/user/Desktop\School/AI_Build_Project_test/Term Projecct/Program/Code/dataset/train/{}"

## Count Files in Folder
def get_files_count(folder_path):
	dirListing = os.listdir(folder_path)
	return len(dirListing)

for i in range(0, 19):
    pm_path_dir = pm_path.format(pm_dir[i])                         ## Directory of Orginal Class
    file_amount = get_files_count(pm_path.format(pm_dir[i]))        ## Files Amount of Original CLass
    file_list = os.listdir(pm_path_dir)                             ## File's name of Original Class
    test_len = 30/100                                               ## Test 30% / Train 70%
    test_amount = math.trunc(file_amount * test_len)
    train_amount = file_amount - test_amount

    if get_files_count(pm_path_dir) <3:
        print(pm_dir[i]," 파일 부족")
    else:
        try:
            for test in range(0, test_amount):                      ## Move Files to Test Folder
                shutil.move(pm_path_dir + "/" + file_list[random.randint(0, file_amount-test)], test_path.format(pm_dir[i]))

            for train in range(0, train_amount):                    ## Move Files to Train Folder
                shutil.move(pm_path_dir + "/" + file_list[train], train_path.format(pm_dir[i]))
                
        except shutil.Error:
            try:
                shutil.move(pm_path_dir, pm_path.format("error"))
            except shutil.Error:
                os.remove(pm_path_dir + "/" + file_list[train])
