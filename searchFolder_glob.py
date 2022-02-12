import pathlib
import os
import time
'''
# os.path.getatime(file) 輸出檔案訪問時間
# os.path.getctime(file) 輸出檔案的建立時間
# os.path.getmtime(file) 輸出檔案最近修改時間
'''

dirPath = r'D:\\'  # searching path

key_word = '**\*CPCProj*.pdf'  # searching keyword

files_list = list(pathlib.Path(dirPath).glob(key_word))

path = 'D:\\' + 'file_list.csv'  # path important!
f = open(path, 'w')  # write
for file in files_list:
    str_file = str(file)
    fileName = str_file.split("\\")[-1]  # 忽略檔案路徑，只顯示檔名
    if not (fileName[0] == "."):        # 開頭為點的隱藏檔案不show出來
        # file_info_time = time.ctime(os.path.getctime(file))
        f.write(fileName + ',' + time.strftime(
            "%Y-%m-%d, %H:%M:%S, %w",  time.gmtime(os.path.getctime(file))) + ',' + str_file + '\n')

f.close()
