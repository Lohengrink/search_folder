import pathlib
import os
import time
import tkinter as tk
from win32com.client import Dispatch
from tkinter import filedialog
from tkinter import *
'''
# os.path.getatime(file) 輸出檔案訪問時間
# os.path.getctime(file) 輸出檔案的建立時間
# os.path.getmtime(file) 輸出檔案最近修改時間
'''

dirPath = r'D:\SEL論文'  # searching path
# dirPath_tmp = input(
#     'input the folder you want to search(ex: F:\):  ')  # searching path
# dirPath = dirPath_tmp + '\\'

key_word = '**\*2014*.pdf'  # searching keyword
# searching keyword
# key_word_tmp = input('input the key word(ex: *.pdf, *.xlxs):  ')
# key_word = '**\\' + key_word_tmp


files_list = list(pathlib.Path(dirPath).glob(key_word))

path = 'D:\\' + 'file_list.csv'  # path important!
xl = Dispatch("Excel.Application")  # 打開excel的應用程式
wb = xl.Workbooks.Open(path)
wb.Close()


f = open(path, 'w')  # write
for file in files_list:
    str_file = str(file)
    fileName = str_file.split("\\")[-1]  # 忽略檔案路徑，只顯示檔名
    if not (fileName[0] == "."):        # 開頭為點的隱藏檔案不show出來
        # file_info_time = time.ctime(os.path.getctime(file))
        f.write(fileName + ',' + time.strftime(
            "%Y-%m-%d, %H:%M:%S, %w",  time.gmtime(os.path.getctime(file))) + ',' + str_file + '\n')

f.close()


xl.Visible = True  # otherwise excel is hidden

# newest excel does not accept forward slash in path
wb = xl.Workbooks.Open(path)


root = tk.Tk()
# 視窗標題
root.title('hello')
# 寬度 200
# 高度 250
# 螢幕位置 X 300
# 螢幕位置 Y 400
root.geometry('200x250+300+400')
# 運行視窗


def browse_button():
    filename = filedialog.askopenfilenames()
    print("filename")
    print(filename)


    # return filename
tk.Button(text="Browse", command=browse_button).pack()
# .grid(row=0, column=3)

# v = StringVar()
root.mainloop()
