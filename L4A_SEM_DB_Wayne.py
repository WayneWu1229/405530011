import tkinter as tk
import matplotlib as plt
import pyodbc
import getpass
import re
import os
import sqlite3
import io

from matplotlib import pyplot as plt
from tkinter import filedialog, ttk,  messagebox
from PIL import Image, ImageTk
from urllib.request import urlopen

# global variable
entry_sheetID = tk.Entry
entry_Position = tk.Entry
entry_type = tk.Entry
entry_value = tk.Entry
feedback = ""

# 設定介面視窗
root = tk.Tk()
root.title('L4A SEM 影像辨識')
root.geometry('1000x680')

# 設定主畫面 tab1、tab 頁籤
tabControl = ttk.Notebook(root)  # Create Tab Control
tabControl.pack(expand=1, fill="both")  # Pack to make visible
tab = ttk.Frame(tabControl)  # Add a second tab
tabControl.add(tab, text='   修改 / 刪除資料庫檔案   ')  # Make second tab visible

sheetID_val = tk.Entry(root, show=None)

# 設定tab開啟介面
sheetID_label_tab = tk.Label(
    tab, text='Sheet ID').place(x=20, y=10, anchor='nw')
sheetID_val_tab = tk.Entry(tab, show=None)
sheetID_val_tab.place(x=90, y=10, anchor='nw')

Position_label_tab = tk.Label(tab, text='點位').place(x=20, y=40, anchor='nw')
Position_val_tab = ttk.Combobox(
    tab, value=['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'J'], width=2)
Position_val_tab.place(x=90, y=40, anchor='nw')


mydb = sqlite3.connect("D:\\SQLite\\SEM_DB.sqlite")
cursor = mydb.cursor()


'#### 刪除資料庫資料 ####'


def deletDB():
    answer = tk.messagebox.askyesno(
        title='刪除', message='確定要刪除資料庫' + str(feedback.split(":")[0]) + '資訊?')
    if answer == True:
        SQL = "DELETE FROM E1  Where (Sheet_ID = '" + feedback.split(":")[0] + "') AND (Position = '" + feedback.split(":")[1].split("-")[
        0] + "') AND (Type = '" + feedback.split(":")[1].split("-")[1] + "')"
        cursor.execute(SQL)
        mydb.commit()
        lb.delete(lb.curselection())
        tk.messagebox.showinfo('提醒', "刪除成功")


def update():
    SQL = "UPDATE E1 SET Value='"+ entry_value.get()+ "' Where (Sheet_ID = '" + feedback.split(":")[0] + "') AND (Position = '" + feedback.split(":")[1].split("-")[
        0] + "') AND (Type = '" + feedback.split(":")[1].split("-")[1] + "')"
    cursor.execute(SQL)
    mydb.commit()

    tk.messagebox.showinfo('提醒', "修改成功")



def search():
    lb.delete(0, 'end')

    list_box = []

    if (sheetID_val_tab.get() == '') or (Position_val_tab.get() == ''):
        tk.messagebox.showerror("錯誤", "請輸入Sheet ID及點位")
    else:
        SQL = "select Sheet_ID, Position, Type, Value from E1 "
        SQL = SQL + "where (Sheet_ID = '" + sheetID_val_tab.get()+"')"
        SQL = SQL + " AND (Position = '" + Position_val_tab.get()+"')"
        # print(SQL)

        cursor.execute(SQL)
        for row in cursor:
            it = str(row[0]) + ':' + str(row[1]) + \
                '-' + str(row[2])
            list_box.append(it)
        
        list_box = list(set(list_box))
        if len(list_box) == 0:
            tk.messagebox.showerror("錯誤", "找不到相關資料")
        else:
            for item in list_box:
                lb.insert('end', item)
            lb.place(width=270, x=20, y=70, anchor='nw')


def CallOn(event):
    global feedback,entry_sheetID,entry_Position,entry_type,entry_value
    feedback = lb.get(lb.curselection())
    SQL = "select Sheet_ID , Position , Value , Type , Image "
    SQL = SQL + "From E1 "
    SQL = SQL + "Where (Sheet_ID = '" + feedback.split(":")[0] + "') AND (Position = '" + feedback.split(":")[1].split("-")[
        0] + "') AND (Type = '" + feedback.split(":")[1].split("-")[1] + "')"
    cursor.execute(SQL)

    img_url_tab = ""
    for row in cursor:
        sheetID_tab = row[0]
        Position_tab = row[1]
        Value_tab = row[2]
        type_tab = row[3]
        img_url_tab = row[4]

    '# 顯示圖片 #'
    image_byte = urlopen(img_url_tab).read()
    data_stream = io.BytesIO(image_byte)
    pil_image = Image.open(data_stream)
    photo = ImageTk.PhotoImage(pil_image)
    label_img = tk.Label(tab, image=photo)
    label_img.image = photo
    label_img.place(x=320, y=70, anchor='nw')

    '# 顯示sheet ID #'
    tk.Label(tab, text='Sheet ID').place(x=30, y=250, anchor='nw')
    v = tk.StringVar(tab, sheetID_tab)
    entry_sheetID = tk.Entry(tab, textvariable=v,state = 'disable')
    entry_sheetID.place(x=100, y=250, anchor='nw')

    '# 顯示點位 #'
    tk.Label(tab, text='點位').place(x=30, y=280, anchor='nw')
    v = tk.StringVar(tab, Position_tab)
    entry_Position = tk.Entry(tab, textvariable=v,state = 'disable')
    entry_Position.place(x=100, y=280, anchor='nw')

    '# 顯示類別 #'
    tk.Label(tab, text='Type').place(x=30, y=310, anchor='nw')
    v = tk.StringVar(tab, type_tab)
    entry_type = tk.Entry(tab, textvariable=v,state = 'disable')
    entry_type.place(x=100, y=310, anchor='nw')

    '# 顯示數值 #'
    tk.Label(tab, text='Value').place(x=30, y=340, anchor='nw')
    v = tk.StringVar(tab, Value_tab)
    entry_value = tk.Entry(tab, textvariable=v)
    entry_value.place(x=100, y=340, anchor='nw')

    '# 修改資料 #'
    btndeletDB = tk.Button(tab, width=29, text="修改資料", command=update,
                           fg='white', bg="green").place(x=400, y=570, anchor='nw')

    '# 刪除資料 #'
    btndeletDB = tk.Button(tab, width=29, text="刪除資料", command=deletDB,
                           fg='white', bg="red").place(x=660, y=570, anchor='nw')


lb = tk.Listbox(tab)
lb.bind('<Double-Button-1>', CallOn)
search_button = tk.Button(tab, text='搜尋', command=search).place(
    x=250, y=35, anchor='nw')

root.mainloop()