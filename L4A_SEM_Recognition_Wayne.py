'''
SEM_Photo 類別:
    屬性:
        img：原始照片
        sheetID：照片的sheet id
        type：照片的分類
        Position：點位
        P_S：面內、面外
        lotID：所屬的lit id
        platform：所屬的平台
        Abbr：所屬的Abbr.
        SPT_Run_Time：SPT_Run貨時間
        SPT_EQ：SPT機台
        AL_SPT_CH
        WMA_Run_Time：WMA_Run貨時間
        WMA_EQ：WMA機台
        OCR_data：辨識後的膜厚陣列
        OCR_taper：辨識後的taper陣列
        file_name：照片的檔名
        dilation2：第二次膨脹的圖
        region：一張SEM影響可能的目標區域陣列
        img_plate：初次裁切的圖
        cut_3：經水平、垂直裁切的圖
      方法:
        preprocess(self)：影像處理
        findPlateNumberRegion(self)：尋找區域輪廓
        cut(self)：裁切目標區域
        OCR(self , x1 , x2 , y1 , y2)：辨識影像上的數字
        catch_data(self)：利用self.sheetID到資料庫撈取資料
        detect(self)：將以上方法彙整再一起，未來只要呼叫detect方法
global變數：
    SEM陣列：儲存每一張SEM物件的陣列，因displayz、save、ExcelPrint、changePic方法都會使用到
    current：目前介面顯示的位置(0:第一張SEM、1:第二張SEM......)，因save、changePic、clear方法都會使用到
    data_val , taper_val：顯示在介面，膜厚和taper的輸入框，因display、clear方法會用到
方法:
   display(index)：介面顯示，index代表顯示的頁數
   save()：儲存修改後的辨識數據(使用者一張一張儲存)
   ExcelPrint()：給’匯出Excel’按鈕呼叫的方法，將數據Print在Excel中
   changePic(flag)：調整上下頁，flag = +1代表下一頁，flag = -1代表上一頁
   btnPreClick()：給’上一頁’按鈕呼叫的方法，當按鈕按下後會執行changePic(-1)方法
   btnNextClick()：給’下一頁’按鈕呼叫的方法，當按鈕按下後會執行changePic(+1)方法
   Clear()：清除數據資料，當跳換上下頁時，需先清除原先顯示的資料
   tunePhoto()：給’瀏覽檔案’按鈕呼叫的方法，可多選圖，並個別建立SEM_Photo物件，執行類別方法
'''


# -*- coding: utf-8 -*-
import tkinter as tk
import numpy as np
import cv2
import pytesseract
import matplotlib as plt
import xlwt
import xlrd
import datetime
import pyodbc
import getpass
import time
import re
import os
import check
import pandas as pd
import sqlite3
import io
import requests

from matplotlib import pyplot as plt
from tkinter import filedialog, ttk,  messagebox
from tkinter.ttk import Progressbar
from PIL import Image, ImageTk
from xlutils.copy import copy as cpy
from urllib.request import urlopen

'''定義照片要對應的參數'''
type_dict = {'M1-TH': ['M1_TOP_MO', 'GSH', 'PV', 'GSH_TAPER', 'PV_TAPER'], 'M1-MO': 'M1_TOP_MO_REMAIN', 'M2-TH': ['M2_TOP_MO', 'PV', 'PV_TAPER'], 'M2-MO': 'M2_TOP_MO_REMAIN', 'SD': ['ASH_REMAIN', 'ASH+AL+N+'], 'AS': 'AS_TAPER', 'AA-M2-TH': 'TH_TAPER'
             }

"""定義 pytesseract 的路徑"""
pytesseract.tesseract_cmd = r'D:\ProgramData\tesseract\tessdata'
tessdata_dir_config = r'--tessdata-dir "D:\ProgramData\tesseract\tessdata"'

# global variable
data_val = {}
taper_val = {}
entry_sheetID = tk.Entry
entry_Position = tk.Entry
entry_type = tk.Entry
entry_value = tk.Entry
login_id = ""
login_dept = ""
feedback = ""

product = ""
TFT_list = []
TP_list = []

# 設定介面視窗
root = tk.Tk()
root.title('L4A SEM 影像辨識')
root.geometry('1000x680')

# 設定主畫面 tab1、tab2 頁籤
tabControl = ttk.Notebook(root)  # Create Tab Control
tab1 = ttk.Frame(tabControl)  # Create a tab
tabControl.add(tab1, text='   檢視SEM影像   ')  # Add the tab
tabControl.pack(expand=1, fill="both")  # Pack to make visible

# ***********************************Tab 1:檢視SEM影像****************************************
# 設定tab1開啟介面
# sheetID_label=tk.Label(tab1 , text='Sheet ID').place(x=20, y=10, anchor='nw')
sheetID_val = tk.Entry(root, show=None)
# sheetID_val.place(x=90, y=10, anchor='nw')


class SEM_Photo:

    def __init__(self, img, sheetID):
        self.img = img
        self.sheetID = sheetID

        self.img_url = ''
        self.file = ''       # SEM影像左上角資訊
        self.Position = ''
        self.type = ''
        self.P_S = ''        # 面內/面外
        self.lotID = ''
        self.platform = ''   # 平台
        self.Abbr = ''
        self.SPT_Run_Time = ''
        self.SPT_EQ = ''
        self.AL_SPT_CH = ''
        self.WMA_Run_Time = ''
        self.WMA_EQ = ''
        self.OCR_data = []
        self.OCR_taper = []

        self.file_name = ''
        self.dilation2 = ''  # 第二次膨脹
        self.region = []    # 一張SEM影響可能的目標區域陣列
        self.img_plate = ''  # 初次裁切的圖
        self.cut_3 = ''     # 經水平、垂直裁切的圖

        # modify by Wayne
        self.doc = ""
        self.week = ""
        self.product = ""
        self.folder = ""
        self.dept = ""

    '#### 方法-影像處理 ####'

    def preprocess(self):

        gray = cv2.cvtColor(self.img, cv2.COLOR_BGR2GRAY)

        # 原圖(灰度圖)
        #cv2.imshow("1.Origin", gray)

        # 裁圖(去除底部)
        cut_1 = gray[0:400, 0:640]
        #cv2.imshow("2.cut", cut_1)

        # 高斯平滑
        gaussian = cv2.GaussianBlur(cut_1, (3, 3), 0, 0, cv2.BORDER_DEFAULT)
        median = cv2.medianBlur(gaussian, 1)
        #cv2.imshow("3.gaussian", median)

        # sobel
        sobel = cv2.Sobel(median, cv2.CV_8U, 1, 0,  ksize=3)
        #cv2.imshow("4.sobel", sobel)

        # 二值化
        ret, binary = cv2.threshold(sobel, 253, 255, cv2.THRESH_BINARY)
        #cv2.imshow("5.binary", binary)

        # 膨脹 I
        element1 = cv2.getStructuringElement(
            cv2.MORPH_RECT, (9, 4))  # 腐蝕操作的核函数
        element2 = cv2.getStructuringElement(
            cv2.MORPH_RECT, (8, 5))  # 膨脹操作的核函数
        dilation = cv2.dilate(binary, element2, iterations=2)
        #cv2.imshow("6.dilationI", dilation)

        # 腐蝕，去掉細節
        erosion = cv2.erode(dilation, element1, iterations=3)
        #cv2.imshow("7.erosion", erosion)

        # 膨脹II
        self.dilation2 = cv2.dilate(erosion, element2, iterations=2)
        #cv2.imshow("8.erosionII", self.dilation2)

    '#### 方法-尋找輪廓區域 ####'

    def findPlateNumberRegion(self):
        # 查找輪廓
        aa, contours, hierarchy = cv2.findContours(
            self.dilation2, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

        # 篩選面積小的
        for i in range(len(contours)):
            cnt = contours[i]
            area = cv2.contourArea(cnt)  # 計算輪廓面積
            #print('面積：' + str(area))

            # 面積小的篩選掉
            if (area < 600):
                continue

            # 找到最小的矩形，該矩形可能有方向
            rect = cv2.minAreaRect(cnt)
            #print ("rect is: ")
            #print (rect)

            # box是四個點的座標
            box = cv2.boxPoints(rect)
            box = np.int0(box)

            # 計算高和寬
            height = abs(box[0][1] - box[2][1])
            width = abs(box[0][0] - box[2][0])

            # 正常情况下長高比在15以內
            ratio = float(width) / float(height)
            #print('ratio：' + str(ratio))
            #print (ratio)
            if (ratio > 15):  # modify by Wayne(10)
                continue
            self.region.append(box)

    '#### 方法-水平、垂直裁切 ###'

    def cut(self, a):
        # 垂直投影
        GrayImage = cv2.cvtColor(
            self.img_plate, cv2.COLOR_BGR2GRAY)  # 將BGR圖轉為灰度圖
        # 將圖片進行二值化（253,255）之間的點均變為255（背景）
        ret, thresh1 = cv2.threshold(GrayImage, 253, 255, cv2.THRESH_BINARY)
        (h, w) = thresh1.shape  # 返回高和寬
        a = [0 for z in range(0, w)]
        for j in range(0, w):  # 遍歷一行
            for i in range(0, h):  # 遍歷一列
                if thresh1[i, j] == 0:  # 如果該點為黑點
                    a[j] += 1  # 該列的計數器加一計數
                    thresh1[i, j] = 255  # 紀錄完後將其變為白色
        for j in range(0, w):  # 遍歷每一列
            for i in range((h-a[j]), h):  # 從該列應該變黑的最頂部的點開始向最底部塗黑
                thresh1[i, j] = 0  # 塗黑
        # print(a)
        # plt.imshow(thresh1,cmap=plt.gray()) # 顯示垂直投影圖
        # plt.show()

        # 裁切(裁切左右)
        count_r = 0
        count_l = 0
        for j in range(0, w):
            if a[j] == h or a[j] == (h-1):
                if j > (w/2):
                    count_r = count_r+1
                else:
                    count_l = count_l+1
        cut_2 = self.img_plate[0:h, (0+count_l):(w-count_r)]

        size1 = cut_2.shape

        # 水平投影
        GrayImage2 = cv2.cvtColor(cut_2, cv2.COLOR_BGR2GRAY)  # 將BGR圖轉為灰度圖
        # 將圖片進行二值化（253,255）之間的點均變為255（背景）
        ret, thresh2 = cv2.threshold(GrayImage2, 253, 255, cv2.THRESH_BINARY)
        (h2, w2) = thresh2.shape  # 返回高和寬
        b = [0 for z in range(0, h2)]
        for j in range(0, h2):
            for i in range(0, w2):
                if thresh2[j, i] == 0:
                    b[j] += 1
                    thresh2[j, i] = 255
        for j in range(0, h2):
            for i in range(0, b[j]):
                thresh2[j, i] = 0
        # print(b)
        # plt.imshow(thresh2,cmap=plt.gray())
        # plt.show()

        # 裁切(裁切上下)
        count_h = 0
        count_b = 0
        for i in range(0, h2):
            if b[i] == w2 or b[i] == (w2-1):
                if i > (h2/2):
                    count_b = count_b+1
                else:
                    count_h = count_h+1
        self.cut_3 = cut_2[(0+count_h):(h2-count_b), 0:w2]

        # 利用雜訊大多不是全黑(0,0,0)，進行像素顏色轉換
        size = self.cut_3.shape  # modify by Wayne
        for i in range(size[0]):
            for j in range(size[1]):
                r, g, b = self.cut_3[i][j][0], self.cut_3[i][j][1], self.cut_3[i][j][2]
                if(r > 4 and g > 4 and b > 4):
                    self.cut_3[i][j][0] = 255
                    self.cut_3[i][j][1] = 255
                    self.cut_3[i][j][2] = 255
            # print(self.cut_3[i])

        #cv2.imshow('Origen'+str(box), thresh11)
        #cv2.imshow(str(self.cut_3), self.cut_3)
        #cv2.imwrite('messigray'+str(a)+'.tiff', self.cut_3)
        #cv2.imwrite('messigray'+'.tiff', self.cut_3)

    '#### 方法-辨識(膜厚、Taper) ####'

    def OCR(self, x1, x2, y1, y2):
        # OCR_result = pytesseract.image_to_string(self.cut_3 , lang='eng', config=tessdata_dir_config)
        # modify by Wayne (Only accept digits & Upper characters),retrained the model
        OCR_result = pytesseract.image_to_string(
            self.cut_3, lang='TrainDigit+TrainSheetID', config="-psm 7 digits")  # 訓練模型更動 TrainTest -> TrainDigit
        OCR_result = OCR_result.upper()
        print(OCR_result)

        temp = []  # 判斷OCR_result 是否包含數字 (避免錯誤辨識)
        for i in range(len(OCR_result)):
            if(OCR_result[i].isdigit()):
                temp.append(OCR_result[i])

        if (OCR_result != '' and temp != []):
            for c in OCR_result:
                if (ord(c) >= 48 and ord(c) <= 57) or (ord(c) >= 65 and ord(c) <= 90) or ord(c) == 46 or ord(c) == 45:  # 只留下數字、英文、點、dash
                    continue
                elif(ord(c) == 8212):  # modify by Wayne(留住 em dash AscII "8212")
                    OCR_result = OCR_result.replace(c, '-')
                OCR_result = OCR_result.replace(c, '')

            dash1 = '-'
            dash2 = '—'
            dot = '.'
            if (dash1 in OCR_result) == True or (dash2 in OCR_result) == True:
                self.file = OCR_result
                self.sheetID = OCR_result.split('-')[0]   # sheetID
                self.Position = OCR_result.split('-')[1]  # 點位
                # type # modify by Wayne (change 0 to O)
                self.type = OCR_result.split(
                    '-', 2)[-1].replace('0', 'O').replace('Q', 'O')
                # 因原始白名單沒有I，針對此產品添加
                if(self.type == "V1A-TH" or self.type == "VUA-TH"):
                    self.type = "VIA-TH"

            else:
                if OCR_result[-1] == 'A':
                    # modify by Wayne 判斷首位是否為數字
                    if(OCR_result[0].isdigit() and len(OCR_result) < 6):
                        self.OCR_data.append(OCR_result[:-1])  # 膜厚
                    else:
                        self.OCR_data.append(OCR_result[1:-1])
                    cv2.putText(self.img, str(len(self.OCR_data)), (x2+2, y1+15),
                                cv2.FONT_HERSHEY_COMPLEX, 0.5, (0, 255, 0), 1)  # 數字顯示易造成誤判 #(x1,y1-5)
                elif (dot in OCR_result) == True:
                    if(OCR_result[0].isdigit()):  # modify by Wayne 判斷首位是否為數字
                        self.OCR_taper.append(OCR_result)  # Taper
                    else:
                        self.OCR_taper.append(OCR_result[1:])
                    cv2.putText(self.img, 'Taper'+str(len(self.OCR_taper)),
                                (x1, y1-5), cv2.FONT_HERSHEY_COMPLEX, 0.5, (255, 0, 0), 1)

    '#### 方法-抓取資料庫 ####'

    def catch_data(self):
        # 回傳分類、lotID、sheetID、平台、Abbr.、SPT Run貨日期、SPT機台、AL SPT CH、WMA Run貨時間、WMA機台
        pass

    '#### 產出各張圖片資訊(膜厚、Taper) ####'

    def detect(self):

        # 影像處理
        self.preprocess()

        # 找出區域輪廓
        self.findPlateNumberRegion()

        # 依各區域進行裁切、辨識
        for box in self.region:
            # cv2.drawContours(self.img, [box], 0, (0, 255, 0), 2) # 需考慮是否於辨識後再畫出邊框(邊框易造成辨識錯誤)
            ys = [box[0, 1], box[1, 1], box[2, 1], box[3, 1]]
            xs = [box[0, 0], box[1, 0], box[2, 0], box[3, 0]]
            ys_sorted_index = np.argsort(ys)  # 由小到大排序索引直
            xs_sorted_index = np.argsort(xs)  # 由小到大排序索引直

            x1 = box[xs_sorted_index[0], 0]
            x2 = box[xs_sorted_index[3], 0]

            y1 = box[ys_sorted_index[0], 1]
            y2 = box[ys_sorted_index[3], 1]

            img_org2 = self.img.copy()
            self.img_plate = img_org2[y1:y2, x1:x2]

            # 垂直、水平裁切
            self.cut(len(box))

            # 辨識數據
            self.OCR(x1, x2, y1, y2)

        for box in self.region:
            cv2.drawContours(self.img, [box], 0, (0, 255, 0), 2)


'#### 介面顯示 ####'


def display(index):
    for widget in tab1.winfo_children():  # 清除物件
        widget.destroy()

    data = SEM[index].OCR_data
    taper = SEM[index].OCR_taper
    global entry_sheetID, entry_Position, entry_type, data_val, taper_val

    try:
        '# 顯示圖片 #'
        ndarray_convert_img = Image.fromarray(SEM[index].img)
        # photo = ImageTk.PhotoImage(ndarray_convert_img.resize((600,450))) #縮放
        photo = ImageTk.PhotoImage(ndarray_convert_img)
        label_img = tk.Label(tab1, image=photo)
        label_img.image = photo
        label_img.place(x=30, y=80, anchor='nw')

        '# 顯示檔名 #'
        tk.Label(tab1, text='檔名').place(x=700, y=120, anchor='nw')
        tk.Label(tab1, text=SEM[index].file_name).place(
            x=800, y=120, anchor='nw')

        '# 顯示sheet ID #'
        tk.Label(tab1, text='Sheet ID').place(x=700, y=150, anchor='nw')
        v = tk.StringVar(tab1, SEM[index].sheetID)
        entry_sheetID = tk.Entry(tab1, textvariable=v)
        entry_sheetID.place(x=800, y=150, anchor='nw')

        '# 顯示點位 #'
        tk.Label(tab1, text='點位').place(x=700, y=180, anchor='nw')
        v = tk.StringVar(tab1, SEM[index].Position)
        entry_Position = tk.Entry(tab1, textvariable=v)
        entry_Position.place(x=800, y=180, anchor='nw')

        '# 顯示類別 #'
        tk.Label(tab1, text='Type').place(x=700, y=210, anchor='nw')
        v = tk.StringVar(tab1, SEM[index].type)
        entry_type = tk.Entry(tab1, textvariable=v)
        entry_type.place(x=800, y=210, anchor='nw')

        '# 顯示膜厚、taper #'
        data_val = {}
        taper_val = {}
        M1TH = ["M1 Top Mo", "GSH", "PV", "GSH Taper",
                "PV Taper", "GSH Taper\nError"]
        TBTH = ["TB M1 Top Mo", "TB PV1 THK", "TB PV1 Taper"]
        # data
        if(SEM[index].product == "PSA" or SEM[index].product == "AHVA" or SEM[index].product == "TN"):
            if(SEM[index].type == "M1-TH"):
                for i in range(5):
                    data_pri = tk.Label(tab1, text=M1TH[i])
                    data_pri.place(x=700, y=(225+((i+1)*30)), anchor='nw')
                if(len(data) == 3):
                    for j in range(3):
                        v = tk.StringVar(tab1, value=data[j])
                        data_val[j] = tk.Entry(tab1, textvariable=v)
                        data_val[j].place(
                            x=800, y=(225+((j+1)*30)), anchor='nw')
                else:
                    for k in range(3):
                        v = tk.StringVar(tab1, value="NA")
                        data_val[k] = tk.Entry(tab1, textvariable=v)
                        data_val[k].place(
                            x=800, y=(225+((k+1)*30)), anchor='nw')
                    tk.messagebox.showinfo("提醒", "此片有遺漏值，請自行補值")

                if(len(taper) == 2):
                    for k in range(2):
                        v = tk.StringVar(tab1, value=taper[k])
                        taper_val[k] = tk.Entry(tab1, textvariable=v)
                        taper_val[k].place(
                            x=800, y=(225+((k+4)*30)), anchor='nw')
                else:  # 因應異常貨，將NG Taper輸出
                    data_pri = tk.Label(tab1, text=M1TH[5])
                    data_pri.place(x=700, y=(225+(6*30)), anchor='nw')
                    for k in range(2, -1, -1):
                        if(k == 0):
                            v = tk.StringVar(tab1, value=taper[k])
                            taper_val[k] = tk.Entry(tab1, textvariable=v)
                            taper_val[k].place(
                                x=800, y=(225+(6*30)), anchor='nw')
                        else:
                            v = tk.StringVar(tab1, value=taper[k])
                            taper_val[k] = tk.Entry(tab1, textvariable=v)
                            taper_val[k].place(
                                x=800, y=(225+((k+3)*30)), anchor='nw')
                    tk.messagebox.showerror("警告", "該片為NG貨")

            elif(SEM[index].type == "M1-MO"):
                data_pri = tk.Label(tab1, text="M1 Top Mo\n Remain")
                data_pri.place(x=700, y=(225+(1*30)), anchor='nw')

                if(len(data) == 1):
                    v = tk.StringVar(tab1, value=data[0])
                    data_val[0] = tk.Entry(tab1, textvariable=v)
                    data_val[0].place(x=800, y=(225+(1*30)), anchor='nw')
                else:
                    v = tk.StringVar(tab1, value="NA")
                    data_val[0] = tk.Entry(tab1, textvariable=v)
                    data_val[0].place(x=800, y=(225+(1*30)), anchor='nw')
                    tk.messagebox.showinfo("提醒", "此片有遺漏值，請自行補值")

            elif(SEM[index].type == "M2-TH"):
                data_pri = tk.Label(tab1, text="M2 Top Mo")
                data_pri.place(x=700, y=(225+(1*30)), anchor='nw')
                if(len(data) == 1):
                    v = tk.StringVar(tab1, value=data[0])
                    data_val[0] = tk.Entry(tab1, textvariable=v)
                    data_val[0].place(x=800, y=(225+(1*30)), anchor='nw')
                else:
                    v = tk.StringVar(tab1, value="NA")
                    data_val[0] = tk.Entry(tab1, textvariable=v)
                    data_val[0].place(x=800, y=(225+(1*30)), anchor='nw')
                    tk.messagebox.showinfo("提醒", "此片有遺漏值，請自行補值")

            elif(SEM[index].type == "M2-MO"):
                data_pri = tk.Label(tab1, text="M2 Top Mo\n Remain")
                data_pri.place(x=700, y=(225+(1*30)), anchor='nw')
                if(len(data) == 1):
                    v = tk.StringVar(tab1, value=data[0])
                    data_val[0] = tk.Entry(tab1, textvariable=v)
                    data_val[0].place(x=800, y=(225+(1*30)), anchor='nw')
                else:
                    v = tk.StringVar(tab1, value="NA")
                    data_val[0] = tk.Entry(tab1, textvariable=v)
                    data_val[0].place(x=800, y=(225+(1*30)), anchor='nw')
                    tk.messagebox.showinfo("提醒", "此片有遺漏值，請自行補值")

            elif(SEM[index].type == "SD"):
                data_pri = tk.Label(tab1, text="ASH+AL+N+")
                data_pri.place(x=700, y=(225+(1*30)), anchor='nw')
                data_pri = tk.Label(tab1, text="ASH Remain")
                data_pri.place(x=700, y=(225+(2*30)), anchor='nw')
                if(len(data) == 2):
                    if(int(data[0]) > int(data[1])):
                        v = tk.StringVar(tab1, value=data[0])
                        data_val[0] = tk.Entry(tab1, textvariable=v)
                        data_val[0].place(x=800, y=(225+(1*30)), anchor='nw')
                        v1 = tk.StringVar(tab1, value=data[1])
                        data_val[1] = tk.Entry(tab1, textvariable=v1)
                        data_val[1].place(x=800, y=(225+(2*30)), anchor='nw')
                    else:
                        v = tk.StringVar(tab1, value=data[1])
                        data_val[0] = tk.Entry(tab1, textvariable=v)
                        data_val[0].place(x=800, y=(225+(1*30)), anchor='nw')
                        v1 = tk.StringVar(tab1, value=data[0])
                        data_val[1] = tk.Entry(tab1, textvariable=v1)
                        data_val[1].place(x=800, y=(225+(2*30)), anchor='nw')
                else:
                    for k in range(2):
                        v = tk.StringVar(tab1, value="NA")
                        data_val[k] = tk.Entry(tab1, textvariable=v)
                        data_val[k].place(
                            x=800, y=(225+((k+1)*30)), anchor='nw')
                    tk.messagebox.showinfo("提醒", "此片有遺漏值，請自行補值")

            elif(SEM[index].type == "AS"):
                data_pri = tk.Label(tab1, text="AS Taper")
                data_pri.place(x=700, y=(225+(1*30)), anchor='nw')
                if(len(taper) == 1):
                    v = tk.StringVar(tab1, value=taper[0])
                    taper_val[0] = tk.Entry(tab1, textvariable=v)
                    taper_val[0].place(x=800, y=(225+(1*30)), anchor='nw')
                else:
                    v = tk.StringVar(tab1, value="NA")
                    taper_val[0] = tk.Entry(tab1, textvariable=v)
                    taper_val[0].place(x=800, y=(225+(1*30)), anchor='nw')
                    tk.messagebox.showinfo("提醒", "此片有遺漏值，請自行補值")

            elif(SEM[index].type == "AA-M2-TH"):
                data_pri = tk.Label(tab1, text="TH Taper")
                data_pri.place(x=700, y=(225+(1*30)), anchor='nw')
                if(len(taper) == 1):
                    v = tk.StringVar(tab1, value=taper[0])
                    taper_val[0] = tk.Entry(tab1, textvariable=v)
                    taper_val[0].place(x=800, y=(225+(1*30)), anchor='nw')
                else:
                    v = tk.StringVar(tab1, value="NA")
                    taper_val[0] = tk.Entry(tab1, textvariable=v)
                    taper_val[0].place(x=800, y=(225+(1*30)), anchor='nw')
                    tk.messagebox.showinfo("提醒", "此片有遺漏值，請自行補值")

        elif(SEM[index].product == "TP" or SEM[index].product == "A2GP"):
            if(SEM[index].type == "VIA-TH"):
                data_pri = tk.Label(tab1, text="VIA PV1 Taper")
                data_pri.place(x=700, y=(225+(1*30)), anchor='nw')
                data_pri = tk.Label(tab1, text="VIA PV1 THK")
                data_pri.place(x=700, y=(225+(2*30)), anchor='nw')
                if(len(taper) == 1):
                    v = tk.StringVar(tab1, value=taper[0])
                    taper_val[0] = tk.Entry(tab1, textvariable=v)
                    taper_val[0].place(x=800, y=(225+(1*30)), anchor='nw')
                else:
                    v = tk.StringVar(tab1, value="NA")
                    taper_val[0] = tk.Entry(tab1, textvariable=v)
                    taper_val[0].place(x=800, y=(225+(1*30)), anchor='nw')
                    tk.messagebox.showinfo("提醒", "此片有遺漏值，請自行補值")

                if(len(data) == 1):
                    v = tk.StringVar(tab1, value=data[0])
                    data_val[0] = tk.Entry(tab1, textvariable=v)
                    data_val[0].place(x=800, y=(225+(2*30)), anchor='nw')
                else:
                    v = tk.StringVar(tab1, value="NA")
                    data_val[0] = tk.Entry(tab1, textvariable=v)
                    data_val[0].place(x=800, y=(225+(2*30)), anchor='nw')
                    tk.messagebox.showinfo("提醒", "此片有遺漏值，請自行補值")

            elif(SEM[index].type == "TB-TH"):
                for i in range(3):
                    data_pri = tk.Label(tab1, text=TBTH[i])
                    data_pri.place(x=700, y=(225+((i+1)*30)), anchor='nw')

                if(len(taper) == 1):
                    v = tk.StringVar(tab1, value=taper[0])
                    taper_val[0] = tk.Entry(tab1, textvariable=v)
                    taper_val[0].place(x=800, y=(225+(3*30)), anchor='nw')
                else:
                    v = tk.StringVar(tab1, value="NA")
                    taper_val[0] = tk.Entry(tab1, textvariable=v)
                    taper_val[0].place(x=800, y=(225+(3*30)), anchor='nw')
                    tk.messagebox.showinfo("提醒", "此片有遺漏值，請自行補值")

                if(len(data) == 2):
                    for j in range(2):
                        v = tk.StringVar(tab1, value=data[j])
                        data_val[j] = tk.Entry(tab1, textvariable=v)
                        data_val[j].place(
                            x=800, y=(225+((j+1)*30)), anchor='nw')
                else:
                    for k in range(2):
                        v = tk.StringVar(tab1, value="NA")
                        data_val[k] = tk.Entry(tab1, textvariable=v)
                        data_val[k].place(
                            x=800, y=(225+((k+1)*30)), anchor='nw')
                    tk.messagebox.showinfo("提醒", "此片有遺漏值，請自行補值")

            elif(SEM[index].type == "TB-MO"):
                data_pri = tk.Label(tab1, text="TB M1 Top\n Mo Remain")
                data_pri.place(x=700, y=(225+(1*30)), anchor='nw')

                if(len(data) == 1):
                    v = tk.StringVar(tab1, value=data[0])
                    data_val[0] = tk.Entry(tab1, textvariable=v)
                    data_val[0].place(x=800, y=(225+(1*30)), anchor='nw')
                else:
                    v = tk.StringVar(tab1, value="NA")
                    data_val[0] = tk.Entry(tab1, textvariable=v)
                    data_val[0].place(x=800, y=(225+(1*30)), anchor='nw')
                    tk.messagebox.showinfo("提醒", "此片有遺漏值，請自行補值")

            elif(SEM[index].type == "PV2-TH"):
                data_pri = tk.Label(tab1, text="PV2 Taper")
                data_pri.place(x=700, y=(225+(1*30)), anchor='nw')
                if(len(taper) == 1):
                    v = tk.StringVar(tab1, value=taper[0])
                    taper_val[0] = tk.Entry(tab1, textvariable=v)
                    taper_val[0].place(x=800, y=(225+(1*30)), anchor='nw')
                else:
                    v = tk.StringVar(tab1, value="NA")
                    taper_val[0] = tk.Entry(tab1, textvariable=v)
                    taper_val[0].place(x=800, y=(225+(1*30)), anchor='nw')
                    tk.messagebox.showinfo("提醒", "此片有遺漏值，請自行補值")

        else:
            tk.messagebox.showerror("警告", "該片異常")

    except:
        tk.messagebox.showerror("警告", "該片異常")

    finally:
        '# 顯示上一頁、下一頁的按鈕 #'
        btnPre = tk.Button(tab1, text="上一張", command=btnPreClick).place(
            x=740, y=515, anchor='nw')
        tk.Label(tab1, text=(str(index+1) + " / " + str(len(SEM))),
                 width=6).place(x=795, y=515, anchor='nw')
        btnNext = tk.Button(tab1, text="下一張", command=btnNextClick).place(
            x=850, y=515, anchor='nw')

        '# 刪除功能 #'
        btndelet = tk.Button(tab1, width=29, text="刪除", command=delet,
                             fg='white', bg="red").place(x=705, y=550, anchor='nw')

        '# 寫入Database #'
        btnWriteXls = tk.Button(tab1, width=29, text="寫入資料庫", command=WriteDB,
                                fg='white', bg="gray").place(x=705, y=600, anchor='nw')


'#### 上、下頁功能 ####'
current = 0


def changePic(flag):  # flag=-1表示上一個，flag=1表示下一個
    global current
    current = current + flag

    if current < 0:
        current = len(SEM)-1
        display(current)
        tk.Label(tab1, text=(str(current+1) + " / " + str(len(SEM))),
                 width=6).place(x=795, y=520, anchor='nw')

    elif current >= len(SEM):
        current = 0
        display(current)
        tk.Label(tab1, text=(str(current+1) + " / " + str(len(SEM))),
                 width=6).place(x=795, y=520, anchor='nw')

    else:
        display(current)
        tk.Label(tab1, text=(str(current+1) + " / " + str(len(SEM))),
                 width=6).place(x=795, y=520, anchor='nw')

# 「上一張」按鈕


def btnPreClick():
    save()
    Clear()
    changePic(-1)

# 「下一張」按鈕


def btnNextClick():
    save()
    Clear()
    changePic(1)


def Clear():
    data_val.clear()
    taper_val.clear()


'#### 儲存功能 ####'


def save():
    # sheetID
    SEM[current].sheetID = entry_sheetID.get()

    # Position
    SEM[current].Position = entry_Position.get()

    # type
    SEM[current].type = entry_type.get()

    # data
    SEM[current].OCR_data = []
    for i in range(len(data_val)):
        if (data_val[i].get() != '0') and (data_val[i].get() != ''):
            SEM[current].OCR_data.append(data_val[i].get())

    # taper
    SEM[current].OCR_taper = []
    for i in range(len(taper_val)):
        if (taper_val[i].get() != '0') and (taper_val[i].get() != ''):
            SEM[current].OCR_taper.append(taper_val[i].get())


'#### 刪除功能 ####'


def delet():

    if len(SEM) > 1:
        answer = tk.messagebox.askyesno(
            title='刪除', message='確定要刪除' + str(SEM[current].file_name) + '?')
        if answer == True:
            del SEM[current]
            changePic(0)
    else:
        tk.messagebox.showwarning(title='警告', message='已是最後一張圖片! 刪除後將關閉程式')
        del SEM[current]
        root.destroy()


'#### 上傳tableau ####'


def data_to_tableau(path):
    data = {'site':'L4A','Path':'ML4AE1'}
    files = [('files',('SEM_DB.sqlite' , open(path + '\\SEM_DB.sqlite','rb')))]
    response = requests.post('http://autceda/files/MultiUpload', files = files,data = data)
    print(response.text)


'#### 寫入DataBase ####'


def WriteDB():
    mydb = sqlite3.connect("D:\\SQLite\\SEM_DB.sqlite")
    cursor = mydb.cursor()
    print("Opened database successfully")

    df = pd.DataFrame()

    for i in range(len(SEM)):
        temp = []
        imgPath = "http://10.88.40.45/SEM_picture/" + \
            SEM[i].folder + "/" + SEM[i].file_name
        doc = SEM[i].doc
        week = SEM[i].week
        sheetID = SEM[i].sheetID
        Position = SEM[i].Position

        if(SEM[i].dept == "ML4AE1"):
            if(SEM[i].product == "PSA" or SEM[i].product == "AHVA" or SEM[i].product == "TN"):
                if(SEM[i].type == "M1-TH"):
                    for j in range(3):
                        if(j == 0):
                            temp.append(
                                [doc, week, sheetID, Position, SEM[i].OCR_data[j], "M1 Top Mo", imgPath])
                        if(j == 1):
                            temp.append(
                                [doc, week, sheetID, Position, SEM[i].OCR_data[j], "GSH", imgPath])
                        if(j == 2):
                            temp.append(
                                [doc, week, sheetID, Position, SEM[i].OCR_data[j], "PV", imgPath])

                    if(len(SEM[i].OCR_taper) > 2):
                        temp.append([doc, week, sheetID, Position,
                                     SEM[i].OCR_taper[1], "GSH Taper", imgPath])
                        temp.append([doc, week, sheetID, Position,
                                     SEM[i].OCR_taper[2], "PV Taper", imgPath])
                        temp.append([doc, week, sheetID, Position,
                                     SEM[i].OCR_taper[0], "PV Taper", imgPath])

                    else:
                        temp.append([doc, week, sheetID, Position,
                                     SEM[i].OCR_taper[0], "GSH Taper", imgPath])
                        temp.append([doc, week, sheetID, Position,
                                     SEM[i].OCR_taper[1], "PV Taper", imgPath])

                elif(SEM[i].type == "M1-MO"):
                    temp.append([doc, week, sheetID, Position,
                                 SEM[i].OCR_data[0], "M1 Top Mo Remain", imgPath])

                elif(SEM[i].type == "M2-TH"):
                    temp.append([doc, week, sheetID, Position,
                                 SEM[i].OCR_data[0], "M2 Top Mo", imgPath])

                elif(SEM[i].type == "M2-MO"):
                    temp.append([doc, week, sheetID, Position,
                                 SEM[i].OCR_data[0], "M2 Top Mo Remain", imgPath])

                elif(SEM[i].type == "SD"):
                    temp.append([doc, week, sheetID, Position,
                                 SEM[i].OCR_data[0], "ASH+AL+N+", imgPath])
                    temp.append([doc, week, sheetID, Position,
                                 SEM[i].OCR_data[1], "ASH Remain", imgPath])

                elif(SEM[i].type == "AS"):
                    temp.append([doc, week, sheetID, Position,
                                 SEM[i].OCR_taper[0], "AS Taper", imgPath])

                elif(SEM[i].type == "AA-M2-TH"):
                    temp.append([doc, week, sheetID, Position,
                                 SEM[i].OCR_taper[0], "TH Taper", imgPath])

                else:
                    temp.append([doc, week, sheetID, Position,
                                 "Error", SEM[i].file_name, imgPath])

            elif(SEM[i].product == "TP" or SEM[i].product == "A2GP"):
                if(SEM[i].type == "VIA-TH"):
                    temp.append([doc, week, sheetID, Position,
                                 SEM[i].OCR_taper[0], "VIA PV1 Taper", imgPath])
                    temp.append([doc, week, sheetID, Position,
                                 SEM[i].OCR_data[0], "VIA PV1 THK", imgPath])

                elif(SEM[i].type == "TB-TH"):
                    temp.append([doc, week, sheetID, Position,
                                 SEM[i].OCR_taper[0], "TB PV1 Taper", imgPath])
                    temp.append([doc, week, sheetID, Position,
                                 SEM[i].OCR_data[0], "TB M1 Top Mo", imgPath])
                    temp.append([doc, week, sheetID, Position,
                                 SEM[i].OCR_data[1], "TB PV1 THK", imgPath])

                elif(SEM[i].type == "TB-MO"):
                    temp.append([doc, week, sheetID, Position,
                                 SEM[i].OCR_data[0], "TB M1 Top Mo Remain", imgPath])

                elif(SEM[i].type == "PV2-TH"):
                    temp.append([doc, week, sheetID, Position,
                                 SEM[i].OCR_taper[0], "PV2 Taper", imgPath])

                else:
                    temp.append([doc, week, sheetID, Position,
                                 "Error", SEM[i].file_name, imgPath])
            df = df.append(temp)

    df = df.rename({0: "Doc", 1: "Week", 2: "Sheet_ID", 3: "Position",
                    4: "Value", 5: "Type", 6: "Image"}, axis='columns')
    df.to_sql(name="E1", con=mydb, if_exists='append', index=False)
    mydb.commit()
    cursor = cursor.execute(
        "delete from E1 where E1.rowid not in (select MAX(E1.rowid) from E1 group by Sheet_ID,Position,Type);")
    mydb.commit()
    mydb.close()
    data_to_tableau("D:\\SQLite")
    tk.messagebox.showinfo('提醒', "寫入成功")


'#### 連接瀏覽檔案按鈕 ####'
click_count = 0
SEM = []


def load():
    path = "./SEM_picture/"
    load_list = check.checkData(path, "local")
    file_name = []  # 存放IMAGE檔案

    for i in range(len(load_list)):
        dir_list = os.listdir(path+load_list[i])
        for file in dir_list:  # 只留下IMAGE檔案
            if(file[0] == 'I'):
                file_name.append(path+load_list[i]+"/"+file)

    canvas = tk.Canvas(tab1, width=465, height=22, bg="white")  # 進度條
    canvas.place(x=300, y=300)
    fill_line = canvas.create_rectangle(
        1.5, 1.5, 0, 23, width=0, fill="green")
    x = len(file_name)
    n = 465 / x  # 465 是矩形填充滿的次數
    tk.Label(tab1, text="辨識中，請稍候",
             width=20).place(x=450, y=350, anchor='nw')

    for i in range(len(file_name)):
        img = cv2.imread(str(file_name[i]))
        sheetID = sheetID_val.get()
        SEM.append(SEM_Photo(img, sheetID))
        name = file_name[i].split('/', )[-1]
        # modify by Wayne 獲得單號、週別
        folder_t = file_name[i].split('/', )[2]
        doc_t = file_name[i].split('/', )[2].split('-', )[2]
        week_t = file_name[i].split('/', )[2].split('-', )[3]
        dept_t = file_name[i].split('/', )[2].split('-', )[4]
        product_t = file_name[i].split('/', )[2].split('-', )[5]
        # print(file_name[i])
        try:
            SEM[i].detect()
        except ValueError:
            continue
        except:
            continue
        finally:
            SEM[i].file_name = name
            SEM[i].img_url = file_name[i]
            SEM[i].folder = folder_t
            SEM[i].doc = doc_t
            SEM[i].week = week_t
            SEM[i].dept = dept_t
            SEM[i].product = product_t

        n = n + 465 / x  # 進度條進度
        canvas.coords(fill_line, (0, 0, n, 60))
        root.update()
        time.sleep(0.02)  # 控制進度條流動的次數

    '# 顯示外框#'
    cv = tk.Canvas(tab1, width=925, height=520)
    cv.create_rectangle(5, 5, 925, 520, outline='gray', fill='')
    cv.place(x=15, y=65)

    display(0)


load()

root.mainloop()
