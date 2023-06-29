# import numpy as np
# import pandas as pd
import yfinance as yf
# import matplotlib.pyplot as plt
import datetime
# from datetime import date, timedelta
# ##新增:將趨勢圖傳送到Line群組
# import requests
# from io import BytesIO
# from PIL import Image
# import lineTool
# import json
# import time
# import sys
# import os
# #import xlwt             # pip install xlwt
# from openpyxl import Workbook, load_workbook
# from openpyxl.drawing.image import Image as XLImage

class Finance:
    def __init__(self):
        self.symbol = {}

    def getDate():
        DateTime = datetime.datetime.now()
        Y1 = DateTime.year
        M1 = DateTime.month
        D1 = DateTime.day
        dateStr = str(Y1)+"/"+str(M1)+"/"+str(D1)    
        return dateStr
    
    # def getData(self, symbol):
    #     # 下載股價資料-最近1年
    #     df = yf.download(symbol, period='1y')
        
    #     # 計算移動平均線
    #     df['MA20'] = df['Close'].rolling(window=20).mean()
    #     df['MA60'] = df['Close'].rolling(window=60).mean()

    #     # 從資料中提取最高價（high）和最低價（low） 
    #     high_values = df['High'].values 
    #     low_values = df['Low'].values