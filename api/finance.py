import numpy as np
import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt
import datetime
from datetime import date, timedelta
##新增:將趨勢圖傳送到Line群組
import requests
from io import BytesIO
from PIL import Image
import lineTool
import json
import time
import sys
import os
#import xlwt             # pip install xlwt
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage

class Finance:
    def __init__(self):
        # ETF, 績優股, 金融股
        self.ETF = {'00878.TW': '國泰永續高股息', '00919.TW': '群益台灣精選高息', '00713.TW': '元大台灣高息低波', '0056.TW': '元大高股息','006208.TW': '富邦台灣50', '00690.TW': '兆豐藍籌30', '00850.TW': '元大臺灣ESG永續', '00701.TW': '國泰股利精選30'}
        self.BlueChip = {'2356.TW': '英業達','2330.TW': '台積電', '2454.TW': '聯發科', '2357.TW': '華碩', '6505.TW': '台塑','1102.TW': '亞泥','2324.TW': '仁寶'}
        self.Financial = {'2890.TW': '永豐金','2886.TW': '兆豐金','2885.TW': '元大金'}

    def getDate():
        DateTime = datetime.datetime.now()
        Y1 = DateTime.year
        M1 = DateTime.month
        D1 = DateTime.day
        dateStr = str(Y1)+"/"+str(M1)+"/"+str(D1)    
        return dateStr
    
    # def getData(self, symbols):
    #     for symbol, name in symbols.items():