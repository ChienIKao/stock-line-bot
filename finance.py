# -*- coding: Big5 -*-
###
#�o�q�{���X�b�p�⧹���ʥ����u�MRSI���Ы�A�ϥ�rolling��ƭp��F20�骺�зǮt�A�îھڼзǮt�p��X�F���L�q�D���W���M�U���C�M��A�{���X�ϥ�loc��Ƨ�X��ѻ��C�󥬪L�q�D�U���BRSI���Фp��30�ɪ�����A�ñNsignal���]�m��1�C�o��ܦb�o�Ǥ���A�ѻ��w�g�^�}�F���L�q�D�U���A�BRSI������ܪѲ��w�g�W��A�o�i��O�@�ӶR�J�I�C�z�i�H�ھڳo�ǫH���i��i�@�B�����R�M�M���C
#�[�Jline�s�ժ��o�e���O:
###

import numpy as np
import pandas as pd
import yfinance as yf
import matplotlib.pyplot as plt
import datetime
from datetime import date, timedelta
##�s�W:�N�Ͷչ϶ǰe��Line�s��
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




#from requests_toolbelt.multipart.encoder import MultipartEncoder                #pip install requests_toolbelt
#pip install xlsxwriter

# �]�wLine Notify�A�Ȫ��s���v��:  �i�ѥ�etf�s�աj
access_token = 'optTVwv7UKvM7QEXFFTZNi9YB4yZZ1WojsWL6hxO2iVw2lCS7fAA/Uf9ztVqJXDb4xnXZdwuBZHgFK+eXz/bE86B8Ge+YBtEt6mEduMjFf5m7Ljqc7LdoDW1ss5BfC/iLJPbS+1UFv8yvk4zTCka1QdB04t89/1O/w1cDnyilFU=' # ��token
#access_token = 'OPkJV3'

# �]�w�Ѳ��N��
#symbol = '00878.TW'
# �]�w�Ѳ��N���C��
#symbols = ['00878.TW', '0056.TW', '00713.TW', '00919.TW']
# �]�w�Ѳ��N���r��

#���`��ETF
#symbols = {'00878.TW': '������򰪪Ѯ�', '00919.TW': '�s�q�x�W��ﰪ��', '00713.TW': '���j�x�W�����C�i', '0056.TW': '���j���Ѯ�','006208.TW': '�I���x�W50', '00690.TW': '�������w30', '00850.TW': '���j�O�WESG����', '00701.TW': '����ѧQ���30'}
#���`���������Z�u��
symbols = {'2356.TW': '�^�~�F','2330.TW': '�x�n�q', '2454.TW': '�p�o��', '2357.TW': '�غ�', '6505.TW': '�x��','1102.TW': '�Ȫd','2324.TW': '���_'}
#���`�����Ī�
#symbols = {'2890.TW': '���ת�','2886.TW': '���ת�','2885.TW': '���j��'}


#symbols = {'00900.TW': '���Ѯ�', '00888.TW': '����', '00692.TW': '����'}
#symbols = {'00878.TW': '������򰪪Ѯ�', '00919.TW': '�s�q�x�W��ﰪ��', '00713.TW': '���j�x�W�����C�i', '0056.TW': '���j���Ѯ�'}
#symbols = {'00878.TW': '������򰪪Ѯ�', '00919.TW': '�s�q�x�W��ﰪ��'}
#symbols = {'2330.TW': '�x�n�q', '2454.TW': '�p�o��', '2357.TW': '�غ�', '6505.TW': '�x��'}
#symbols = {'006208.TW': '�I���x�W50', '00690.TW': '�������w30', '00850.TW': '���j�O�WESG����', '00701.TW': '����ѧQ���30'}
#symbols = {'2324.TW': '���_', '2357.TW': '�غ�', '1101.TW': '�x�d', '2356.TW': '�^�~�F'}
#symbols = {'1102.TW': '�Ȫd', '2886.TW': '���ת�', '2890.TW': '���ת�', '2885.TW': '���j��'}
#symbols = {'1102.TW': '�Ȫd', '2890.TW': '���ת�'}
#symbols = {'00919.TW': '�s�q�x�W��ﰪ��'}
#symbols = {'00713.TW': '���j�x�W�����C�i'}
#symbols = {'0056.TW': '���j���Ѯ�'}
#symbols = {'00900.TW': '�I�����Ѯ�'}
#symbols = {'2890.TW': '���ת�'}
#symbols = {'2356.TW': '�^�~�F'}
#symbols = {'0050.TW': '���j�x�W50'}
#symbols = {'00690.TW': '�������w30'}

'''
#��k1: �ϥΥx�W�Ҩ����Ҫ��Ѳ��N���ഫAPI�A�N�Ѳ��N���ഫ������W��
url = 'https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch=' + symbol
try:
    response = requests.get(url)
    # �Nresponse���e�ରutf-8
    response.encoding = 'utf-8'
    name = response.json()['msgArray'][0]['n']
except requests.exceptions.JSONDecodeError:
    name = ""

#��k2: �ϥΥx�W�Ҩ����Ҫ��Ѳ��N���d�ߺ����A����Ѳ�����W��
url = 'http://mis.twse.com.tw/stock/api/getStockInfo.jsp?c=' + symbol
response = requests.get(url)
pattern = re.compile(symbol + r'\s+(.*?)\s+')
match = pattern.search(response.text)
if match:
    symbol_name = match.group(1)
else:
    symbol_name = symbol      
#�o�ӿ��~�q�`�O�ѩ�API��^���ƾڮ榡�����T�εL�k�ѪR�Ӥް_���C�z�i�H���ըϥ�try-except�y�y����JSONDecodeError���`�A�æb����첧�`�ɳB�z���C
'''

#��X������
tonow = datetime.datetime.now()
Y1 = tonow.year
M1 = tonow.month
D1 = tonow.day
todaystr = str(Y1)+"/"+str(M1)+"/"+str(D1)

#for symbol in symbols:  #####################################�s��̦��d��
'''
# �إ�Excel�ɮ�
# �إߤ@�ӪŪ�Excel�ɮ�
#writer = pd.ExcelWriter('stock_data.xlsx', engine='xlsxwriter')
# �ˬdExcel�ɮ׬O�_�s�b
if os.path.isfile('stock_data.xlsx'):
    # �p�G�s�b�A�N�}�ҳo���ɮ�
    ccdf = pd.read_excel('stock_data.xlsx', sheet_name=symbols[symbol])
    startrow = ccdf.shape[0]
    writer = pd.ExcelWriter('stock_data.xlsx', engine='openpyxl')
    writer.book = openpyxl.load_workbook('stock_data.xlsx')
else:
    # �p�G���s�b�A�N�s�ؤ@���ɮ�
    startrow = 0
    writer = pd.ExcelWriter('stock_data.xlsx', engine='xlsxwriter')
'''
###-----------------------------------------------�H�U�}�l�̭ӪѶ]�^��
for symbol, name in symbols.items():

    # �ˬd�ɮ׬O�_�s�b
    if os.path.isfile('999.xlsx'):
        # �p�G�ɮצs�b�A���J��
        wb = load_workbook('999.xlsx')
        # �s��S�w����ƪ�
        if name in wb.sheetnames:
            ws = wb[name]
            # �b�o�̶i���ƪ��s��
        else:
            # �p�G�S�w����ƪ��s�b�A�h�إߤ@�ӷs����ƪ�
            ws = wb.create_sheet(name)
            # �b�o�̶i��s��ƪ���l��
            listtitle=['�N��','�ѦW','���','���L��(��)','����q','RSI','MFI','ADL','Williams[20]','���L�W��(��)','���L�U��(��)','�맡�u(��)','�u���u(��)','�J���ѻ�[45](��)','�ާ@��ĳ']
            ws.append(listtitle)
    else:
        # �p�G�ɮפ��s�b�A�إߤ@�ӷs��xlsx�ɮ�
        wb = Workbook()
        # �إ߯S�w����ƪ�
        ws = wb.create_sheet(name)
        # �b�o�̶i��s��ƪ���l��
        listtitle=['�N��','�ѦW','���','���L��(��)','����q','RSI','MFI','ADL','Williams[20]','���L�W��(��)','���L�U��(��)','�맡�u(��)','�u���u(��)','�J���ѻ�[45](��)','�ާ@��ĳ']
        ws.append(listtitle)

    todaydate = ">>"+todaystr+"�����Ͷ�>>"+"\n"+" ["+str(symbol)+"]-[ "+name+" ] ���R�K�n:"


    # �U���ѻ����-�̪�1�~
    df = yf.download(symbol, period='1y')
    #print(df)
    #df���榡: [  Date       Open       High        Low      Close  Adj Close     Volume   ]
    #sys.exit(1)

    # �p�Ⲿ�ʥ����u
    df['MA20'] = df['Close'].rolling(window=20).mean()
    df['MA60'] = df['Close'].rolling(window=60).mean()

    # �q��Ƥ������̰����]high�^�M�̧C���]low�^ 
    high_values = df['High'].values 
    low_values = df['Low'].values

    # �p��william���� 
    #df['WILLIAMS'] = (high_values - low_values.shift()) / low_values.shift() 
    # �N low_values �ഫ�� Pandas �� Series
    #low_series = pd.Series(low_values)
    # �ϥ� shift() ��k
    #df['WILLIAMS'] = (high_values - low_series.shift()) / low_series.shift()
    #�ϥγ�·�骺WILLIAM���ƪ��p�⤽�����G(�̰��� - ���L��) / �̰��� - �̧C��) * -100
    #df['WILLIAMS'] = ((df['High'] - df['Close']) / (df['High'] - df['Low'])) * -100
    #�ϥιL�h20�Ѷg�����p��覡:WILLIAM���ƪ��p�⤽�����G(�L�h45�Ѧ��L�̰���-��馬�L��) / �L�h45�ѳ̰��� - �L�h45�ѳ̧C��) * -100
    window_size = 20
    high_values = df['High'].rolling(window_size).max()
    low_values = df['Low'].rolling(window_size).min()
    df['WILLIAMS'] = ((high_values - df['Close'][-1]) / (high_values - low_values)) * 100


    # �p��MFI����y�V����--------------------------------------------------------
    #����MFI���ЩMRSI���Ы������A�����h�Ҷq�F����q���]���A�Ӥ���RSI�u�Ҷq����]���C
    #����y�q���СB����y�V����(Money Flow Index , MFI)�A�O�޳N���R���ΨӿŶq�R�����O���u��A���PRSI�ۦ����]�t�F����q�A�Ӥ���RSI�u�]�A����C
    #MFI�Q�{�����@�w������N�q�A�q�`�Ψӹw���������աA�z�LMFI�[��X���W�R�B�W��T���A�i�H�����H�ѧO��b���f��C
    #MFI����80�ΥH�W�n�`�N���઺�ͶաB�����b�j�l���U�^�ͶդU�~��U�^�ɡAMFI�ȥi��C��20�A�h�N��W�檺���p�C
    #�p�G���樫�թMMFI�����ܬۤϡA�N�|�X�{�I���T���A�ӥ��Ѫi�ʪ��ݺ��ݶ^�]�Ψӹw���R�J�ν�X�����|�C
    #MFI����y�V���Ъ��p�⤽�����G100 - (100 / (1 + money_flow_ratio))
    #close_values = df['Close'].values 
    #df['FMI'] = (high_values + low_values + close_values) / 3
    #�ϥβ�?/��??�]Accumulation/Distribution Line�AADL�^??��CADL�O�@����_����q�Mɲ��?�ƪ���?��?�A�Τ_�Ŷq??��???�O�CMFI�O��_ADL?�⪺�A�]���]�Q??MFI/ADL��?�C
    # ?��ADL
    adl = ((df['Close'] - df['Low']) - (df['High'] - df['Close'])) / (df['High'] - df['Low']) * df['Volume']
    adl = adl.cumsum()

    #df['Volume'] = df['Volume'].fillna(0)
    # �p�� MFI
    typical_price = (df['High'] + df['Low'] + df['Close']) / 3
    money_flow = typical_price * df['Volume']
    positive_flow = np.where(typical_price > typical_price.shift(1), money_flow, 0)
    negative_flow = np.where(typical_price < typical_price.shift(1), money_flow, 0)
    positive_flow_sum = pd.Series(positive_flow).rolling(window=14).sum()
    negative_flow_sum = pd.Series(negative_flow).rolling(window=14).sum()
    money_flow_ratio = np.where(negative_flow_sum == 0, 0.5, positive_flow_sum / negative_flow_sum)
    mfi = 100 - (100 / (1 + money_flow_ratio))
    # ��
    df['MFI'] = mfi
    df['ADL'] = adl
    #�ˬdMFI���ЬO�_�p�⦨�\
    print(df['MFI'])
    print(df['ADL'])
    #----------------------------------------------------------------------------------------
    #sys.exit(1)

    # �p��RSI����
    delta = df['Close'].diff()
    gain = delta.where(delta > 0, 0)
    loss = -delta.where(delta < 0, 0)
    avg_gain = gain.rolling(window=14).mean()
    avg_loss = loss.rolling(window=14).mean()
    rs = avg_gain / avg_loss
    df['RSI'] = 100 - (100 / (1 + rs))
    #print(df['RSI'][-1])

    # �p�⥬�L�q�D
    df['std'] = df['Close'].rolling(window=20).std()
    df['upper'] = df['MA20'] + 2 * df['std']
    df['lower'] = df['MA20'] - 2 * df['std']

    #��ʳ]�w�R�J����M���l����
    #buy_price = 17.5
    #stop_loss_price = 18.5

    '''
    # �Q�γ̪�@�Ѫ����L���A�۰ʭp��̨ζR�J����M���l����
    df['diff'] = df['upper'] - df['lower']
    df['buy_price'] = df['lower'] + 0.2 * df['diff']
    df['stop_loss_price'] = df['lower'] + 0.1 * df['diff']
    #print('��ĳ���̨ζR�J�I:' + str(df['buy_price']))
    '''
    # �Q�Ϋe�@�Ӥ몺���L�q�D�W�U���ӭp��̨ζR�J����M���l����
    df['diff'] = df['upper'] - df['lower']
    df['buy_price'] = df['lower'].rolling(window=45).mean() + 0.2 * df['diff'].rolling(window=45).mean()
    #�o�q�{���X�O�Ψӭp��̨ζR�J���檺�����A�䤤 df['lower'].rolling(window=20).mean() �O���p��L�h45�Ѫ����L�q�D�U�������ȡA�� 
    #df['diff'].rolling(window=20).mean() �O���p��L�h45�Ѫ����L�q�D�W�U���϶��t�������ȡC�]���A�o�q�{���X���N��O�G�Q�ιL�h45�Ѫ����L�q�D�U�������ȥ[�W�L�h45�Ѫ����L�q�D�W�U���϶��t�������ȭ��W0.2�A�ӭp��̨ζR�J����C
    #�o�Ӥ������ت��O�n�����H�b�ѻ��B��C�I�ɶR�i�A�åB�]�w�@�Ӱ��l����A�H����I�C�o�Ӥ������֤߬O�Q�Υ��L�q�D���U���ӧP�_�ѻ��O�_�B��C�I�A�A�[�W�@�ӽw�İ϶��A�H�קK�R�i�ɪѻ��w�g�ϼu�C�ӳo�ӽw�İ϶����j�p�A�N�O��0.2�������L�q�D�W�U���϶��t�ӨM�w���C
    df['stop_loss_price'] = df['lower'].rolling(window=45).mean() + 0.1 * df['diff'].rolling(window=45).mean()

    # ��X�̨ζR�J�I
    df['signal'] = 0
    #df.loc[(df['Close'] < df['lower']) & (df['RSI'] < 30) & (df['Close'] <= buy_price), 'signal'] = 1
    #df.loc[(df['Close'] < df['lower']) & (df['RSI'] < 30) & (df['Close'] <= df['buy_price']), 'signal'] = 1
    df.loc[(df['Close'] < df['lower'] * 1.05) & (df['RSI'] < 30) & (df['Close'] <= df['buy_price']), 'signal'] = 1
    #�o�q�{���X�O�ΨӧP�_�O�_���R�J�T��������A�䤤�]�A�ѻ��C�󥬪L�q�D�U�y�u�]lower�^��1.05���BRSI���Фp��30�B�B�ѻ��p�󵥩�R�J����]buy_price�^�C�p�G�ŦX�o�Ǳ���A�h�|�b��������ƦC���аOsignal��1�A��ܦ��R�J�T���C
    #�ܩ�p��M�w�̨ζR�J�I�A�o�q�`�ݭn�Ҽ{�h�ئ]���A�Ҧp�޳N���СB�򥻭����R�B�����Ͷյ����C�o�Ǧ]���i�H�ھڭӤH����굦���M���I���n�i���X�Ҽ{�A�H��X�̨Ϊ��R�J�I�C��ĳ�b�i����e�A���i��R������s�M���R�A�è�w�X���T����굦���M���I����I�C

    '''
    # �P�_�O�_�ӶR�i�B��X���[��
    last_close = df['Close'][-1]
    if df['signal'][-1] == 1:
        print(symbol+f'�̪�@�Ѧ��L���w�C�󥬪L�U���BRSI���Фp��30�A��ĳ�R�i�A�R�J���欰{last_close:.2f}��')
        print(symbol+f'��ĳ�]�w���l���欰{stop_loss_price:.2f}��')
    elif last_close > df['MA20'][-1] and last_close > df['upper'][-1]:
        print(symbol+f'�̪�@�Ѧ��L���w���󤤽u�B���񥬪L�W���A��ĳ��X�A��X���欰{last_close:.2f}��')
    else:
        print(symbol+'�̪�@�Ѧ��L���b���L�q�D���A��ĳ�[��C')
    '''
    # �P�_�O�_�ӶR�i�B��X���[��
    last_close = df['Close'][-1]
    #print('last_close:'+str(last_close))
    last_rsi = df['RSI'][-1]
    #print('last_rsi:'+str(last_rsi))
    last_buy_price = df['buy_price'][-1]
    print('�Q�γ̪�45�ѥ���ƾڱ��⪺�̨ζR�I:'+ str(last_buy_price))
    last_stop_loss_price = df['stop_loss_price'][-1]
    print('���l���欰:'+str(last_stop_loss_price))

    if df['signal'][-1] == 1:
        keyword = symbol+f'�̪�@�Ѧ��L��[{last_close:.2f}��]�A�w�g�X�{�R�J�T���F!'+'\n'+'**�R�J�T���]�w����: �ѻ��w�C�󥬪L�q�D�U�y�u��1.05���B�BRSI���Фp��30�B�B�ѻ��C����⪺�X�z�R�J����C'+'\n'+'��ĳ���ֶR�i�A�R�J�Ѧһ��欰'+str(round(last_buy_price*0.995,2))+'���C(���o�j�]~)'+'\n'+'�t���ѰѦҰ��l���欰'+str(round(last_stop_loss_price,2))+'���C'    
    elif last_close >= df['MA20'][-1] and last_close >= df['MA60'][-1] and last_close >= df['upper'][-1]*0.995:
        keyword = symbol+f'�̪�@�Ѧ��L��[{last_close:.2f}��]�A�w����MA20�맡�u��MA60�u���u�B�W�L���L�W���C'+'\n'+'�j�P��ĳ�i��X�C'+'\n'+'��X�Ѧһ���:'+str(round(last_close*1.005,2))+'���C(�o�j�]�F~)'+'\n'+'RSI���Ь�'+str(round(last_rsi,2))   
    elif last_close >= df['MA20'][-1] and last_close >= df['upper'][-1]*0.995:
        keyword = symbol+f'�̪�@�Ѧ��L��[{last_close:.2f}��]�A�w����MA20�맡�u�B�W�L���L�W���C'+'\n'+'�Y����X�p�e�A��ĳ�i��X�C'+'\n'+'��X�Ѧһ��欰'+str(round(last_close*1.005,2))+'���C'+'\n'+'RSI���Ь�'+str(round(last_rsi,2)) 
    elif last_close <= df['MA20'][-1] and last_close <= df['MA60'][-1] and last_close <= df['lower'][-1]*1.005:
        keyword = symbol+f'�̪�@�Ѧ��L��[{last_close:.2f}��]�A�w�C��MA20�맡�u��MA60�u���u�B�C�󥬪L�U���C'+'\n'+'�j�P��ĳ�i�[�X�R�i!!! (���z�o�]~)'+'\n'+'RSI���Ь�'+str(round(last_rsi,2))    
    elif last_close <= df['MA20'][-1] and last_close <= df['lower'][-1]*1.005:
        keyword = symbol+f'�̪�@�Ѧ��L��[{last_close:.2f}��]�A�w�C��MA20�맡�u�B�C�󥬪L�U���C'+'\n'+'�Y���n�[�X�R�i���p�e�A��ĳ�i�H�ǳƤF��C (���z�o�]~)'+'\n'+'RSI���Ь�'+str(round(last_rsi,2))    
    #elif last_close >= df['upper'][-1] * 0.95:
    #    keyword = symbol+f'�̪�@�Ѧ��L��[{last_close:.2f}��]�A�w���󥬪L�W��0.95���w�İ϶��C'+'\n'+'�Y���n��X���p�e�A��ĳ�i�H�ǳƤF�C'+'\n'+'RSI���Ь�'+str(round(last_rsi,2))    
    else:
        keyword = symbol+f'�̪�@�Ѧ��L��[{last_close:.2f}��]�A�٦b���L�q�D���C'+'\n'+'��ĳ�����[��~'+'\n'+'RSI���Ь�'+str(round(last_rsi,2))
        

    #RS���ЬO�@�ا޳N���R���СA�Ω�Ŷq�Ѳ����W�R�M�W��{�סCRSI���Ъ��p��覡�O�ھڤ@�q�ɶ����Ѳ����^�T�ת������ȡA�ӭp��Ѳ����j�z�{�סCRSI���Ъ����Ƚd��0��100�A�@��{��RSI���Цb30�H�U��ܪѲ��B��W�檬�A�A�Ӧb70�H�W��ܪѲ��B��W�R���A�C
    #�b�{���X����ĳRSI���Фp��30�ɶR�i�Ѳ��A�O�]����RSI���Фp��30�ɡA��ܪѲ��B��W�檬�A�A���������i��w�g�L�״d�[�A�Ѳ�����i��w�g�Q�C���C���ɡA�p�G�Ѳ����򥻭��S���o�ͭ��j�ܤơA�i��O�@�ӶR�J���n�ɾ��C��M�A�o�u�O�@�ӰѦҡA�����ٻݭn�ھڦۤv�����I�Ө���O�M���ؼСA���X������M�ǽT���M���C
    '''
    ####�ǳƱN���G�s�JEXCEL��
    # �N�C�ӪѲ��N������Ƽg�JExcel�ɮ�
    # Ū���{������ƪ�A�p�G���s�b�N�إߤ@�ӷs��
    '''
    #�|�إߤ@�ӦW��stock_data.xlsx��Excel�ɮסA�óv�@�B�zsymbols�r�夤���C�ӪѲ��N���C
    #���C�ӪѲ��A���|�U���̪�1�~���ѻ���ơA�p�Ⲿ�ʥ����u�BMFI�MADL�A�M��N��Ƽg�JExcel�ɮפ��C�C�ӪѲ�����ƪ�W�ٷ|�H�Ѳ��W�٩R�W�A���榡�]�|�]�w�n�C�̫�A���|�x�sExcel�ɮסC

    
    # �N��Ƽg�JExcel�ɮ�
    #�U���|�g�J�Ҧ������
    #df.to_excel(writer, sheet_name=sheet_name, index=False)
    # �N�Ҧ���Ƽg�JExcel�ɮ�
    #df.to_excel(writer, sheet_name=sheet_name, startrow=startrow, index=False)    

    # �s�W���榡
    '''
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    number_format = workbook.add_format({'num_format': '#,##0.00'})
    worksheet.set_column('A:A', 12, date_format)
    worksheet.set_column('B:B', None, number_format)
    worksheet.set_column('C:C', None, number_format)
    worksheet.set_column('D:D', None, number_format)
    worksheet.set_column('E:E', None, number_format)
    worksheet.set_column('F:F', None, number_format)
    worksheet.set_column('G:G', None, number_format)
    worksheet.set_column('H:H', None, number_format)
    # ���o�̪�@�Ѫ����
    latest_data = df.iloc[-1:]

    #�N�̪�@�Ѫ���Ƽg�JExcel�ɮ�
    latest_data.to_excel(writer, sheet_name=symbol, startrow=startrow+1, index=False, header=True)
    # �N��Ƽg�JExcel�ɮ�
    #df.to_excel(writer, sheet_name=symbol, startrow=startrow, index=False)
    '''
    # �b�o�̶i��excel��L����ƪ�s��
    # �v�@�ˬd�C�� C ��쪺�ȤF�C�q�ĤG�C�}�l�A�q�ĤT���ĤT��]�]�N�O C ���^�v�@���o�C���x�s��
    c2_value = ""
    for row in ws.iter_rows(min_row=2, min_col=3, max_col=3):
        for cell in row:
            if cell.value == todaystr:
                # �p�G C ��쪺�ȵ��� 'EEE'�A����o�̪��{���X��
                c2_value = 'YES'
                break
            else:
                # �p�G C ��쪺�Ȥ����� 'EEE'�A����o�̪��{���X��
                c2_value = 'NO'

    
    # �P�_ C2 �x�s�檺�ȬO�_���� 'todaystr'   ��ܤ��Ѫ���Ƥw�g�g�L�F
    '''
    if c2_value is None:
        # �p�G������ todaystr�A��ܸӤ������Ƥ��s�b�A�h�~��s�W�@�C
        # �b�ĤG�C���J�@��A�û��g�J�Ĥ@�C
        #ws.insert_rows(1)
        #listtitle=['�N��','�ѦW','���','���L��(��)','����q','RSI','MFI','ADL','Williams[20]','���L�W��(��)','���L�U��(��)','�맡�u(��)','�u���u(��)','�J���ѻ�[45](��)','�ާ@��ĳ']
        listdatas=[symbol , name , todaystr , str(round(last_close,2)) , str(round(df['Volume'][-1],2)) , str(round(last_rsi,2)) , str(round(df['MFI'][-1],2)) , str(round(df['ADL'][-1]/10000,2)) , str(round(df['WILLIAMS'][-1],2))+'%' , str(round(df['upper'][-1],2)) , str(round(df['lower'][-1],2)) , str(round(df['MA20'][-1],2)) , str(round(df['MA60'][-1],2)) , str(round(last_buy_price,2)) , keyword]
        # �N��Ƽg�J�ĤG�C
        ws.append(listdatas)    
    '''
    if c2_value == 'YES':
        # �p�G���� todaystr�A��ܭL�����Ƥw�s�b�A�h�έק諸�覡�мg���
        # �b�ĤG�C���J�@��A�û��g�J�Ĥ@�C
        # �ק�ĤG�C�����
        ws.cell(row=2, column=1, value=symbol)
        ws.cell(row=2, column=2, value=name)
        ws.cell(row=2, column=3, value=todaystr)
        ws.cell(row=2, column=4, value=str(round(last_close,2)))
        ws.cell(row=2, column=5, value=str(round(df['Volume'][-1],2)))
        ws.cell(row=2, column=6, value=str(round(last_rsi,2)))
        ws.cell(row=2, column=7, value=str(round(df['MFI'][-1],2)))
        ws.cell(row=2, column=8, value=str(round(df['ADL'][-1]/10000,2)))
        ws.cell(row=2, column=9, value=str(round(df['WILLIAMS'][-1],2))+'%')
        ws.cell(row=2, column=10, value=str(round(df['upper'][-1],2)))
        ws.cell(row=2, column=11, value=str(round(df['lower'][-1],2)))
        ws.cell(row=2, column=12, value=str(round(df['MA20'][-1],2)))
        ws.cell(row=2, column=13, value=str(round(df['MA60'][-1],2)))
        ws.cell(row=2, column=14, value=str(round(last_buy_price,2)))
        ws.cell(row=2, column=15, value=keyword)
    
    else:
        # �p�G������ todaystr�A��ܸӤ������Ƥ��s�b�A�h�~��s�W�@�C
        # �b�ĤG�C���J�@��A�û��g�J�Ĥ@�C
        #ws.insert_rows(1)
        #listtitle=['�N��','�ѦW','���','���L��(��)','����q','RSI','MFI','ADL','Williams[20]','���L�W��(��)','���L�U��(��)','�맡�u(��)','�u���u(��)','�J���ѻ�[45](��)','�ާ@��ĳ']
        listdatas=[symbol , name , todaystr , str(round(last_close,2)) , str(round(df['Volume'][-1],2)) , str(round(last_rsi,2)) , str(round(df['MFI'][-1],2)) , str(round(df['ADL'][-1]/10000,2)) , str(round(df['WILLIAMS'][-1],2))+'%' , str(round(df['upper'][-1],2)) , str(round(df['lower'][-1],2)) , str(round(df['MA20'][-1],2)) , str(round(df['MA60'][-1],2)) , str(round(last_buy_price,2)) , keyword]
        # �N��Ƽg�J�ĤG�C
        ws.append(listdatas)

    # �x�s�ɮ�
    #wb.save('999.xlsx')


    ####��Xline�]��-----------------------------------------------------------------------------------
    msgsouce = "��45�ѥ���ƾڪ��̨ζR�I: "+ str(round(last_buy_price,2)) + "��"+"\n"
    msgsouce = msgsouce + "*���L�q�D�W��: " + str(round(df['upper'][-1],2)) + "\n"
    msgsouce = msgsouce + "*�맡�u�Ѧҭ�: " + str(round(df['MA20'][-1],2)) + "\n"
    msgsouce = msgsouce + "*�u���u�Ѧҭ�: " + str(round(df['MA60'][-1],2)) + "\n"
    msgsouce = msgsouce + "*���L�q�D�U��: " + str(round(df['lower'][-1],2)) + "\n"+ "\n"
    msgsouce = msgsouce + "*Williams[20]:    " + str(round(df['WILLIAMS'][-1],2)) +"%"+"\n"
    msgsouce = msgsouce + "*MFI����y�V����:  " + str(round(df['MFI'][-1],2))+"\n"
    msgsouce = msgsouce + "*A/D Line����:  " + str(round(df['ADL'][-1]/10000,2))    
    msg = "\n"+ todaydate +"\n"+ str(msgsouce)
    r00 = lineTool.lineNotify(access_token,msg)


    # ø�s�Ͷչ�
    fig, ax = plt.subplots(figsize=(16, 9))

    ax.plot(df.index, df['Close'], lw=3,marker='.',color='#FF0000', label='C_Price')
    ax.plot(df.index, df['MA20'], lw=3,color='#EE7700', linestyle='-', label='MA20 day')
    ax.plot(df.index, df['MA60'], lw=3,color='#7700BB', linestyle='-', label='MA60 day')
    ax.plot(df.index, df['upper'], color='k', linestyle='--', label='Bollinger Bands_Upper')
    ax.plot(df.index, df['lower'], color='k', linestyle='--', label='Bollinger Bands_Lower')

    ax.fill_between(df.index, df['upper'], df['lower'], alpha=0.2)
    #�U���O�аO�̪�@��'�����L��:
    #ax.axhline(df['Close'][-1], color='b', linestyle='--', label='Last Close')
    #ax.axhline(df['upper'][-1], color='r', linestyle='--', label='Upper')
    #ax.axhline(df['lower'][-1], color='g', linestyle='--', label='Lower')

    # ��̪ܳ�@�Ѧ��L����RSI��
    ax2 = ax.twinx()
    ax2.plot(df.index, df['RSI'], lw=1,color='b', label='RSI')
    ax2.axhline(y=30, lw=1,color='g', linestyle='--')
    ax2.axhline(y=70, lw=1,color='g', linestyle='--')
    ax2.set_ylim([0, 100])
    ax2.set_ylabel('RSI')
    ax2.legend(loc='upper left')

    ax.legend()
    ax.set_title(symbol+'~'+todaystr)
    ax.set_xlabel('Date')
    ax.set_ylabel('Price')

    # Ū���Ͷչ��ഫ���N��O�s��@��BytesIO��H��
    '''
    fig, ax9 = plt.subplots()
    ax9.plot(df.index, df['Close'])
    ax9.set_title('Stock Price')
    ax9.set_xlabel('Date')
    ax9.set_ylabel('Price')
    '''
    buffer = BytesIO()
    plt.savefig(buffer, format='png')
    buffer.seek(0)
    #image = Image.open(buffer)   

    #plt.show()


    ###
    #�o�q�{���X�ϥΤFmatplotlib�w��ø�s�ͶչϡC�{���X�|ø�s�ѻ��B20��M60�骺���ʥ����u�A�H�Υ��L�q�D���W���M�U���C�{���X�٨ϥ�fill_between��ƶ�R�F���L�q�D���ϰ�C�̫�A�{���X��ܤFø�s�n���ͶչϡC
    ###
    '''
    # �`���̪�@�Ѫ����L��
    last_close = df['Close'][-1]
    # �P�_�O�_�ӶR�i�B��X���[��
    if last_close < df['MA20'][-1] and last_close < df['lower'][-1]:
        print(symbol+'�̪�@�Ѧ��L���w�C�󤤽u�B���񥬪L�U���A��ĳ�R�i�C')
    elif last_close > df['MA20'][-1] and last_close > df['upper'][-1]:
        print(symbol+'�̪�@�Ѧ��L���w���󤤽u�B���񥬪L�W���A��ĳ��X�C')
    else:
        print(symbol+'�̪�@�Ѧ��L���b���L�q�D���A��ĳ�[��C')
    '''



    # �N�Ͷչ��ഫ���G�i��ƾ�
    image_binary = buffer.getvalue()

    # �ǰe�ͶչϨ�Line�s��
    url = 'https://notify-api.line.me/api/notify'
    headers = {'Authorization': 'Bearer ' + access_token}
            #data = {'message': 'Stock Price', 'imageFile': ('image.png', buffer, 'image/png')}
            #response = requests.post(url, headers=headers, files=data)
    data = {'message': "\n"+keyword}
    files = {'imageFile': ('image.png', image_binary, 'image/png')}
    response = requests.post(url, headers=headers, data=data, files=files)     #�̥��T�y�k

    # �ˬdHTTP�T��
    print(response.status_code)
    print(response.text)

    #�ڭ̭���Ū���ͶչϨñN��O�s��@��BytesIO��H���C�M��A�ڭ̨ϥ�BytesIO��H��getvalue()
    #��k�N�Ͷչ��ഫ���G�i��ƾڡC���U�ӡA�ڭ̨ϥ�HTTP�ШD�VLine Notify API�o�e�����A�]�A�������D�M�ͶչϡC�bHTTP�ШD���A�ڭ̱N�ͶչϪ��G�i��ƾڧ@�����ǰe�A�ñN�������D�@�����ƾڶǰe�C�̫�A�ڭ̥i�H�ˬdHTTP�T���H�T�w�����O�_���\�o�e�C

    #�b�W�����{���X���A�ڭ̭����]�wLine Notify�A�Ȫ��s���v���A�M��Ū���ͶչϨñN���ഫ��PNG�榡�C���U�ӡA�ڭ̨ϥ�HTTP�ШD�VLine�o�e�q�������A�]�A�������D�M�ͶչϡC�̫�A�ڭ̥i�H�ˬdHTTP�T���H�T�w�����O�_���\�o�e�C
    #�Ъ`�N�ALine Notify�A�Ȧ��@�ǭ���A�Ҧp�C�ѳ̦h�u��o�e1000�������A�C�Ӯ����̤j�u��]�t1MB���ƾڵ��C�p�G�z�ݭn�o�e��h�������Χ�j���ƾڡA�ЦҼ{�ϥΨ�L�������A�ȡA�ҦpTelegram Bot���C
    msg_end = todaystr+"///////�o�]"
    r99 = lineTool.lineNotify(access_token,msg_end)

    #�N���ɼg�JEXCEL��
    '''
    # ���R���u�@������w�g�s�b���Ϥ��A�קK�ɮ׹L�j
    for imga in ws._images:
        if imga.anchor == 'A15':
            del ws._images[imga]
    '''
    # �R���u�@�����Ҧ��Ϥ�
    ws._images = []
    
    # �N�Ϥ��g�J�u�@��
    imgg = XLImage(BytesIO(image_binary))
    ws.add_image(imgg, 'A15')
    # �x�sEXCEL�ɮ�
    wb.save('999.xlsx')

    time.sleep(5)  # �Ȱ�5����




