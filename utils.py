from pandas import Series, DataFrame
import pandas as pd
import locale
import os
import time
import win32com.client


# PLUS 공통 OBJECT
g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')





'''
# 연결 여부 체크
'''
def connect_status() :
    objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
    bConnect = objCpCybos.IsConnect
    if (bConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음. ")
        exit()


'''
종목명 / 종목코드 구하기
'''
def kospi_kosdaq_dict() :
    objCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
    codeList = objCpCodeMgr.GetStockListByMarket(1)  # 거래소
    codeList2 = objCpCodeMgr.GetStockListByMarket(2)  # 코스닥
    kospi_name_dict = {objCpCodeMgr.CodeToName(v) : v  for v in codeList}
    kosdaq_name_dict = {objCpCodeMgr.CodeToName(v) : v for v in codeList2}
    print('코스피 전체 종목 : {}'.format(len(codeList)))
    print('코스닥 전체 종목 : {}'.format(len(codeList2)))
    return kospi_name_dict, kosdaq_name_dict

