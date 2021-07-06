import win32com.client
import os, sys, ctypes
import pandas as pd
from datetime import datetime
import time, calendar
from bs4 import BeautifulSoup
from urllib.request import urlopen
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from utils import *


# 크레온 플러스 공통 OBJECT
# 각종 코드정보 및 코드 리스트를 얻을 수 있습니다.
cpStockCd = win32com.client.Dispatch('CpUtil.CpStockCode')
# CYBOS의 각종 상태를 확인할 수 있음.
cpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
# 설명 : 주문 오브젝트를 사용하기 위해 필요한 초기화 과정들을 수행한다
# 모든 주문오브젝트는 사용하기 전에, 필수적으로 TradeInit을 호출한 후에 사용할 수 있다.
# 전역변수(글로벌 변수) 로 선언하여 사용하여야 합니다.
cpTradeUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')
# 주식 종목의 현재가에 관련된 데이터
cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
# 주식, 업종, ELW의 차트데이터를 수신합니다.
cpOhlc = win32com.client.Dispatch('CpSysDib.StockChart')
# 계좌별 잔고 및 주문체결 평가 현황 데이터를 요청하고 수신한다
cpBalance = win32com.client.Dispatch('CpTrade.CpTd6033')
# 계좌별 매수주문 가능 금액/수량 데이터를 요청하고 수신한다
cpCash = win32com.client.Dispatch('CpTrade.CpTdNew5331A')
# 장내주식/코스닥주식/ELW 주문(현금 주문) 데이터를 요청하고 수신한다
cpOrder = win32com.client.Dispatch('CpTrade.CpTd0311')
# CYBOS에서 사용되는 주식코드 조회 작업을 함.
cpCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')

# 종목코드 리스트 구하기

def code_name() :
    codeList = cpCodeMgr.GetStockListByMarket(1)  # 거래소
    codeList2 = cpCodeMgr.GetStockListByMarket(2)  # 코스닥

    print('거래소 전체 종목 : {}'.format(len(codeList)))
    print('코스닥 전체 종목 : {}'.format(len(codeList2)))
    print("거래소 + 코스닥 종목코드 ", len(codeList) + len(codeList2))
    print("거래소 종목코드", len(codeList))
    kospi = {}
    for i, code in enumerate(codeList):
        secondCode = cpCodeMgr.GetStockSectionKind(code)
        name = cpCodeMgr.CodeToName(code)
        stdPrice = cpCodeMgr.GetStockStdPrice(code)
        kospi[name] = secondCode

    print("코스닥 종목코드", len(codeList2))
    kosdaq = {}
    for i, code in enumerate(codeList2):
        secondCode = cpCodeMgr.GetStockSectionKind(code)
        name = cpCodeMgr.CodeToName(code)
        kosdaq[name] = secondCode
    return kospi, kosdaq

def check_creon_system():
    """크레온 플러스 시스템 연결 상태를 점검한다."""
    # 관리자 권한으로 프로세스 실행 여부
    if not ctypes.windll.shell32.IsUserAnAdmin():
        print('check_creon_system() : admin user -> FAILED')
        return False

    # 연결 여부 체크
    if (cpStatus.IsConnect == 0):
        print('check_creon_system() : connect to server -> FAILED')
        return False

    # 주문 관련 초기화 - 계좌 관련 코드가 있을 때만 사용
    if (cpTradeUtil.TradeInit(0) != 0):
        print('check_creon_system() : init trade -> FAILED')
        return False
    return True


def get_current_price(code):
    """인자로 받은 종목의 현재가, 매수호가, 매도호가를 반환한다."""
    cpStock.SetInputValue(0, code)  # 종목코드에 대한 가격 정보
    cpStock.BlockRequest()
    item = {}
    item['cur_price'] = cpStock.GetHeaderValue(11)   # 현재가
    item['ask'] =  cpStock.GetHeaderValue(16)        # 매수호가
    item['bid'] =  cpStock.GetHeaderValue(17)        # 매도호가
    return item['cur_price'], item['ask'], item['bid']



def get_ohlc(code, qty):
    """
    인자로 받은 종목의 OHLC 가격 정보를 qty 개수만큼 반환한다.
    OHLC : 'open', 'high', 'low', 'close' 시가 / 고가 / 저가 / 종가
    """
    cpOhlc.SetInputValue(0, code)           # 종목코드
    cpOhlc.SetInputValue(1, ord('2'))        # 1:기간, 2:개수
    cpOhlc.SetInputValue(4, qty)             # 요청개수
    cpOhlc.SetInputValue(5, [0, 2, 3, 4, 5]) # 0:날짜, 2~5:OHLC
    cpOhlc.SetInputValue(6, ord('D'))        # D:일단위
    cpOhlc.SetInputValue(9, ord('1'))        # 0:무수정주가, 1:수정주가
    cpOhlc.BlockRequest()
    count = cpOhlc.GetHeaderValue(3)   # 3:수신개수
    columns = ['open', 'high', 'low', 'close']
    index = []
    rows = []
    for i in range(count):
        index.append(cpOhlc.GetDataValue(0, i))
        rows.append([cpOhlc.GetDataValue(1, i), cpOhlc.GetDataValue(2, i),
            cpOhlc.GetDataValue(3, i), cpOhlc.GetDataValue(4, i)])
    df = pd.DataFrame(rows, columns=columns, index=index)
    return df


def get_stock_balance(code):
    """인자로 받은 종목의 종목명과 수량을 반환한다."""
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]      # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체, 1:주식, 2:선물/옵션
    cpBalance.SetInputValue(0, acc)         # 계좌번호
    cpBalance.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
    cpBalance.SetInputValue(2, 50)          # 요청 건수(최대 50)
    cpBalance.BlockRequest()
    if code == 'ALL':
        print('계좌명: ' + str(cpBalance.GetHeaderValue(0)))
        print('결제잔고수량 : ' + str(cpBalance.GetHeaderValue(1)))
        print('평가금액: ' + str(cpBalance.GetHeaderValue(3)))
        print('평가손익: ' + str(cpBalance.GetHeaderValue(4)))
        print('종목수: ' + str(cpBalance.GetHeaderValue(7)))
        print('수익율: ' + str(cpBalance.GetHeaderValue(8)))

    stocks = []
    for i in range(cpBalance.GetHeaderValue(7)):
        stock_code = cpBalance.GetDataValue(12, i)  # 종목코드
        stock_name = cpBalance.GetDataValue(0, i)   # 종목명
        stock_qty = cpBalance.GetDataValue(15, i)   # 수량
        stock_rate = cpBalance.GetHeaderValue(8, i)   # 수익율

        if code == 'ALL':
            print(str(i+1) + ' ' + stock_code + '(' + stock_name + ')'
                + ':' + str(stock_qty) + '/' + str(stock_rate))
            stocks.append({'code': stock_code, 'name': stock_name,
                'qty': stock_qty, 'rate' : stock_rate})
        if stock_code == code:
            return stock_name, stock_qty
    if code == 'ALL':
        return stocks
    else:
        stock_name = cpCodeMgr.CodeToName(code)
        return stock_name, 0


def get_current_cash():
    """증거금 100% 주문 가능 금액을 반환한다."""
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]    # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체, 1:주식, 2:선물/옵션
    cpCash.SetInputValue(0, acc)              # 계좌번호
    cpCash.SetInputValue(1, accFlag[0])      # 상품구분 - 주식 상품 중 첫번째
    cpCash.BlockRequest()
    return cpCash.GetHeaderValue(9) # 증거금 100% 주문 가능 금액


def get_target_price(code):
    """매수 목표가를 반환한다."""
    try:
        time_now = datetime.now()
        str_today = time_now.strftime('%Y%m%d')
        ohlc = get_ohlc(code, 10)
        if str_today == str(ohlc.iloc[0].name):
            today_open = ohlc.iloc[0].open
            lastday = ohlc.iloc[1]
        else:
            lastday = ohlc.iloc[0]
            today_open = lastday[3]
        lastday_high = lastday[1]
        lastday_low = lastday[2]
        target_price = today_open + (lastday_high - lastday_low) * 0.5
        return target_price
    except Exception as ex:
        print("`get_target_price() -> exception! " + str(ex) + "`")
        return None


def get_movingaverage(code, window):
    """인자로 받은 종목에 대한 이동평균가격을 반환한다."""
    try:
        time_now = datetime.now()
        str_today = time_now.strftime('%Y%m%d')
        ohlc = get_ohlc(code, 20)
        if str_today == str(ohlc.iloc[0].name):
            lastday = ohlc.iloc[1].name
        else:
            lastday = ohlc.iloc[0].name
        closes = ohlc['close'].sort_index()
        ma = closes.rolling(window=window).mean()
        return ma.loc[lastday]
    except Exception as ex:
        print('get_movingavrg(' + str(window) + ') -> exception! ' + str(ex))
        return None


def buy_etf(code, buy_amount):
    """인자로 받은 종목을 최유리 지정가 FOK 조건으로 매수한다."""
    try:
        global bought_list  # 함수 내에서 값 변경을 하기 위해 global로 지정
        if code in bought_list:  # 매수 완료 종목이면 더 이상 안 사도록 함수 종료
            # printlog('code:', code, 'in', bought_list)
            return False
        time_now = datetime.now()
        current_price, ask_price, bid_price = get_current_price(code)
        target_price = get_target_price(code)  # 매수 목표가
        ma5_price = get_movingaverage(code, 5)  # 5일 이동평균가
        ma10_price = get_movingaverage(code, 10)  # 10일 이동평균가
        buy_qty = 0  # 매수할 수량 초기화
        if ask_price > 0:  # 매수호가가 존재하면
            buy_qty = buy_amount // ask_price
        stock_name, stock_qty = get_stock_balance(code)  # 종목명과 보유수량 조회
        # printlog('bought_list:', bought_list, 'len(bought_list):',
        #    len(bought_list), 'target_buy_count:', target_buy_count)
        if current_price > target_price and current_price > ma5_price \
                and current_price > ma10_price:
            print(stock_name + '(' + str(code) + ') ' + str(buy_qty) +
                     'EA : ' + str(current_price) + ' meets the buy condition!`')
            cpTradeUtil.TradeInit()
            acc = cpTradeUtil.AccountNumber[0]  # 계좌번호
            accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:전체,1:주식,2:선물/옵션
            # 최유리 FOK 매수 주문 설정
            cpOrder.SetInputValue(0, "2")  # 2: 매수
            cpOrder.SetInputValue(1, acc)  # 계좌번호
            cpOrder.SetInputValue(2, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
            cpOrder.SetInputValue(3, code)  # 종목코드
            cpOrder.SetInputValue(4, buy_qty)  # 매수할 수량
            cpOrder.SetInputValue(7, "2")  # 주문조건 0:기본, 1:IOC, 2:FOK
            cpOrder.SetInputValue(8, "12")  # 주문호가 1:보통, 3:시장가
            # 5:조건부, 12:최유리, 13:최우선
            # 매수 주문 요청
            ret = cpOrder.BlockRequest()
            print('최유리 FoK 매수 ->', stock_name, code, buy_qty, '->', ret)
            if ret == 4:
                remain_time = cpStatus.LimitRequestRemainTime
                print('주의: 연속 주문 제한에 걸림. 대기 시간:', remain_time / 1000)
                time.sleep(remain_time / 1000)
                return False
            time.sleep(2)
            print('현금주문 가능금액 :', buy_amount)
            stock_name, bought_qty = get_stock_balance(code)
            print('get_stock_balance :', stock_name, stock_qty)
            if bought_qty > 0:
                bought_list.append(code)
                print("`buy_etf(" + str(stock_name) + ' : ' + str(code) +
                       ") -> " + str(bought_qty) + "EA bought!" + "`")
    except Exception as ex:
        print("`buy_etf(" + str(code) + ") -> exception! " + str(ex) + "`")


def sell_all():
    """
    acc : 계좌번호
    target_code
    보유한 모든 종목을 최유리 지정가 IOC 조건으로 매도한다.

    IOC : 남은것들 중 체결되지 못한 잔량만 전부 취소
    FOK : 단 1주라도 체결이 안되면 주문 전체가 취소


    """
    try:
        cpTradeUtil.TradeInit()
        acc = cpTradeUtil.AccountNumber[0]  # 계좌번호
        accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:전체, 1:주식, 2:선물/옵션
        while True:
            stocks = get_stock_balance('ALL')
            total_qty = 0
            for s in stocks:
                total_qty += s['qty']
            if total_qty == 0:
                return True
            for s in stocks:
                if s['qty'] != 0:
                    cpOrder.SetInputValue(0, "1")  # 1:매도, 2:매수
                    cpOrder.SetInputValue(1, acc)  # 계좌번호
                    cpOrder.SetInputValue(2, accFlag[0])  # 주식상품 중 첫번째
                    cpOrder.SetInputValue(3, s['code'])  # 종목코드
                    cpOrder.SetInputValue(4, s['qty'])  # 매도수량
                    cpOrder.SetInputValue(7, "1")  # 조건 0:기본, 1:IOC, 2:FOK
                    cpOrder.SetInputValue(8, "12")  # 호가 12:최유리, 13:최우선
                    # 최유리 IOC 매도 주문 요청
                    ret = cpOrder.BlockRequest()
                    print('최유리 IOC 매도', s['code'], s['name'], s['qty'],
                             '-> cpOrder.BlockRequest() -> returned', ret)
                    if ret == 4:
                        remain_time = cpStatus.LimitRequestRemainTime
                        print('주의: 연속 주문 제한, 대기시간:', remain_time / 1000)
                time.sleep(1)
            time.sleep(30)
    except Exception as ex:
        print("sell_all() -> exception! " + str(ex))



# 현재잔고 확인하기
def deposit_chk() :
    b_connect = cpStatus.IsConnect
    if b_connect == 0:
        print(209, "PLUS가 정상적으로 연결되지 않음. ")
    #### 초기화
    instCheck = cpTradeUtil.TradeInit(0)
    if (instCheck != 0):
        print("Initialization Fail")
    else:
        print("Initialization Success")

    ### 계좌정보 확인
    account = cpTradeUtil.AccountNumber[0]
    accFlag = cpTradeUtil.GoodsList(account, 1)
    cpCash.SetInputValue(0, account)
    cpCash.SetInputValue(1, accFlag[0])
    cpCash.BlockRequest()
    print('계좌번호 : {}'.format(account))
    print('예수금(증거금 100%) : {}'.format(cpCash.GetHeaderValue(9)))
    return

#주식매수주문
'''
company_code : 종목코드
buy_quantity : 매수수량
buy_price : 주문단가
'''
def buy(company_code = None, buy_quantity = None, buy_price = None ) :
    #### 초기화
    instCheck = cpTradeUtil.TradeInit(0)
    if (instCheck != 0):
        print("Initialization Fail")
    else:
        print("Initialization Success")
    # 주식 매수 주문
    acc = cpTradeUtil.AccountNumber[0]  # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1)  # 주식상품 구분
    print(acc, accFlag[0])
    objStockOrder = win32com.client.Dispatch("CpTrade.CpTd0311")
    objStockOrder.SetInputValue(0, "2")  # 2: 매수
    objStockOrder.SetInputValue(1, acc)  # 계좌번호
    objStockOrder.SetInputValue(2, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
    objStockOrder.SetInputValue(3, company_code)  # 종목코드 - 필요한 종목으로 변경 필요
    objStockOrder.SetInputValue(4, buy_quantity)  # 매수수량 - 요청 수량으로 변경 필요
    objStockOrder.SetInputValue(5, buy_price)  # 주문단가 - 필요한 가격으로 변경 필요
    objStockOrder.SetInputValue(7, "0")  # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
    objStockOrder.SetInputValue(8, "01")  # 주문호가 구분코드 - 01: 보통
    # 매수 주문 요청
    nRet = objStockOrder.BlockRequest()
    if (nRet != 0):
        print("주문요청 오류", nRet)
        # 0: 정상,  그 외 오류, 4: 주문요청제한 개수 초과
        exit()
    return

# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client):
        self.client = client

    def OnReceived(self):
        code = self.client.GetHeaderValue(0)  # 초
        name = self.client.GetHeaderValue(1)  # 초
        timess = self.client.GetHeaderValue(18)  # 초
        exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그
        cprice = self.client.GetHeaderValue(13)  # 현재가
        diff = self.client.GetHeaderValue(2)  # 대비
        cVol = self.client.GetHeaderValue(17)  # 순간체결수량
        vol = self.client.GetHeaderValue(9)  # 거래량

        if (exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            print("실시간(예상체결)", name, timess, "*", cprice, "대비", diff, "체결량", cVol, "거래량", vol)
        elif (exFlag == ord('2')):  # 장중(체결)
            print("실시간(장중 체결)", name, timess, cprice, "대비", diff, "체결량", cVol, "거래량", vol)

# CpStockCur: 실시간 현재가 요청 클래스
class CpStockCur:
    def Subscribe(self, code):
        self.objStockCur = win32com.client.Dispatch("DsCbo1.StockCur")
        handler = win32com.client.WithEvents(self.objStockCur, CpEvent)
        self.objStockCur.SetInputValue(0, code)
        handler.set_params(self.objStockCur)
        self.objStockCur.Subscribe()

    def Unsubscribe(self):
        self.objStockCur.Unsubscribe()


class Cp7043:
    def __init__(self):
        # 통신 OBJECT 기본 세팅
        self.objRq = win32com.client.Dispatch("CpSysDib.CpSvrNew7043")
        self.objRq.SetInputValue(0, ord('0'))  # 거래소 + 코스닥
        self.objRq.SetInputValue(1, ord('2'))  # 상승
        self.objRq.SetInputValue(2, ord('1'))  # 당일
        self.objRq.SetInputValue(3, 21)  # 전일 대비 상위 순
        # self.objRq.SetInputValue(3, 21)  # 전일 대비 상위 순
        self.objRq.SetInputValue(4, ord('1'))  # 관리 종목 제외
        self.objRq.SetInputValue(5, ord('0'))  # 거래량 전체
        self.objRq.SetInputValue(6, ord('0'))  # '표시 항목 선택 - '0': 시가대비
        self.objRq.SetInputValue(7, 0)  # 등락율 시작
        self.objRq.SetInputValue(8, 30)  # 등락율 끝

    # 실제적인 7043 통신 처리
    def rq7043(self, retcode, data):
        self.objRq.BlockRequest()
        # 현재가 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        rqRet = self.objRq.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        cnt = self.objRq.GetHeaderValue(0)
        cntTotal = self.objRq.GetHeaderValue(1)
        print(cnt, cntTotal)
        for i in range(cnt):
            code = self.objRq.GetDataValue(0, i)  # 코드
            retcode.append(code)
            name = self.objRq.GetDataValue(1, i)  # 종목명
            price = self.objRq.GetDataValue(2, i)  # 종목명
            diffflag = self.objRq.GetDataValue(3, i)  #
            diff = self.objRq.GetDataValue(4, i)
            rate = self.objRq.GetDataValue(5, i)
            vol = self.objRq.GetDataValue(6, i)  # 거래량
            # print(code, name, diffflag, diff, vol)
            data.append((code, name, price, diffflag, diff, rate, vol))
            if len(retcode) >= 200:  # 최대 200 종목만,
                break
        return data

    def Request(self, retCode, data):
        self.rq7043(retCode, data)
        # 연속 데이터 조회 - 200 개까지만.
        while self.objRq.Continue:
            (self.rq7043(retCode, data))
            # print(len(retCode))
            if len(retCode) >= 200:
                break

        # #7043 상승하락 서비스를 통해 받은 상승률 상위 200 종목
        # size = len(retCode)
        # for i in range(size):
        #    print(retCode[i])
        return True


# CpMarketEye : 복수종목 현재가 통신 서비스
def CpMarketEye_v2(codes):
    # 연결 여부 체크
    objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
    bConnect = objCpCybos.IsConnect
    if (bConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음. ")
        return False

    # 관심종목 객체 구하기
    objRq = win32com.client.Dispatch("CpSysDib.MarketEye")
    # 요청 필드 세팅 - 종목코드, 시간, 대비부호, 대비, 현재가,시가, 매도호가, 매수호가, 거래량, 거래대금, 전일거래량, 체결강도
    rqField = [0, 1, 2, 3, 4, 5, 8, 9, 10, 11, 22, 24]  # 요청 필드
    objRq.SetInputValue(0, rqField)  # 요청 필드
    objRq.SetInputValue(1, codes)  # 종목코드 or 종목코드 리스트
    objRq.BlockRequest()

    # 현재가 통신 및 통신 에러 처리
    rqStatus = objRq.GetDibStatus()
    rqRet = objRq.GetDibMsg1()
    print("통신상태", rqStatus, rqRet)
    if rqStatus != 0:
        return False

    cnt = objRq.GetHeaderValue(2)

    data = []
    for i in range(cnt):
        # rpCode = objRq.GetDataValue(0, i)  # 코드
        # # rpName = objRq.GetDataValue(1, i)  # 종목명
        # rpTime = objRq.GetDataValue(1, i)  # 시간
        # rpDiffFlag = objRq.GetDataValue(2, i)  # 대비부호
        # rpDiff = objRq.GetDataValue(3, i)  # 대비
        # rpCur = objRq.GetDataValue(4, i)  # 현재가
        # rpStart = objRq.GetDataValue(5, i)  # 현재가
        # sellCall = objRq.GetDataValue(6, i)  # 매도호가
        # buyCall = objRq.GetDataValue(7, i)  # 매수호가
        # rpVol = objRq.GetDataValue(8, i)  # 거래량
        # rpPrice = objRq.GetDataValue(9, i)  # 거래대금
        # rpPrice_d_1 = objRq.GetDataValue(10, i)  # 전일거래량
        # rpPower = objRq.GetDataValue(11, i)  # 체결강도
        # if i % 100 == 0 :
        #     print(rpCode, rpTime, rpDiffFlag, rpDiff, rpCur, rpStart, sellCall, buyCall, rpVol, rpPrice, rpPrice_d_1, rpPower)
        #     data.append((rpCode, rpTime, rpDiffFlag, rpDiff, rpCur,rpStart, sellCall, buyCall, rpVol, rpPrice, rpPrice_d_1, rpPower))
        tmp = [objRq.GetDataValue(v,i) for v in range(len(rqField))]
        data.append(tmp)
        if i % 100 == 0 :
            print(tmp)
    return data

#주식차트 조회(분차트 / 틱차트)

class CpStockChart():
    def __init__(self):
        self.objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")

    # 차트 요청 - 기간 기준으로
    def RequestFromTo(self, code, fromDate, toDate, caller):
        print(code, fromDate, toDate)
        # 연결 여부 체크
        bConnect = cpStatus.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return False
        self.objStockChart.SetInputValue(0, code)  # 종목코드
        self.objStockChart.SetInputValue(1, ord('1'))  # 기간으로 받기
        self.objStockChart.SetInputValue(2, toDate)  # To 날짜
        self.objStockChart.SetInputValue(3, fromDate)  # From 날짜
        # self.objStockChart.SetInputValue(4, 500)  # 최근 500일치
        self.objStockChart.SetInputValue(5, [0, 2, 3, 4, 5, 8])  # 날짜,시가,고가,저가,종가,거래량
        self.objStockChart.SetInputValue(6, ord('D'))  # '차트 주기 - 일간 차트 요청
        self.objStockChart.SetInputValue(9, ord('1'))  # 수정주가 사용
        self.objStockChart.BlockRequest()

        rqStatus = self.objStockChart.GetDibStatus()
        rqRet = self.objStockChart.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            exit()

        len = self.objStockChart.GetHeaderValue(3)

        caller.dates = []
        caller.opens = []
        caller.highs = []
        caller.lows = []
        caller.closes = []
        caller.vols = []
        for i in range(len):
            caller.dates.append(self.objStockChart.GetDataValue(0, i))
            caller.opens.append(self.objStockChart.GetDataValue(1, i))
            caller.highs.append(self.objStockChart.GetDataValue(2, i))
            caller.lows.append(self.objStockChart.GetDataValue(3, i))
            caller.closes.append(self.objStockChart.GetDataValue(4, i))
            caller.vols.append(self.objStockChart.GetDataValue(5, i))

        print(len)
    # 차트 요청 - 최근일 부터 개수 기준
    def RequestDWM(self, code, dwm, count, caller):
        # 연결 여부 체크
        bConnect = cpStatus.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return False

        self.objStockChart.SetInputValue(0, code)  # 종목코드
        self.objStockChart.SetInputValue(1, ord('2'))  # 개수로 받기
        self.objStockChart.SetInputValue(4, count)  # 최근 500일치
        self.objStockChart.SetInputValue(5, [0, 2, 3, 4, 5, 8])  # 요청항목 - 날짜,시가,고가,저가,종가,거래량
        self.objStockChart.SetInputValue(6, dwm)  # '차트 주기 - 일/주/월
        self.objStockChart.SetInputValue(9, ord('1'))  # 수정주가 사용
        self.objStockChart.BlockRequest()

        rqStatus = self.objStockChart.GetDibStatus()
        rqRet = self.objStockChart.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            exit()

        len = self.objStockChart.GetHeaderValue(3)

        caller.dates = []
        caller.opens = []
        caller.highs = []
        caller.lows = []
        caller.closes = []
        caller.vols = []
        caller.times = []
        for i in range(len):
            caller.dates.append(self.objStockChart.GetDataValue(0, i))
            caller.opens.append(self.objStockChart.GetDataValue(1, i))
            caller.highs.append(self.objStockChart.GetDataValue(2, i))
            caller.lows.append(self.objStockChart.GetDataValue(3, i))
            caller.closes.append(self.objStockChart.GetDataValue(4, i))
            caller.vols.append(self.objStockChart.GetDataValue(5, i))

        print(len)

        return
    # 차트 요청 - 분간, 틱 차트
    def RequestMT(self, code, dwm, count, caller):
        # 연결 여부 체크
        bConnect = cpStatus.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return False

        self.objStockChart.SetInputValue(0, code)  # 종목코드
        self.objStockChart.SetInputValue(1, ord('2'))  # 개수로 받기
        self.objStockChart.SetInputValue(4, count)  # 조회 개수
        self.objStockChart.SetInputValue(5, [0, 1, 2, 3, 4, 5, 8])  # 요청항목 - 날짜, 시간,시가,고가,저가,종가,거래량
        self.objStockChart.SetInputValue(6, dwm)  # '차트 주기 - 분/틱
        self.objStockChart.SetInputValue(7, 1)  # 분틱차트 주기
        self.objStockChart.SetInputValue(9, ord('1'))  # 수정주가 사용
        self.objStockChart.BlockRequest()

        rqStatus = self.objStockChart.GetDibStatus()
        rqRet = self.objStockChart.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            exit()

        len = self.objStockChart.GetHeaderValue(3)
        caller.dates = []
        caller.opens = []
        caller.highs = []
        caller.lows = []
        caller.closes = []
        caller.vols = []
        caller.times = []
        for i in range(len):
            caller.dates.append(self.objStockChart.GetDataValue(0, i))
            caller.times.append(self.objStockChart.GetDataValue(1, i))
            caller.opens.append(self.objStockChart.GetDataValue(2, i))
            caller.highs.append(self.objStockChart.GetDataValue(3, i))
            caller.lows.append(self.objStockChart.GetDataValue(4, i))
            caller.closes.append(self.objStockChart.GetDataValue(5, i))
            caller.vols.append(self.objStockChart.GetDataValue(6, i))

        print(len)

        return


# 현재가 정보 조회
class CpStockMst():
    def __init__(self):
        # 연결 여부 체크
        connect_status()

    def Request(self, code):

        # 현재가 객체 구하기
        objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
        objStockMst.SetInputValue(0, code)  # 종목 코드 - 삼성전자
        objStockMst.BlockRequest()

        # 현재가 통신 및 통신 에러 처리
        rqStatus = objStockMst.GetDibStatus()
        rqRet = objStockMst.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        # 현재가 정보 조회
        code = objStockMst.GetHeaderValue(0)  # 종목코드
        name = objStockMst.GetHeaderValue(1)  # 종목명
        time = objStockMst.GetHeaderValue(4)  # 시간
        cprice = objStockMst.GetHeaderValue(11)  # 종가
        diff = objStockMst.GetHeaderValue(12)  # 대비
        open = objStockMst.GetHeaderValue(13)  # 시가
        high = objStockMst.GetHeaderValue(14)  # 고가
        low = objStockMst.GetHeaderValue(15)  # 저가
        offer = objStockMst.GetHeaderValue(16)  # 매도호가
        bid = objStockMst.GetHeaderValue(17)  # 매수호가
        vol = objStockMst.GetHeaderValue(18)  # 거래량
        vol_value = objStockMst.GetHeaderValue(19)  # 거래대금

        print("코드 이름 시간 현재가 대비 시가 고가 저가 매도호가 매수호가 거래량 거래대금")
        print(code, name, time, cprice, diff, open, high, low, offer, bid, vol, vol_value)
        return True


'''
주문 초기화
'''
def init_trade() :
    objTrade =  win32com.client.Dispatch("CpTrade.CpTdUtil")
    initCheck = objTrade.TradeInit(0)
    if (initCheck != 0):
        print("주문 초기화 실패")
        exit()

'''
실시간 수신
'''


################################################
# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client, name, caller):
        self.client = client  # CP 실시간 통신 object
        self.name = name  # 서비스가 다른 이벤트를 구분하기 위한 이름
        self.caller = caller  # callback 을 위해 보관

        # 구분값 : 텍스트로 변경하기 위해 딕셔너리 이용
        self.dicflag12 = {'1': '매도', '2': '매수'}
        self.dicflag14 = {'1': '체결', '2': '확인', '3': '거부', '4': '접수'}
        self.dicflag15 = {'00': '현금', '01': '유통융자', '02': '자기융자', '03': '유통대주',
                          '04': '자기대주', '05': '주식담보대출', '07': '채권담보대출',
                          '06': '매입담보대출', '08': '플러스론',
                          '13': '자기대용융자', '15': '유통대용융자'}
        self.dicflag16 = {'1': '정상주문', '2': '정정주문', '3': '취소주문'}
        self.dicflag17 = {'1': '현금', '2': '신용', '3': '선물대용', '4': '공매도'}
        self.dicflag18 = {'01': '보통', '02': '임의', '03': '시장가', '05': '조건부지정가'}
        self.dicflag19 = {'0': '없음', '1': 'IOC', '2': 'FOK'}

    def OnReceived(self):
        # 실시간 처리 - 현재가 주문 체결
        if self.name == 'stockcur':
            code = self.client.GetHeaderValue(0)  # 초
            name = self.client.GetHeaderValue(1)  # 초
            timess = self.client.GetHeaderValue(18)  # 초
            exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그
            cprice = self.client.GetHeaderValue(13)  # 현재가
            diff = self.client.GetHeaderValue(2)  # 대비
            cVol = self.client.GetHeaderValue(17)  # 순간체결수량
            vol = self.client.GetHeaderValue(9)  # 거래량

            item = {}
            item['code'] = code
            # rpName = self.objRq.GetDataValue(1, i)  # 종목명
            # rpDiffFlag = self.objRq.GetDataValue(3, i)  # 대비부호
            item['diff'] = diff
            item['cur'] = cprice
            item['vol'] = vol

            # 현재가 업데이트
            self.caller.updateJangoCurPBData(item)

        # 실시간 처리 - 주문체결
        elif self.name == 'conclution':
            # 주문 체결 실시간 업데이트
            conc = {}

            # 체결 플래그
            conc['체결플래그'] = self.dicflag14[self.client.GetHeaderValue(14)]

            conc['주문번호'] = self.client.GetHeaderValue(5)  # 주문번호
            conc['주문수량'] = self.client.GetHeaderValue(3)  # 주문/체결 수량
            conc['주문가격'] = self.client.GetHeaderValue(4)  # 주문/체결 가격
            conc['원주문'] = self.client.GetHeaderValue(6)
            conc['종목코드'] = self.client.GetHeaderValue(9)  # 종목코드
            conc['종목명'] = cpStockCd.CodeToName(conc['종목코드'])

            conc['매수매도'] = self.dicflag12[self.client.GetHeaderValue(12)]

            flag15 = self.client.GetHeaderValue(15)  # 신용대출구분코드
            if (flag15 in self.dicflag15):
                conc['신용대출'] = self.dicflag15[flag15]
            else:
                conc['신용대출'] = '기타'

            conc['정정취소'] = self.dicflag16[self.client.GetHeaderValue(16)]
            conc['현금신용'] = self.dicflag17[self.client.GetHeaderValue(17)]
            conc['주문조건'] = self.dicflag19[self.client.GetHeaderValue(19)]

            conc['체결기준잔고수량'] = self.client.GetHeaderValue(23)
            loandate = self.client.GetHeaderValue(20)
            if (loandate == 0):
                conc['대출일'] = ''
            else:
                conc['대출일'] = str(loandate)
            flag18 = self.client.GetHeaderValue(18)
            if (flag18 in self.dicflag18):
                conc['주문호가구분'] = self.dicflag18[flag18]
            else:
                conc['주문호가구분'] = '기타'

            conc['장부가'] = self.client.GetHeaderValue(21)
            conc['매도가능수량'] = self.client.GetHeaderValue(22)

            print(conc)
            self.caller.updateJangoCont(conc)

            return

'''
주식 잔고 조회
'''


class Cp6033:
    def __init__(self):
        acc = cpTradeUtil.AccountNumber[0]  # 계좌번호
        accFlag = cpTradeUtil.GoodsList(acc, 1)  # 주식상품 구분
        print(acc, accFlag[0])

        self.objRq = win32com.client.Dispatch("CpTrade.CpTd6033")
        self.objRq.SetInputValue(0, acc)  # 계좌번호
        self.objRq.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objRq.SetInputValue(2, 50)  # 요청 건수(최대 50)
        self.dicflag1 = {ord(' '): '현금',
                         ord('Y'): '융자',
                         ord('D'): '대주',
                         ord('B'): '담보',
                         ord('M'): '매입담보',
                         ord('P'): '플러스론',
                         ord('I'): '자기융자',
                         }

    # 실제적인 6033 통신 처리
    def requestJango(self, caller):
        while True:
            self.objRq.BlockRequest()
            # 통신 및 통신 에러 처리
            rqStatus = self.objRq.GetDibStatus()
            rqRet = self.objRq.GetDibMsg1()
            print("통신상태", rqStatus, rqRet)
            if rqStatus != 0:
                return False

            cnt = self.objRq.GetHeaderValue(7)
            print(cnt)

            for i in range(cnt):
                item = {}
                code = self.objRq.GetDataValue(12, i)  # 종목코드
                item['종목코드'] = code
                item['종목명'] = self.objRq.GetDataValue(0, i)  # 종목명
                item['현금신용'] = self.dicflag1[self.objRq.GetDataValue(1, i)]  # 신용구분
                print(code, '현금신용', item['현금신용'])
                item['대출일'] = self.objRq.GetDataValue(2, i)  # 대출일
                item['잔고수량'] = self.objRq.GetDataValue(7, i)  # 체결잔고수량
                item['매도가능'] = self.objRq.GetDataValue(15, i)
                item['장부가'] = self.objRq.GetDataValue(17, i)  # 체결장부단가
                # item['평가금액'] = self.objRq.GetDataValue(9, i)  # 평가금액(천원미만은 절사 됨)
                # item['평가손익'] = self.objRq.GetDataValue(11, i)  # 평가손익(천원미만은 절사 됨)
                # 매입금액 = 장부가 * 잔고수량
                item['매입금액'] = item['장부가'] * item['잔고수량']
                item['현재가'] = 0
                item['대비'] = 0
                item['거래량'] = 0

                # 잔고 추가
                #                key = (code, item['현금신용'],item['대출일'] )
                key = code
                caller.jangoData[key] = item

                if len(caller.jangoData) >= 200:  # 최대 200 종목만,
                    break

            if len(caller.jangoData) >= 200:
                break
            if (self.objRq.Continue == False):
                break
        return True





