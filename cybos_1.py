'''
CYBOS-PLUS 홈페이지 예제 정리

written by woosikyang
date : 2020-09-03


전체 TITLE LIST

- 실시간 종목조회
- 주식 현재가 조회
- 주식 일자별 조회
- 주식 현금 매수주문
- 주식 현금 매도주문







'''


'''
TITLE : 실시간 종목조회
'''
import win32com.client

# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()

# 종목코드 리스트 구하기
objCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
codeList = objCpCodeMgr.GetStockListByMarket(1)  # 거래소
codeList2 = objCpCodeMgr.GetStockListByMarket(2)  # 코스닥
print('거래소 전체 종목 : {}'.format(len(codeList)))
print('코스닥 전체 종목 : {}'.format(len(codeList2)))
print("거래소 + 코스닥 종목코드 ", len(codeList) + len(codeList2))


print("거래소 종목코드", len(codeList))
for i, code in enumerate(codeList):
    secondCode = objCpCodeMgr.GetStockSectionKind(code)
    name = objCpCodeMgr.CodeToName(code)
    stdPrice = objCpCodeMgr.GetStockStdPrice(code)
    print(i, code, secondCode, stdPrice, name)



print("코스닥 종목코드", len(codeList2))
for i, code in enumerate(codeList2):
    secondCode = objCpCodeMgr.GetStockSectionKind(code)
    name = objCpCodeMgr.CodeToName(code)
    stdPrice = objCpCodeMgr.GetStockStdPrice(code)
    print(i, code, secondCode, stdPrice, name)


'''
TITLE : 주식 현재가 조회
'''


# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")


# 현재가 객체 구하기
objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
objStockMst.SetInputValue(0, 'A005930')  # 종목 코드 - 삼성전자
objStockMst.BlockRequest()

# 현재가 통신 및 통신 에러 처리
rqStatus = objStockMst.GetDibStatus()
rqRet = objStockMst.GetDibMsg1()
print("통신상태", rqStatus, rqRet)
'''
통신상태 0 0027 조회가 완료되었습니다.(stock.new.mst)
'''

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

# 예상 체결관련 정보
exFlag = objStockMst.GetHeaderValue(58)  # 예상체결가 구분 플래그
exPrice = objStockMst.GetHeaderValue(55)  # 예상체결가
exDiff = objStockMst.GetHeaderValue(56)  # 예상체결가 전일대비
exVol = objStockMst.GetHeaderValue(57)  # 예상체결수량

print("코드", code)
print("이름", name)
print("시간", time)
print("종가", cprice)
print("대비", diff)
print("시가", open)
print("고가", high)
print("저가", low)
print("매도호가", offer)
print("매수호가", bid)
print("거래량", vol)
print("거래대금", vol_value)

'''결과 -2017.10.29 기준
코드 A005930
이름 삼성전자
시간 1556
종가 2654000
대비 34000
시가 2620000
고가 2666000
저가 2607000
매도호가 2657000
매수호가 2654000
거래량 147850
거래대금 39046225
'''


'''
TITLE : 주식 일자별 조회
'''


def ReqeustData(obj):
    # 데이터 요청
    obj.BlockRequest()

    # 통신 결과 확인
    rqStatus = obj.GetDibStatus()
    rqRet = obj.GetDibMsg1()
    print("통신상태", rqStatus, rqRet)
    if rqStatus != 0:
        return False

    # 일자별 정보 데이터 처리
    count = obj.GetHeaderValue(1)  # 데이터 개수
    for i in range(count):
        date = obj.GetDataValue(0, i)  # 일자
        open = obj.GetDataValue(1, i)  # 시가
        high = obj.GetDataValue(2, i)  # 고가
        low = obj.GetDataValue(3, i)  # 저가
        close = obj.GetDataValue(4, i)  # 종가
        diff = obj.GetDataValue(5, i)  # 종가
        vol = obj.GetDataValue(6, i)  # 종가
        print(date, open, high, low, close, diff, vol)

    return True


# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()

# 일자별 object 구하기
objStockWeek = win32com.client.Dispatch("DsCbo1.StockWeek")
objStockWeek.SetInputValue(0, 'A005930')  # 종목 코드 - 삼성전자

# 최초 데이터 요청
ret = ReqeustData(objStockWeek)
if ret == False:
    exit()

# 연속 데이터 요청
# 예제는 5번만 연속 통신 하도록 함.
NextCount = 1
while objStockWeek.Continue:  # 연속 조회처리
    NextCount += 1;
    if (NextCount > 5):
        break
    ret = ReqeustData(objStockWeek)
    if ret == False:
        exit()


'''
TITLE : 주식 현금 매수주문
'''

# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()

# 주문 초기화
objTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")
initCheck = objTrade.TradeInit(0)
if (initCheck != 0):
    print("주문 초기화 실패")
    exit()

# 주식 매수 주문
acc = objTrade.AccountNumber[0]  # 계좌번호
accFlag = objTrade.GoodsList(acc, 1)  # 주식상품 구분
print(acc, accFlag[0])
objStockOrder = win32com.client.Dispatch("CpTrade.CpTd0311")
objStockOrder.SetInputValue(0, "2")  # 2: 매수
objStockOrder.SetInputValue(1, acc)  # 계좌번호
objStockOrder.SetInputValue(2, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
objStockOrder.SetInputValue(3, "A003540")  # 종목코드 - A003540 - 대신증권 종목
objStockOrder.SetInputValue(4, 10)  # 매수수량 10주
objStockOrder.SetInputValue(5, 14100)  # 주문단가  - 14,100원
objStockOrder.SetInputValue(7, "0")  # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
objStockOrder.SetInputValue(8, "01")  # 주문호가 구분코드 - 01: 보통

# 매수 주문 요청
objStockOrder.BlockRequest()

rqStatus = objStockOrder.GetDibStatus()
rqRet = objStockOrder.GetDibMsg1()
print("통신상태", rqStatus, rqRet)
if rqStatus != 0:
    exit()



'''
TITLE : 주식 현금 매도주문
'''

# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()

# 주문 초기화
objTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")
initCheck = objTrade.TradeInit(0)
if (initCheck != 0):
    print("주문 초기화 실패")
    exit()

# 주식 매도 주문
acc = objTrade.AccountNumber[0]  # 계좌번호
accFlag = objTrade.GoodsList(acc, 1)  # 주식상품 구분
print(acc, accFlag[0])
objStockOrder = win32com.client.Dispatch("CpTrade.CpTd0311")
objStockOrder.SetInputValue(0, "1")  # 1: 매도
objStockOrder.SetInputValue(1, acc)  # 계좌번호
objStockOrder.SetInputValue(2, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
objStockOrder.SetInputValue(3, "A003540")  # 종목코드 - A003540 - 대신증권 종목
objStockOrder.SetInputValue(4, 10)  # 매도수량 10주
objStockOrder.SetInputValue(5, 14100)  # 주문단가  - 14,100원
objStockOrder.SetInputValue(7, "0")  # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
objStockOrder.SetInputValue(8, "01")  # 주문호가 구분코드 - 01: 보통

# 매도 주문 요청
objStockOrder.BlockRequest()

rqStatus = objStockOrder.GetDibStatus()
rqRet = objStockOrder.GetDibMsg1()
print("통신상태", rqStatus, rqRet)
if rqStatus != 0:
    exit()


'''
TITLE : 차트 데이터 구하는 예제
'''

# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()

# 차트 객체 구하기
objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")

objStockChart.SetInputValue(0, 'A005930')  # 종목 코드 - 삼성전자
objStockChart.SetInputValue(1, ord('2'))  # 개수로 조회
objStockChart.SetInputValue(4, 100)  # 최근 100일 치
objStockChart.SetInputValue(5, [0, 2, 3, 4, 5, 8])  # 날짜,시가,고가,저가,종가,거래량
objStockChart.SetInputValue(6, ord('D'))  # '차트 주가 - 일간 차트 요청
objStockChart.SetInputValue(9, ord('1'))  # 수정주가 사용
objStockChart.BlockRequest()

len = objStockChart.GetHeaderValue(3)

print("날짜", "시가", "고가", "저가", "종가", "거래량")
print("빼기빼기==============================================-")

for i in range(len):
    day = objStockChart.GetDataValue(0, i)
    open = objStockChart.GetDataValue(1, i)
    high = objStockChart.GetDataValue(2, i)
    low = objStockChart.GetDataValue(3, i)
    close = objStockChart.GetDataValue(4, i)
    vol = objStockChart.GetDataValue(5, i)
    print(day, open, high, low, close, vol)