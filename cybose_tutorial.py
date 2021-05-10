import win32com.client



objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
cpCash = win32com.client.Dispatch('CpTrade.CpTdNew5331A')
CpTdUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')



# 현재잔고 확인하기
def deposit_chk() :
    b_connect = objCpCybos.IsConnect
    if b_connect == 0:
        print(209, "PLUS가 정상적으로 연결되지 않음. ")
    #### 초기화
    instCheck = CpTdUtil.TradeInit(0)
    if (instCheck != 0):
        print("Initialization Fail")
    else:
        print("Initialization Success")

    ### 계좌정보 확인
    account = CpTdUtil.AccountNumber[0]
    accFlag = CpTdUtil.GoodsList(account, 1)
    cpCash.SetInputValue(0, account)
    cpCash.SetInputValue(1, accFlag[0])
    cpCash.BlockRequest()
    print('계좌번호 : {}'.format(account))
    print('예수금(증거금 100%) : {}'.format(cpCash.GetHeaderValue(9)))

#주식매수주문
'''
company_code : 종목코드
buy_quantity : 매수수량
buy_price : 주문단가
'''
def buy(company_code = None, buy_quantity = None, buy_price = None ) :
    #### 초기화
    instCheck = CpTdUtil.TradeInit(0)
    if (instCheck != 0):
        print("Initialization Fail")
    else:
        print("Initialization Success")
    # 주식 매수 주문
    acc = CpTdUtil.AccountNumber[0]  # 계좌번호
    accFlag = CpTdUtil.GoodsList(acc, 1)  # 주식상품 구분
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

