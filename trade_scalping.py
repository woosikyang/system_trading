'''
초단타 스켈핑

Step 1 : 상승률 상위 200종목 선정

1-1 종목 선정 주기
-> 특정 조건에 따라?
-> 특정 시간에 따라?

Step 2 : 상위 200 중 조건에 맞는 종목 필터링

Step 3 : 해당 종목 매수

Step 4 : 해당 종목 매도(특정 조건 하)


- 매 스텝 당 slack으로 메세지 전송


'''
from slack import post_message
import win32com.client
import os, sys, ctypes
import pandas as pd
from datetime import datetime
import time, calendar
from cybos_base import *
from configs import *



if __name__ == '__main__':
    try:
        symbol_list = ['A122630', 'A252670', 'A233740', 'A250780', 'A225130',
                       'A280940', 'A261220', 'A217770', 'A295000', 'A176950']
        bought_list = []  # 매수 완료된 종목 리스트
        target_buy_count = 5  # 매수할 종목 수
        buy_percent = 0.19
        print('check_creon_system() :', check_creon_system())  # 크레온 접속 점검
        stocks = get_stock_balance('ALL')  # 보유한 모든 종목 조회
        total_cash = int(get_current_cash())  # 100% 증거금 주문 가능 금액 조회
        buy_amount = total_cash * buy_percent  # 종목별 주문 금액 계산
        print('100% 증거금 주문 가능 금액 :', total_cash)
        print('종목별 주문 비율 :', buy_percent)
        print('종목별 주문 금액 :', buy_amount)
        print('시작 시간 :', datetime.now().strftime('%m/%d %H:%M:%S'))
        soldout = False;

        while True:
            t_now = datetime.now()
            t_9 = t_now.replace(hour=9, minute=0, second=0, microsecond=0)
            t_start = t_now.replace(hour=9, minute=5, second=0, microsecond=0)
            t_sell = t_now.replace(hour=15, minute=15, second=0, microsecond=0)
            t_exit = t_now.replace(hour=15, minute=20, second=0, microsecond=0)
            today = datetime.today().weekday()
            if today == 5 or today == 6:  # 토요일이나 일요일이면 자동 종료
                print('Today is', 'Saturday.' if today == 5 else 'Sunday.')
                sys.exit(0)
            if t_9 < t_now < t_start and soldout == False:
                soldout = True
                sell_all()
                text = '매도 완료'
                post_message(slack_api_token, "#주식", text)
            if t_start < t_now < t_sell:  # AM 09:05 ~ PM 03:15 : 매수
                for sym in symbol_list:
                    if len(bought_list) < target_buy_count:
                        buy_etf(sym)
                        time.sleep(1)
                        text = '매수 완료'
                        post_message(slack_api_token, "#주식", text)
                if t_now.minute == 30 and 0 <= t_now.second <= 5:
                    get_stock_balance('ALL')
                    text = '전체 종목 개수'
                    post_message(slack_api_token, "#주식", text)
                    time.sleep(5)
            if t_sell < t_now < t_exit:  # PM 03:15 ~ PM 03:20 : 일괄 매도
                if sell_all() == True:
                    print('`sell_all() returned True -> self-destructed!`')
                    sys.exit(0)
            if t_exit < t_now:  # PM 03:20 ~ :프로그램 종료
                print('`self-destructed!`')
                sys.exit(0)
            time.sleep(3)
    except Exception as ex:
        print('`main -> exception! ' + str(ex) + '`')

