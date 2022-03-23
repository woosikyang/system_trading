'''
종목 필터링을 위한 함수들

input : 튜블 형태로 담긴 종목코드 및 종목정보들에 대한 리스트

return : 종목 정보로부터 설정된 조건에 맞춘 종목코드
'''
from cybos_base import get_current_price, CpStockMst
import time

def filtering(data_list) :
    '''

    :param data_list: 종목코드, 시간, 대비부호, 대비, 현재가,시가, 매도호가, 매수호가, 거래량, 거래대금, 전일거래량, 체결강도

    # 조건

    대비부호
    -> 상승인 종목만

    상승률 : 현재가 / 시가
    -> 3퍼 이상 20퍼 이하인 것들만

    체결강도 : 특정 시점 매수 / 매도 * 100
    -> 100 초과

    :return:
    '''
    # 조건 1
    # data_list = [v for v in data_list if v[2] =='2']
    # 조건 2
    data_list = [v for v in data_list if (1.03 < v[4]/v[5]) and (v[4]/v[5] < 1.05)]
    # 조건 3
    data_list = [v for v in data_list if v[11] > 150]

    return data_list



def filtering_v2(data_list) :
    '''

    :param data_list: 종목코드, 시간, 대비부호, 대비, 현재가,시가, 매도호가, 매수호가, 거래량, 거래대금, 전일거래량, 체결강도

    # 시초가 매매법
    # 조건

    # 1. 갭 상승 종목
    -> 현재가 > 시가
    # 2. 이평선 기준 지지선 이상

    # 3. 상승률 : 현재가 / 시가
    -> 추가 연구 필요

    # 4. 체결강도 : 특정 시점 매수 / 매도 * 100
    -> 100 초과

    :return:
    '''
    # 조건 1
    data_list = [v for v in data_list if (v[4] > v[5]) and (v[4]/ v[5] < 1.05)]
    # 조건 2
    #data_list = [v for v in data_list if (1.03 < v[4]/v[5]) and (v[4]/v[5] < 1.05)]
    # 조건 3
    data_list = [v for v in data_list if v[11] > 150]

    return data_list



def condition(stocks) :
    # code = stocks[0]['code']
    for stock in stocks :
        # 수익률 + 체결 강도 적용 필요
        if stock['rate'] > 3 or stock['rate'] < -2 :
            return True
        else :
            return False



def condition_v2(stocks) :
    # code = stocks[0]['code']
    for stock in stocks :
        # 수익률 + 체결 강도 적용 필요
        if stock['rate'] > 3 or stock['rate'] < -2 :
            return True
        else :
            return False
