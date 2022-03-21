import pymysql
import configs


'''
RISE : 급상승 종목 속성 저장 테이블


'''


conn = pymysql.connect(host=configs.ip,
                       user='root',
                       password=configs.password,
                       charset='utf8',
                       database='uprising',
                       port=3306)

cur = conn.cursor()
# 종목코드, 시간, 대비부호, 대비, 현재가, 시가, 매도호가, 매수호가, 거래량, 거래대금, 전일거래량, 체결강도
sql = "create table if not exists RISE(code varchar(10), cur_time char(5), daebi_buho char(1), daebi varchar(5)," \
      " cur_price int, start_price int, ask_price int, sell_price int, amout int, money_amount int, previous_amount int, power float)"


#sql 실행
cur.execute(sql)

#저장
conn.commit()

#종료
conn.close()