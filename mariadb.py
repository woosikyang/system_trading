import pymysql
import configs



conn = pymysql.connect(host=configs.ip,
                       user='root',
                       password=configs.password,
                       charset='utf8',
                       port=3306)

cur = conn.cursor()

sql = "create table if not exists userTable (id char(4), userName char(10))"

cur.execute(sql)
