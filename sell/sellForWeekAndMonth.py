import pymysql
import datetime

db = pymysql.connect(host='localhost', user='root', password='ashy1256', db='bflow_data', charset='utf8')
cursor = db.cursor()

sql = '''select product_order_number, payment_at, week, month from sell where payment_at is not null and week is null;'''

cursor.execute(sql)
rows = cursor.fetchall()

insertSql = '''INSERT INTO `bflow_data`.`sell` (product_order_number, week, month)
 VALUES (%s, %s, %s) ON DUPLICATE KEY UPDATE week = %s, month = %s'''

for row in rows:
    product_order_number = row[0]
    payment_at = row[1]
    week = payment_at.isocalendar()[1]
    month = payment_at.month

    values = (
        product_order_number, week, month, week, month
    )
    print(insertSql, values)
    cursor.execute(insertSql, values)
    db.commit()

db.close()
