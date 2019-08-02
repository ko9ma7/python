import pymysql
import datetime


db = pymysql.connect(host='localhost', user='root', password='ashy1256', db='bflow_data', charset='utf8')
cursor = db.cursor()

sql = '''
 select sell.payment_at, sell.product_order_number, sell.total_amount,
 calculate.calculate, calculate.channel_calculate, calculate.margin 
 from sell
 inner join calculate on sell.product_order_number = calculate.product_order_number;
'''

cursor.execute(sql)
rows = cursor.fetchall()

insertSql = '''INSERT INTO `bflow_data`.`sell` (
payment_at,
product_order_number,
total_amount,
calculate,
calculate_channel,
margin,
profit
) VALUES (%s, %s, %s, %s, %s, %s, %s)
ON DUPLICATE KEY UPDATE calculate = %s, calculate_channel = %s, margin = %s, profit = %s
'''

for row in rows:
    payment_at = row[0]
    product_order_number = row[1]
    total_amount = row[2]
    calculate = row[3]
    calculate_channel = row[4]
    margin = row[5]
    profit = round(margin/total_amount*100, 2)
    print(payment_at, product_order_number, total_amount, calculate, calculate_channel, margin, profit)

    values = (
        payment_at,
        product_order_number,
        total_amount,
        calculate,
        calculate_channel,
        margin,
        profit,
        calculate,
        calculate_channel,
        margin,
        profit
    )

    cursor.execute(insertSql, values)
    db.commit()

db.close()
