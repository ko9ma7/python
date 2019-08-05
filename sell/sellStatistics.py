import pymysql
import datetime

db = pymysql.connect(host='localhost', user='root', password='ashy1256', db='bflow_data', charset='utf8')
cursor = db.cursor()

# brich, gmarket, auction, 11st, wemakeprice, interpark, coupang, ssg, g9

channel_name = ['brich', 'gmarket', 'auction', '11st', 'wemakeprice', 'interpark', 'coupang', 'ssg', 'g9']

for name in channel_name:
    sql = f'''select payment_at, channel, sum(total_amount), sum(quantity), sum(calculate), sum(calculate_channel)
    from sell 
    where not order_state in ('결제취소')
     and channel = '{name}'
    and payment_at is not null 
    group by payment_at;'''
    
    cursor.execute(sql)
    rows = cursor.fetchall()

    print(sql)

    insertSql = f'''insert into `bflow_data`.`sellstatistics` 
    (
    date,
    week,
    month,
    {name}_total_amount,
    {name}_qty,
    {name}_ct,
    {name}_margin,
    {name}_profit
    )
    values (%s, %s, %s, %s, %s, %s, %s, %s) ON DUPLICATE KEY UPDATE
    week = %s,
    month = %s,
    {name}_total_amount = %s,
    {name}_qty = %s,
    {name}_ct = %s,
    {name}_margin = %s,
    {name}_profit = %s '''

    for row in rows:
        date = row[0]
        week = date.isocalendar()[1]
        month = date.month
        total_amount = row[2]
        qty = row[3]
        ct = round(total_amount/qty, 0)
        calculate = row[4] or 0
        calculate_channel = row[5] or 0
        margin = calculate_channel - calculate
        profit = round(margin/total_amount, 2) or 0

        values = (
            date,
            week,
            month,
            total_amount,
            qty,
            ct,
            margin,
            profit,
            week,
            month,
            total_amount,
            qty,
            ct,
            margin,
            profit
        )

        print(insertSql, values)
        cursor.execute(insertSql, values)
        db.commit()

db.close()

