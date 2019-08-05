import pymysql
import datetime
from openpyxl import Workbook

# wb = Workbook()

# ws = wb.active()

db = pymysql.connect(host='localhost', user='root', password='ashy1256', db='bflow_data', charset='utf8')
cursor = db.cursor()

for i in range(1, 30):
    sql = f'''
    SELECT 
    week,
    min(date),
    max(date),
    sum(brich_total_amount),
    sum(brich_qty),
    sum(brich_margin),
    sum(gmarket_total_amount),
    sum(gmarket_qty),
    sum(gmarket_margin),
    sum(auction_total_amount),
    sum(auction_qty),
    sum(auction_margin),
    sum(11st_total_amount),
    sum(11st_qty),
    sum(11st_margin),
    sum(wemakeprice_total_amount),
    sum(wemakeprice_qty),
    sum(wemakeprice_margin),
    sum(interpark_total_amount),
    sum(interpark_qty),
    sum(interpark_margin),
    sum(coupang_total_amount),
    sum(coupang_qty),
    sum(coupang_margin),
    sum(ssg_total_amount),
    sum(ssg_qty),
    sum(ssg_margin),
    sum(g9_total_amount),
    sum(g9_qty),
    sum(g9_margin)
    FROM sellstatistics
    WHERE week = {i}
    GROUP BY week
    '''

    cursor.execute(sql)
    rows = cursor.fetchall()

    no = 0

    for row in rows:
        week = row[0]
        weekstr = datetime.datetime.strftime(row[1], '%Y-%m-%d') + "~" + datetime.datetime.strftime(row[2], '%Y-%m-%d')
        brich_total_amount = row[3]
        brich_qty = row[4]
        brich_CT = round(brich_total_amount/brich_qty,0)
        brich_margin = row[5] or 0
        brich_profit = round(brich_margin/brich_total_amount,2)
        gmarket_total_amount = row[6]
        gmarket_qty = row[7]
        gmarket_CT = round(gmarket_total_amount/gmarket_qty,0)
        gmarket_margin = row[8] or 0
        gmarket_profit = round(gmarket_margin/gmarket_total_amount,2)
        auction_total_amount = row[9]
        auction_qty = row[10]
        auction_CT = round(auction_total_amount/auction_qty,0)
        auction_margin = row[11] or 0
        auction_profit = round(auction_margin/auction_total_amount,2)
        st_total_amount = row[12]
        st_qty = row[13]
        st_CT = round(st_total_amount/st_qty,0)
        st_margin = row[14] or 0
        st_profit = round(st_margin/st_total_amount,2)
        wemakeprice_total_amount = row[15] 
        wemakeprice_qty = row[16] 
        wemakeprice_CT = round(wemakeprice_total_amount/wemakeprice_qty,0) or 0
        wemakeprice_margin = row[17] or 0
        wemakeprice_profit = round(wemakeprice_margin/wemakeprice_total_amount,2) or 0
        interpark_total_amount = row[18]
        interpark_qty = row[19]
        interpark_CT = round(interpark_total_amount/interpark_qty,0)
        interpark_margin = row[20] or 0
        interpark_profit = round(interpark_margin/interpark_total_amount,2)
        coupnag_total_amount = row[21]
        coupnag_qty = row[22]
        coupnag_CT = round(coupnag_total_amount/coupnag_qty,0)
        coupnag_margin = row[23] or 0
        coupnag_profit = round(coupnag_margin/coupnag_total_amount,2)
        ssg_total_amount = row[24]
        ssg_qty = row[25]
        ssg_CT = round(ssg_total_amount/ssg_qty,0)
        ssg_margin = row[26] or 0
        ssg_profit = round(ssg_margin/ssg_total_amount,2)
        g9_total_amount = row[27]
        g9_qty = row[28]
        g9_CT = round(g9_total_amount/g9_qty,0)
        g9_margin = row[29] or 0
        g9_profit = round(g9_margin/g9_total_amount,2)

        print(
            week,
            weekstr,
            brich_total_amount,
            brich_qty,
            brich_CT,
            brich_margin,
            brich_profit,
            gmarket_total_amount,
            gmarket_qty,
            gmarket_CT,
            gmarket_margin,
            gmarket_profit,
            auction_total_amount,
            auction_qty,
            auction_CT,
            auction_margin,
            auction_profit,
            st_total_amount,
            st_qty,
            st_CT,
            st_margin,
            st_profit,
            wemakeprice_total_amount,
            wemakeprice_qty,
            wemakeprice_CT,
            wemakeprice_margin,
            wemakeprice_profit,
            interpark_total_amount,
            interpark_qty,
            interpark_CT,
            interpark_margin,
            interpark_profit,
            coupnag_total_amount,
            coupnag_qty,
            coupnag_CT,
            coupnag_margin,
            coupnag_profit,
            ssg_total_amount,
            ssg_qty,
            ssg_CT,
            ssg_margin,
            ssg_profit,
            g9_total_amount,
            g9_qty,
            g9_CT,
            g9_margin,
        )



