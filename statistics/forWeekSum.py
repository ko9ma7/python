import pymysql
import datetime
from openpyxl import Workbook
from openpyxl.styles import Border, Side

# wb = Workbook()

# ws = wb.active()

db = pymysql.connect(host='localhost', user='root', password='ashy1256', db='bflow_data', charset='utf8')
cursor = db.cursor()

border = Border(
left=Side(border_style='thin', color='000000'),
right=Side(border_style='thin', color='000000'),
top=Side(border_style='thin', color='000000'),
bottom=Side(border_style='thin', color='000000'),
)

def noZerodiv(a, b):
    if a != None and b != None:
        return round(b/a, 2)
    else:
        return None
    

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


    for row in rows:
        week = row[0]
        weekstr = datetime.datetime.strftime(row[1], '%Y-%m-%d') + "~" + datetime.datetime.strftime(row[2], '%Y-%m-%d')
        brich_total_amount = row[3]
        brich_qty = row[4]
        brich_CT = noZerodiv(brich_qty, brich_total_amount)
        brich_margin = row[5]
        brich_profit = noZerodiv(brich_total_amount, brich_margin)
        gmarket_total_amount = row[6]
        gmarket_qty = row[7]
        gmarket_CT = noZerodiv(gmarket_qty, gmarket_total_amount)
        gmarket_margin = row[8]
        gmarket_profit = noZerodiv(gmarket_total_amount, gmarket_margin)
        auction_total_amount = row[9]
        auction_qty = row[10]
        auction_CT = noZerodiv(auction_qty, auction_total_amount)
        auction_margin = row[11]
        auction_profit = noZerodiv(auction_total_amount, auction_margin)
        st_total_amount = row[12]
        st_qty = row[13]
        st_CT = noZerodiv(st_qty, st_total_amount)
        st_margin = row[14]
        st_profit = noZerodiv(st_total_amount, st_margin)
        wemakeprice_total_amount = row[15] 
        wemakeprice_qty = row[16] 
        wemakeprice_CT = noZerodiv(wemakeprice_qty, wemakeprice_total_amount)
        wemakeprice_margin = row[17]
        wemakeprice_profit = noZerodiv(wemakeprice_total_amount, wemakeprice_margin)
        interpark_total_amount = row[18]
        interpark_qty = row[19]
        interpark_CT = noZerodiv(interpark_qty, interpark_total_amount)
        interpark_margin = row[20]
        interpark_profit = noZerodiv(interpark_total_amount, interpark_margin)
        coupnag_total_amount = row[21]
        coupnag_qty = row[22]
        coupnag_CT = noZerodiv(coupnag_qty, coupnag_total_amount)
        coupnag_margin = row[23]
        coupnag_profit = noZerodiv(coupnag_total_amount, coupnag_margin)
        ssg_total_amount = row[24]
        ssg_qty = row[25]
        ssg_CT = noZerodiv(ssg_qty, ssg_total_amount)
        ssg_margin = row[26]
        ssg_profit = noZerodiv(ssg_total_amount, ssg_margin)
        g9_total_amount = row[27]
        g9_qty = row[28]
        g9_CT = noZerodiv(g9_qty, g9_total_amount)
        g9_margin = row[29]
        g9_profit = noZerodiv(g9_total_amount, g9_margin)

        

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



