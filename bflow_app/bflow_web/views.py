from django.shortcuts import render
from .models import Sell
import openpyxl
import datetime

def replacedate(text):
    if text is None:
        return None
    else:
        text = text[0:10]
        return text


def replacenone(text):
    if text is None:
        return None
    else:
        text = str(text)
        return text


def replaceint(text):
    if text is None:
        return None
    else:
        text = int(text)
        return text

def sell_list(request):
    if "GET" == request.method:
        return render(request, 'bflow_web/sell_list.html', {})
    else:
        excel_file = request.FILES["excel_file"]

        wb = openpyxl.load_workbook(excel_file)

        ws = wb.active
        print(ws)

        excel_data = list()
        iter_rows = iter(ws.rows)
        next(iter_rows)

        
        for row in iter_rows:
            row_data = list()
            
            
            sell = Sell.objects.update_or_create(
                product_order_number = replacenone(row[0].value),
                order_number = replacenone(row[1].value),
                payment_at = replacedate(row[2].value),
                order_state = replacenone(row[3].value),
                claim = replacenone(row[4].value),
                provide_name = replacenone(row[5].value),
                product_name = replacenone(row[6].value),
                product_option = replacenone(row[7].value),
                channel = replacenone(row[8].value),
                product_number = replacenone(row[18].value),
                product_amount = replaceint(row[19].value),
                option_amount = replaceint(row[20].value),
                seller_discount = replaceint(row[21].value),
                quantity = replaceint(row[22].value),
                total_amount = replaceint(row[23].value),
                delivery_at = replacedate(row[24].value),
                delivery_complete = replacedate(row[25].value),
                order_complete_at = replacedate(row[26].value),
                auto_complete_at = replacedate(row[27].value),
                category_number = replacenone(row[39].value),
                buyer_email = replacenone(row[40].value),
                buyer_gender = replacenone(row[41].value),
                buyer_age = replacenone(row[42].value),
                crawler = replacenone(row[43].value),
            )
            


            for cell in row:
                row_data.append(str(cell.value))
            excel_data.append(row_data)
        
        print (excel_data[0][1])
        return render(request, 'bflow_web/sell_list.html', {"excel_data": excel_data})


# Create your views here.
