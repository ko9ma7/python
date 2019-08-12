from django.db import models

# Create your models here.


class Calculate(models.Model):
    product_order_number = models.CharField(unique=True, max_length=50, blank=True, null=True)
    order_number = models.CharField(max_length=50, blank=True, null=True)
    channel_order_number = models.CharField(max_length=50, blank=True, null=True)
    order_state = models.CharField(max_length=50, blank=True, null=True)
    delivery_complete_at = models.DateField(blank=True, null=True)
    provide_name = models.CharField(max_length=50, blank=True, null=True)
    channel = models.CharField(max_length=50, blank=True, null=True)
    quantity = models.IntegerField(blank=True, null=True)
    order_amount = models.IntegerField(blank=True, null=True)
    fees = models.FloatField(blank=True, null=True)
    calculate = models.IntegerField(blank=True, null=True)
    channel_calculate = models.IntegerField(blank=True, null=True)
    complete_at = models.DateField(blank=True, null=True)
    matching_at = models.DateField(blank=True, null=True)
    week = models.IntegerField(blank=True, null=True)
    month = models.IntegerField(blank=True, null=True)
    margin = models.IntegerField(blank=True, null=True)


           

    def __str__(self):
        return self.product_order_number


class Sell(models.Model):
    product_order_number = models.CharField(unique=True, max_length=50, blank=True, null=True)
    order_number = models.CharField(max_length=50, blank=True, null=True)
    payment_at = models.DateField(blank=True, null=True)
    order_state = models.CharField(max_length=50, blank=True, null=True)
    claim = models.CharField(max_length=50, blank=True, null=True)
    provide_name = models.CharField(max_length=50, blank=True, null=True)
    product_name = models.CharField(max_length=255, blank=True, null=True)
    product_option = models.CharField(max_length=255, blank=True, null=True)
    channel = models.CharField(max_length=50, blank=True, null=True)
    product_number = models.CharField(max_length=50, blank=True, null=True)
    product_amount = models.IntegerField(blank=True, null=True)
    option_amount = models.IntegerField(blank=True, null=True)
    seller_discount = models.IntegerField(blank=True, null=True)
    quantity = models.IntegerField(blank=True, null=True)
    total_amount = models.IntegerField(blank=True, null=True)
    delivery_at = models.DateField(blank=True, null=True)
    delivery_complete = models.DateField(blank=True, null=True)
    order_complete_at = models.DateField(blank=True, null=True)
    auto_complete_at = models.DateField(blank=True, null=True)
    category_number = models.CharField(max_length=50, blank=True, null=True)
    buyer_email = models.CharField(max_length=50, blank=True, null=True)
    buyer_gender = models.CharField(max_length=50, blank=True, null=True)
    buyer_age = models.CharField(max_length=50, blank=True, null=True)
    crawler = models.CharField(max_length=50, blank=True, null=True)
    calculate = models.IntegerField(blank=True, null=True)
    calculate_channel = models.IntegerField(blank=True, null=True)
    margin = models.IntegerField(blank=True, null=True)
    profit = models.FloatField(blank=True, null=True)
    month = models.IntegerField(blank=True, null=True)
    week = models.IntegerField(blank=True, null=True)

    def __str__(self):
        return self.product_order_number


class Sellstatistics(models.Model):
    date = models.DateField(unique=True, blank=True, null=True)
    week = models.IntegerField(blank=True, null=True)
    month = models.IntegerField(blank=True, null=True)
    brich_total_amount = models.IntegerField(blank=True, null=True)
    brich_qty = models.IntegerField(blank=True, null=True)
    brich_ct = models.IntegerField(blank=True, null=True)  # Field name made lowercase.
    brich_margin = models.IntegerField(blank=True, null=True)
    brich_profit = models.FloatField(blank=True, null=True)
    gmarket_total_amount = models.IntegerField(blank=True, null=True)
    gmarket_qty = models.IntegerField(blank=True, null=True)
    gmarket_ct = models.IntegerField(blank=True, null=True)  # Field name made lowercase.
    gmarket_margin = models.IntegerField(blank=True, null=True)
    gmarket_profit = models.FloatField(blank=True, null=True)
    auction_total_amount = models.IntegerField(blank=True, null=True)
    auction_qty = models.IntegerField(blank=True, null=True)
    auction_ct = models.IntegerField(blank=True, null=True)  # Field name made lowercase.
    auction_margin = models.IntegerField(blank=True, null=True)
    auction_profit = models.FloatField(blank=True, null=True)
    st_total_amount = models.IntegerField(blank=True, null=True)  # Field renamed because it wasn't a valid Python identifier.
    st_qty = models.IntegerField(blank=True, null=True)  # Field renamed because it wasn't a valid Python identifier.
    st_ct = models.IntegerField(blank=True, null=True)  # Field name made lowercase. Field renamed because it wasn't a valid Python identifier.
    st_margin = models.IntegerField(blank=True, null=True)  # Field renamed because it wasn't a valid Python identifier.
    st_profit = models.FloatField(blank=True, null=True)  # Field renamed because it wasn't a valid Python identifier.
    wemakeprice_total_amount = models.IntegerField(blank=True, null=True)
    wemakeprice_qty = models.IntegerField(blank=True, null=True)
    wemakeprice_ct = models.IntegerField(blank=True, null=True)  # Field name made lowercase.
    wemakeprice_margin = models.IntegerField(blank=True, null=True)
    wemakeprice_profit = models.FloatField(blank=True, null=True)
    interpark_total_amount = models.IntegerField(blank=True, null=True)
    interpark_qty = models.IntegerField(blank=True, null=True)
    interpark_ct = models.IntegerField(blank=True, null=True)  # Field name made lowercase.
    interpark_margin = models.IntegerField(blank=True, null=True)
    interpark_profit = models.FloatField(blank=True, null=True)
    coupang_total_amount = models.IntegerField(blank=True, null=True)
    coupang_qty = models.IntegerField(blank=True, null=True)
    coupang_ct = models.IntegerField(blank=True, null=True)  # Field name made lowercase.
    coupang_margin = models.IntegerField(blank=True, null=True)
    coupang_profit = models.FloatField(blank=True, null=True)
    ssg_total_amount = models.IntegerField(blank=True, null=True)
    ssg_qty = models.IntegerField(blank=True, null=True)
    ssg_ct = models.IntegerField(blank=True, null=True)  # Field name made lowercase.
    ssg_margin = models.IntegerField(blank=True, null=True)
    ssg_profit = models.FloatField(blank=True, null=True)
    g9_total_amount = models.IntegerField(blank=True, null=True)
    g9_qty = models.IntegerField(blank=True, null=True)
    g9_ct = models.IntegerField(blank=True, null=True)  # Field name made lowercase.
    g9_margin = models.IntegerField(blank=True, null=True)
    g9_profit = models.FloatField(blank=True, null=True)

    def __str__(self):
        return self.date
