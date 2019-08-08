from django.urls import path
from . import views

app_name = "bflow_web"

urlpatterns = [
    path('sell/list/', views.sell_list, name='sell_list'),
    path('sell/create/', views.sell_create, name='sell_create'),
    path('calaulate/list/', views.sell_list, name='calaulate_list'),
    path('calaulate/create/', views.calaulate_create, name='calaulate_create'),
]