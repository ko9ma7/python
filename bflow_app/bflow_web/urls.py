from django.urls import path
from . import views

app_name = "bflow_web"

urlpatterns = [
    path('sell/list/', views.sell_list, name='sell_list'),
    path('sell/create/', views.sell_create, name='sell_create'),
    path('calculate/list/', views.calculate_list, name='calculate_list'),
    path('calculate/create/', views.calculate_create, name='calculate_create'),
]