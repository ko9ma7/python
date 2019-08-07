from django.urls import path
from . import views

app_name = "bflow_web"

urlpatterns = [
    path('list/', views.sell_list, name='sell_list'),
    path('create/', views.sell_create, name='sell_create'),
]