from django.urls import path
from . import views

app_name = "bflow_web"

urlpatterns = [
    path('', views.sell_list, name='sell_list'),
]