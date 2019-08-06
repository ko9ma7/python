from django.urls import path
from . import views

urlpatterns = [
    path('', views.sell_list, name='sell_list'),
]