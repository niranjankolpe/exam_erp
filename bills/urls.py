from django.urls import path 
from . import views

urlpatterns = [
    path('bills',views.bills, name="bills"),
    
    path('bills/electricity', views.electricity),
    path('bills/expert_lab_peon', views.expert_lab_peon, name='expert_lab_peon'),

    path('bills/total_lab_and_staff', views.total_lab_and_staff),
   
    path('bills/int_and_ext_bills', views.int_and_ext_bills),

]