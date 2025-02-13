from django.urls import path 
from . import views

urlpatterns = [
    path('statements',views.statement),
    path('statements/labstaff',views.labstaff, name="labstaff"),
    path('statements/internal_external_bill', views.internal_external_bill, name="internal_external")
]