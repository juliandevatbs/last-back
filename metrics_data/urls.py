from django.urls import path
from . import views

urlpatterns = [


    path('api/reporters/', views.get_reporters, name='reporters')


]