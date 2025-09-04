from django.contrib import admin
from django.urls import path

from read_data.views import ReadFile

urlpatterns = [
    path('upload/', ReadFile.as_view(), name="upload-file"),
]
