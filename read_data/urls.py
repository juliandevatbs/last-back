from django.contrib import admin
from django.urls import path

from read_data.views import UploadFile

urlpatterns = [
    path('upload/', UploadFile.as_view(), name="upload-file"),
]
