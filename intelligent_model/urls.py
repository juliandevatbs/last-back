from django.contrib import admin
from django.urls import path, include

from intelligent_model.views import FeedBackProvider

urlpatterns = [

    path('feedback-generator/', FeedBackProvider.as_view(), name='feedback_provider')

]