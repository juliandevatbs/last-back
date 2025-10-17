from django.shortcuts import render
from rest_framework import generics
from rest_framework.decorators import api_view
from rest_framework import generics

from .models import Reporter
from .serializer import MetricsDataSerializer
from rest_framework.response import Response


@api_view(['GET'])
def get_reporters(request):

    queryset = Reporter.objects.all()
    serializer_class = MetricsDataSerializer(queryset, many=True)

    print(serializer_class.data)
    return Response(serializer_class.data)

