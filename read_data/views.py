from django.shortcuts import render
from django.template.context_processors import request
from rest_framework import viewsets
from rest_framework import status
from rest_framework.response import Response
from rest_framework.views import APIView


class UploadFile(APIView):

    # Post method
    def post(self ,request, *args, **kwargs):

        #print(request.data)
        #print(request.FILES)

        # The file comes in form-data
        received_file = request.FILES.get("file")

        if not received_file:
            # Return a bad error response
            return Response({ "error": "The file was not sent" }, status=status.HTTP_400_BAD_REQUEST)

        # If the file was uploaded succesfully
        return Response({ "success": "File uploaded succesfully"}, status = status.HTTP_200_OK)



