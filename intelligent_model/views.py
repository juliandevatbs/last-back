import threading

from django.shortcuts import render
from rest_framework import status
from rest_framework.response import Response
from rest_framework.views import APIView

from core.tasks import get_feedback_from_gemini


class FeedBackProvider(APIView):

    def post(self, request, *args, **kwargs):


        # Get the docx file from the formdata
        docx_file = request.FILES.get('file')


        if not docx_file:
            # Return a bad request
            return Response({"error": "Word file not provided"}, status=status.HTTP_400_BAD_REQUEST)

        file_bytes = docx_file.read()

        # Launch a thread to process the word file without block the response
        thread = threading.Thread(target=get_feedback_from_gemini, args=(file_bytes, ))
        thread.start()


        # If everything goes well, return status 200
        return Response({"data": "Reading the word"}, status=status.HTTP_200_OK)