from openpyxl.reader.excel import load_workbook
from rest_framework import viewsets
from rest_framework import status
from rest_framework.response import Response
from rest_framework.views import APIView
import threading
from core.threads import main_flow
import json

from core.threads.main_flow import MainFlow


class ReadFile(APIView):

    # Post method (Receives the file and the selected template for the report generation)
    def post(self ,request, *args, **kwargs):

        print("POST:", request.POST, "FILES:", request.FILES)
        # Get the file from the formData
        chain_of_custody = request.FILES.get('file')

        # Get the template data from the formData
        template_name = request.POST.get('template')



        if not chain_of_custody:
            return Response({"error": "The file was not sent"}, status=status.HTTP_400_BAD_REQUEST)

        if not template_name:
            return Response({"error": "The template was not sent"}, status=status.HTTP_400_BAD_REQUEST)


        file_bytes = chain_of_custody.read()


        # Launch a thread to process the main flow
        thread = threading.Thread(target=MainFlow().main_flow_caller , args=(file_bytes, template_name
                    ))
        thread.start()

        #Return 200 OK response
        return Response({"Success": "Data received"}, status=status.HTTP_200_OK)

