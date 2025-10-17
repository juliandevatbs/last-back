from openpyxl.reader.excel import load_workbook
from rest_framework import viewsets
from rest_framework import status
from rest_framework.response import Response
from rest_framework.views import APIView
import threading
from core.tasks import general_task, main_thread
from read_data.services.excel_reader import read_main_sheet_excel, read_chain_of_custody
import json

class ReadFile(APIView):

    # Post method (Receives the file and the selected template for the report generation)
    def post(self ,request, *args, **kwargs):

        # Get the file from the formData
        uploaded_file = request.FILES.get('file')

        # Get the template data from the formData
        selected_template_str = request.POST.get('template')
        selected_template = json.loads(selected_template_str) if selected_template_str else None

        # Get the options from the formData
        selected_options = request.POST.get('options')

        # Get the reporter from the formData
        reporter = request.POST.get('reporter')


        if not uploaded_file:
            return Response({"error": "The file was not sent"}, status=status.HTTP_400_BAD_REQUEST)

        if not selected_template:
            return Response({"error": "The template was not sent"}, status=status.HTTP_400_BAD_REQUEST)

        if not reporter:
            return Response({"error": "The reporter infor was not sent"}, status=status.HTTP_400_BAD_REQUEST)

        file_bytes = uploaded_file.read()

        # Launch a thread to process the excel without block the response
        thread = threading.Thread(target=general_task , args=(file_bytes, ))
        thread.start()

        #Return 200 OK response
        return Response({"Success": "File opened succesfully"}, status=status.HTTP_200_OK)

