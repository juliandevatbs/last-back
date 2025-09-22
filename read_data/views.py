from openpyxl.reader.excel import load_workbook
from rest_framework import viewsets
from rest_framework import status
from rest_framework.response import Response
from rest_framework.views import APIView
import threading
from core.tasks import read_excel, general_task
from read_data.services.excel_reader import read_main_sheet_excel, read_chain_of_custody

class ReadFile(APIView):

    # Post method (Receives the file and the selected template for the report generation)
    def post(self ,request, *args, **kwargs):

        # Get the file from the formData
        uploaded_file = request.FILES.get('file')

        # Get the template data from the formData
        selected_template = request.POST.get('template')

        # Get the options from the formData
        selected_options = request.POST.get('options')

        #print(f"SELECTED TEMPLATE {selected_template}")
        #print(f"SELECTED OPTIONS {selected_options}")

        if not uploaded_file:
            return Response({"error": "The file was not sent"}, status=status.HTTP_400_BAD_REQUEST)

        if not selected_template:
            return Response({"error": "The template was not sent"}, status=status.HTTP_400_BAD_REQUEST)

        file_bytes = uploaded_file.read()

        # Launch a thread to process the excel without block the response
        thread = threading.Thread(target=general_task , args=(file_bytes, selected_template, ))
        thread.start()

        #Return 200 OK response
        return Response({"Success": "File opened succesfully"}, status=status.HTTP_200_OK)



