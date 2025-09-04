from openpyxl.reader.excel import load_workbook
from rest_framework import viewsets
from rest_framework import status
from rest_framework.response import Response
from rest_framework.views import APIView
from read_data.services.excel_reader import read_main_sheet_excel, read_chain_of_custody

class ReadFile(APIView):

    # Post method
    def post(self ,request, *args, **kwargs):

        uploaded_file = request.FILES.get('file')
        #print(uploaded_file.name)

        if not uploaded_file:

            return Response({"error": "The file was not sent"}, status=status.HTTP_400_BAD_REQUEST)

        #Open the workbook from here to avoid opening it multiple times from the reading functions
        workbook = load_workbook(uploaded_file)

        # Read the different sheets
        try:
            read_main_sheet_excel(workbook)

        except Exception as ex:

            print(f"Error reading the basic data sheet -> {ex}")

        try:
            read_chain_of_custody(workbook)

        except Exception as ex:

            print(f"Error reading the chain of custody -> {ex}")


        #Close workbook to free up resources
        workbook.close()

        #Return 200 response
        return Response({"Success": "File opened succesfully"}, status=status.HTTP_200_OK)



