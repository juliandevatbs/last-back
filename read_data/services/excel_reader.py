from openpyxl.reader.excel import load_workbook


def read_main_sheet_excel(workbook):


    try:

        main_sheet = workbook.worksheets[0]

        #print(main_sheet)

        # Dict to storage the data
        data_client = {}
        sampling_data = {}

        # Storage client data
        data_client["client_name"] = main_sheet["B2"].value or "Not client found"
        data_client["client_contact"] = main_sheet["B6"].value or "Not client contact found"
        data_client["prepared_by"] = main_sheet["E2"].value or "No manufacturer found"


        #print(data_client)

    except Exception as ex:

        print(f"File error -> {ex}")


def read_chain_of_custody(workbook):

    return True




