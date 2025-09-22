import io

from openpyxl.reader.excel import load_workbook

from core.services.server_client import ServerClient
from read_data.services.excel_reader import read_main_sheet_excel, read_chain_of_custody, data_constructor
from write_data.services.data_writer import WordService


def read_excel(file_bytes):

    #Open the workbook from here to avoid opening it multiple times from the reading functions
    workbook = load_workbook(io.BytesIO(file_bytes), data_only=True)

    # Read the different sheets
    try:
        data_dictionary = data_constructor(workbook)

    except Exception as ex:

        print(f"Error reading the data sheets -> {ex}")

    # Close workbook to free up resources
    workbook.close()

    return data_dictionary

def write_report(data, selected_template):


     # Workflow tasks to write the data in the template

    server_client_instance = ServerClient()

    template_doc = server_client_instance.get_selected_template(selected_template)

    word_service = WordService(template_doc, data)

    #First validate the template

    if not word_service.validate_template():

        raise Exception("The template is not valid")

    #word_service.write_header_table()
    #word_service.first_page()
    #word_service.write_info_table_f_page()
    #word_service.second_part_page()

    word_service.write_first_table()



    word_service.write_objective()

    #word_service.duplicate_table_rows_after_title()

    #word_service.fill_monitoring_table_data()

    word_service.insert_normative_text()



    word_service.setup_monitoring_table()

    word_service.insert_sampling_methodology_text()


    #word_service.update_table_of_contents()


    #word_service.recreate_table_of_contents()

    word_service.save()

# Caller for the functions
def general_task(file_bytes, selected_template, ):

    data = read_excel(file_bytes)
    write_report(data, selected_template)



