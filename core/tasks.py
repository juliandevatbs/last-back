import io

from openpyxl.reader.excel import load_workbook

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

def write_report(data):

     # Workflow tasks to write the data in the template
    template_path = "C:\\Code\\automatizacion_informes\\BackEnd\\templates\\ecopetrol\\ECOPETROL_TEMPLATE.docx"

    word_service = WordService(template_path, data)
    word_service.load_template()
    word_service.write_header_table()
    word_service.first_page()
    word_service.write_info_table_f_page()
    word_service.second_part_page()

    word_service.save(template_path)

# Caller for the functions
def general_task(file_bytes):

    data = read_excel(file_bytes)
    write_report(data)



