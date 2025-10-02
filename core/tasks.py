import io
import logging

from openpyxl.reader.excel import load_workbook

from core.services.server_client import ServerClient
from intelligent_model.prompts.load_prompt import load_prompt
from intelligent_model.services.gemini_feedback import gemini_feedback
from intelligent_model.services.gemini_service import ask_gemini
from read_data.services.excel_reader import read_main_sheet_excel, read_chain_of_custody, data_constructor
from read_data.services.extract_text_docx import extract_text_docx
from write_data.services.data_writer import WordService
from write_data.services.json_builder import JsonBuilder
from write_data.services.writer import Writer

logger = logging.getLogger(__name__)

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

    word_service.write_first_table()

    word_service.write_objective()

    word_service.insert_normative_text()

    word_service.setup_monitoring_table()

    word_service.insert_sampling_methodology_text()

    word_service.save()


def get_feedback_from_gemini(file_bytes):
    try:
        # Extract the file content
        document_content = extract_text_docx(file_bytes)

        if not document_content:
            print("Could not extract content from the document")
            return None

        logger.info(f"Documento extraido {len(document_content)} caracteres")

        # Build the prompt
        prompt = load_prompt("docx_feedback", documento_text=document_content)

        # Send the file content to gemini service
        feedback = ask_gemini(prompt)

        if feedback:
            logger.info("Feedback generado exitosamente")
            print("=== FEEDBACK GENERADO ===")
            print(feedback)
            return feedback
        else:
            logger.error("No se recibio feedback de gemini")
            print("No se recibio feedback de gemini")
            return None

    except FileNotFoundError as ex:
        logger.error(f"Archivo de prompt no encontrado: {ex}")
        print(f"Archivo de prompt no encontrado: {ex}")
        return None

    except Exception as e:
        logger.error(f"Error procesando documento: {e}")
        print(f"Error procesando documento: {e}")
        return None


# Main Caller for the functions
def general_task(file_bytes,selected_template, selected_options, selected_reporter):


    data = read_excel(file_bytes)

    # Create a Writer instance for call all its functions
    writer_instance = Writer()

    # Create a JsonBuilder instance for call all its functions
    json_builder_instance = JsonBuilder(file_bytes)

    # Load json file
    json_builder_instance.load_json()

    # Write excel data to json file
    json_builder_instance.update_json()

    # Load the updated json data to the writer
    writer_instance.load_json_config()

    # Open word template to write
    writer_instance.load_word_template()

    # Main writer calls all of writer functions in a specific order
    writer_instance.main_writer()

    # Clen the json config to reuse with other report
    json_builder_instance.clean_json()

    # Save the document with all changes
    writer_instance.save_document("templates/final.docx")

    #write_report(data, selected_template)



