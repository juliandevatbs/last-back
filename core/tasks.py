import io
import logging

from openpyxl.reader.excel import load_workbook

from core.services.server_client import ServerClient
from intelligent_model.prompts.load_prompt import load_prompt
from intelligent_model.services.gemini_feedback import gemini_feedback
from intelligent_model.services.gemini_service import ask_gemini
from read_data.services.excel_reader import read_main_sheet_excel, read_chain_of_custody, data_constructor, \
    read_sample_information
from read_data.services.extract_text_docx import extract_text_docx
from read_data.services.readers.ExcelReaderMain import ExcelReaderMain
from read_data.services.readers.ph_reader import ph_reader
from write_data.services.docx_writer.DocxWriterMain import DocxWriterMain
from write_data.services.data_writer import WordService
from write_data.services.json_builder import JsonBuilder
from write_data.services.writer import Writer

logger = logging.getLogger(__name__)


def main_thread(file_bytes):
    """
    Función principal que integra lectura de Excel y escritura en Word
    usando tanto DocxWriterMain como Writer con JsonBuilder
    """

    # ===== LECTURA DE DATOS =====
    workbook = load_workbook(io.BytesIO(file_bytes), data_only=True)

    excel_reader_main_instance = ExcelReaderMain()
    excel_reader_main_instance.load_work_book(workbook)

    basic_data, samples_data, ph_data, ph_data_2 = excel_reader_main_instance.caller()

    # ===== ESCRITURA CON DOCXWRITERMAIN =====
    docx_write_main_instance = DocxWriterMain()
    docx_write_main_instance.load_docx()
    docx_write_main_instance.load_data(basic_data, samples_data, ph_data, ph_data_2)
    docx_write_main_instance.caller()

    # ===== CONFIGURACIÓN JSON Y WRITER =====
    # Create a JsonBuilder instance for call all its functions
    json_builder_instance = JsonBuilder(file_bytes)

    # Load json file
    json_builder_instance.load_json()

    # Write excel data to json file
    json_builder_instance.update_json()

    # Create a Writer instance for call all its functions
    writer_instance = Writer()

    # Load the updated json data to the writer
    writer_instance.load_json_config()

    # Open word template to write
    writer_instance.word_template = docx_write_main_instance.docx

    # Main writer calls all of writer functions in a specific order
    writer_instance.main_writer()

    # ===== GUARDAR DOCUMENTO FINAL =====
    docx_write_main_instance.save_doc()

    print("=== Documento generado exitosamente ===")


def get_feedback_from_gemini(file_bytes):
    """Obtiene feedback de Gemini para un documento"""
    try:
        document_content = extract_text_docx(file_bytes)

        if not document_content:
            print("Could not extract content from the document")
            return None

        logger.info(f"Documento extraido {len(document_content)} caracteres")

        prompt = load_prompt("docx_feedback", documento_text=document_content)
        feedback = ask_gemini(prompt)

        if feedback:
            logger.info("Feedback generado exitosamente")
            print("=== FEEDBACK GENERADO ===")
            print(feedback)
            return feedback
        else:
            logger.error("No se recibio feedback de gemini")
            return None

    except FileNotFoundError as ex:
        logger.error(f"Archivo de prompt no encontrado: {ex}")
        return None
    except Exception as e:
        logger.error(f"Error procesando documento: {e}")
        return None


# Función principal para llamar desde el endpoint
def general_task(file_bytes):
    """
    Función principal que orquesta todo el proceso
    """
    try:
        main_thread(file_bytes)
        print("Proceso completado exitosamente")

    except Exception as e:
        logger.error(f"Error en general_task: {e}")
        print(f"Error en general_task: {e}")
        raise