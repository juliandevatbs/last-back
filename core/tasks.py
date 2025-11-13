import io
import json
import logging
import os

from openpyxl.reader.excel import load_workbook

from core.exceptions import NotJsonConfigFound
from read_data.services.readers.ExcelReaderMain import ExcelReaderMain
from write_data.services.json_builder import JsonBuilder
from write_data.services.writer import Writer

logger = logging.getLogger(__name__)


class Tasks:

    # This class contains functions to main flow

    def __init__(self):

        # Directory of the json folder
        self.JSON_FOLDER_URL = 'fields_config/'



    # This function validates that the selected template has a configuration file
    def search_template_config(self, template_name:str):

        template_name += ".json"

        if template_name not in os.listdir(self.JSON_FOLDER_URL):

            raise NotJsonConfigFound(f"The configuration file for the {template_name} template does not exist")

        config_path = os.path.join(self.JSON_FOLDER_URL, template_name )

        try:

            with open(config_path, 'r', encoding='utf-8') as file:

                json_content = json.load(file)

            logger.info(f"Config loaded succesfully {template_name}")


            return json_content

        except json.JSONDecodeError as e:

            logger.error(f"Invalid json in {template_name}: {e}")
            raise

        except Exception as e:

            logger.error(f"Error reading {template_name} config: {e}")
            raise

    # This function cleans the last data from the json config
    def clean_json(self, json_to_clean):


        if json_to_clean:

            if isinstance(json_to_clean, dict):
                return {key: self.clean_json(value) for key, value in json_to_clean.items()}

            elif isinstance(json_to_clean, list):
                return [self.clean_json(item) for item in json_to_clean]

            elif isinstance(json_to_clean, str):
                return ""

            elif isinstance(json_to_clean, bool):
                return False

            elif isinstance(json_to_clean, (int, float)):
                return 0

            else:

                return None

        return json_to_clean











def process_excel_to_word(file_bytes, config_name, template_name, output_name):
    try:
        workbook = load_workbook(io.BytesIO(file_bytes), data_only=True)

        reader = ExcelReaderMain()
        reader.load_work_book(workbook)
        reader.load_json_data(config_name)


        basic_data, samples_data, specific_data = reader.caller()
        workbook.close()


        # Json class instance
        json_instance = JsonBuilder()
        json_instance.load_json("PLANTILLA_CPF_CUPIAGUA_ACEITOSAS_ARI_ACBB")
        json_instance.update_json_normal_values(basic_data)
        json_instance.update_json_samples(samples_data)
        json_instance.save_json("PLANTILLA_CPF_CUPIAGUA_ACEITOSAS_ARI_ACBB")

        writer = Writer(
            config_path=f"fields_config/{config_name}.json",
            template_path=f"templates/{template_name}.docx"
        )

        writer.load_word_template()
        writer.load_json_config()
        writer.write_data(basic_data, samples_data, specific_data)
        writer.save_document(f"templates/{output_name}rty.docx")

        return True

    except Exception as e:
        print(f"Error en proceso: {e}")
        return False