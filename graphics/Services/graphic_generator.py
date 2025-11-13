import json
import logging
import os
import traceback

from core.exceptions import NotJsonConfigFound
from graphics.Services.read.read_graphics_data import ReadGraphicsData
from graphics.Services.write.write_insitu_data import WriteInsituData

logger = logging.getLogger(__name__)


class GraphicGenerator:

    def __init__(self):


        self.graphics_config = None
        self.chain_of_custody = None

        #Constants
        self.JSON_FOLDER_URL = "fields_config/"

        #Instances
        self.reader_graphics_data = ReadGraphicsData()
        self.writer_graphics_data = WriteInsituData()


    def get_chain_of_custody(self, chain_of_custody):


        if not chain_of_custody:

            logger.error("Chain of custody not loaded")

            return

        self.chain_of_custody = chain_of_custody




    def search_graphics_config(self, report_name: str):

        report_name = "GRAFICAS_" + report_name + ".json"

        if report_name not in os.listdir(self.JSON_FOLDER_URL):
            raise NotJsonConfigFound(f"The configuration file for the {report_name} graphics does not exist")

        config_path = os.path.join(self.JSON_FOLDER_URL, report_name)

        try:

            with open(config_path, 'r', encoding='utf-8') as file:

                self.graphics_config = json.load(file)

            logger.info(f"Grapchis config loaded succesfully {report_name}")

        except json.JSONDecodeError as e:

            logger.error(f"Invalid json in {report_name}: {e}")
            raise

        except Exception as e:

            logger.error(f"Error reading {report_name}  graphics config: {e}")
            raise

    def main(self, report_name: str, chain_of_custody):

        try:

            self.get_chain_of_custody(chain_of_custody)
            self.reader_graphics_data.get_chain_custody(self.chain_of_custody)
            self.search_graphics_config(report_name)
            self.reader_graphics_data.get_cod_muestras(self.graphics_config)
            self.writer_graphics_data.load_excel_to_write(self.graphics_config["RUTA_EXCEL_GRAFICAS"])
            tables_data = self.reader_graphics_data.read_table_ph_solidos_caudales()
            formatted_tables_data = self.reader_graphics_data.complete_table_data()
            self.writer_graphics_data.load_insitu_data(formatted_tables_data)
            self.writer_graphics_data.load_json_config(self.graphics_config)
            self.writer_graphics_data.write_insitu_data()

        except Exception as ex:

            logger.error(f"Error in the main thread {ex}")
            logger.error(f"Traceback: {traceback.format_exc()}")








        
        


