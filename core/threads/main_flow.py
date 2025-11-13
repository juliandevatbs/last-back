import json
import logging
import os

from core.tasks import Tasks
from graphics.Services.graphic_generator import GraphicGenerator

logger = logging.getLogger(__name__)


class MainFlow:

    def __init__(self):

        # Instances
        self.general_tasks = Tasks()
        self.graphics_core = GraphicGenerator()

        # Template config
        self.template_config = None

    def load_json_config(self, template_name: str):

        try:

            config = self.general_tasks.search_template_config(template_name)

            logger.info(f"Configuration loaded succesfully for {template_name}")

            return config

        except Exception as e:

            logger.error(f"Error loading json template config {e}")

            raise

    def clean_json_config(self, template_name: str):

            try:

                if not template_name.endswith('.json'):

                    template_name += '.json'

                config_path = os.path.join(self.general_tasks.JSON_FOLDER_URL, template_name)

                with open(config_path, 'w', encoding='utf-8') as file:

                    json.dump(self.template_config, file, indent=4)

                    logger.info("Json cleaned succesfully")

            except Exception as e:

                logger.error("Error cleaning the json")

    def main_flow_caller(self, chain_of_custody, template_name: str):

        # Orchestrates the complete flow
        try:

            self.template_config = self.load_json_config(template_name)

            self.clean_json_config(template_name)

            #Graphics
            self.graphics_core.main(template_name, chain_of_custody)

        except Exception as ex:

            logger.error(f"Error opening the template config {ex}")

            raise
