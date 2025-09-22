import logging
import os

from django.contrib.sites import requests
from dotenv import load_dotenv

from core.exceptions import ServerClientException
import requests as http_requests
from django.conf import settings
from docx import Document

load_dotenv()

logger = logging.getLogger(__name__)

class ServerClient:


    def __init__(self):

        self.SERVER_HOST= os.getenv("SERVER_HOST")
        self.F2025_PROJECTS = os.getenv("SERVER_URL_PROJECTS_25")
        self.TEMPLATES_FOLDER_URL = os.getenv('SERVER_URL_FOLDER_TEMPLATES')
        self.LOCAL_TEMPLATES_PATH = "templates/"


    def get_template_folders(self) -> list[dict]:

        # Get the list of folders into the 2025 projects folder

        try:

            response = http_requests.get(
                # Build the complete url
                f"{self.SERVER_HOST}{self.F2025_PROJECTS}",
            )

            response.raise_for_status()
            return response.json()

        except http_requests.RequestException as e:

            logger.error(f"Error accessing template server: {e}")
            raise ServerClientException(f"Failed to fetch template names: {e}")

    def get_selected_template(self, selected_name_template: str):

        if not selected_name_template:
            logger.error("No template option was provided")
            raise ValueError("A template must be selected")

        # Ensure the file has .docx extension

        if not selected_name_template.endswith('.docx'):

            selected_name_template +='.docx'
            logger.debug(f"Added .docx extension: {selected_name_template}")

        # Build the full local path
        template_path= os.path.join("templates", selected_name_template)






        try:

            if not os.path.exists(template_path):

                logger.error(f"Template not found in the url -> {template_path}")


                raise ServerClientException(

                    f"Template not found: {selected_name_template}"



                )
            logger.info(f"Loading template from {template_path}")

            template_doc = Document(template_path)

            return template_doc

        except ServerClientException:

            #Re throw exceptions that have already been catch
            raise

        except Exception as e:

            logger.error(f"Error loading local template {selected_name_template}: {e}")
            raise ServerClientException(f"The template could not be loaded {e}")













