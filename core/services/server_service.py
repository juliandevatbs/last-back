from core.exceptions import ServerClientException
from core.services.server_client import ServerClient


class ServerService:


    def __init__(self):

        self.client = ServerClient()


    def get_avaible_templates(self):

        try:

            templates = self.client.get_template_folders()

            print(templates)

            return templates


        except ServerClientException as e:

            print(f"Get templates failed {e}")


