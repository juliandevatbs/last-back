import json

class Writer():



    def __init__(self):

        with open('fields_config/fields.json', 'r') as json_file:

            self.data = json.load(json_file)






    def get_json_config(self, template_name: str) -> dict:

        print(self.data)












