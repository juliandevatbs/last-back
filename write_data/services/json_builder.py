import io
import json
from typing import Optional, Dict, Any

from openpyxl.reader.excel import load_workbook



class JsonBuilder():
    """Manages excel data extraction and JSON configuration updates"""

    def __init__(self):
        """
        Initialize JsonBuilder with excel file bytes

        args:
            file_bytes = Excel file content as bytes
        """
        self.json_object = None
        self.json_name = None



    def load_json(self, json_name: str):

        try:
            self.json_name = json_name
            with open(f'fields_config/{json_name}.json', 'r', encoding='utf-8') as config_inf:
                self.json_object = json.load(config_inf)

        except Exception as ex:
            raise Exception(f"Error cargando JSON: {ex}")

    def update_json_normal_values(self, basic_data):
        if self.json_object:
            encabezado_values = self.json_object["ENCABEZADO_CONFIG"]["VARIABLES"]
            portada_values = self.json_object["PORTADA_CONFIG"]["VARIABLES"]
            generales_values = self.json_object["HOJAS_GENERALES_CONFIG"]["VARIABLES"]

            updated_keys = []
            for key, item in basic_data.items():
                if key in encabezado_values:
                    encabezado_values[key] = item
                    updated_keys.append(f"ENCABEZADO: {key}")
                elif key in portada_values:
                    portada_values[key] = item
                    updated_keys.append(f"PORTADA: {key}")
                elif key in generales_values:
                    generales_values[key] = item
                    updated_keys.append(f"GENERALES: {key}")

            print("Variables actualizadas:", updated_keys)  # Debug

    def update_json_samples(self, samples_data):

        if self.json_object and "HOJAS_MUESTRAS" in self.json_object:

            if "DATOS_MUESTRAS" not in self.json_object:

                self.json_object["DATOS_MUESTRAS"] = {}

            self.json_object["DATOS_MUESTRAS"] = samples_data

    def save_json(self, json_name: str):

        save_name = json_name if json_name else self.json_name

        if self.json_object and self.json_name:
            try:
                with open(f'fields_config/{self.json_name}.json', 'w', encoding='utf-8') as config_inf:
                    json.dump(self.json_object, config_inf, indent=2, ensure_ascii=False)
            except Exception as ex:
                raise Exception(f"Error guardando JSON: {ex}")
        else:
            raise Exception("No hay JSON cargado para guardar")












