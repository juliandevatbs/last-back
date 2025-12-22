import json
from read_data.services.readers.ph_reader import ph_reader
from read_data.services.readers.read_chain_custody import read_chain_custody
from read_data.services.readers.read_main_sheet import read_main_sheet


class ExcelReaderMain:
    def __init__(self):
        self.workbook = None
        self.config_data = None
        self.basic_data = None
        self.samples_data = None
        self.specific_data = {}

    def load_work_book(self, workbook):
        self.workbook = workbook

    def load_json_data(self, inf_name: str):

        try:
            with open(f'fields_config/{inf_name}.json', 'r', encoding='utf-8') as config_inf:
                self.config_data = json.load(config_inf)
        except Exception as ex:
            raise Exception(f"Error cargando JSON: {ex}")

    def caller(self):
        try:
            basic_sheet = self.config_data['HOJAS_IMPORTANTES']['HOJA_BASICOS']
            custody_sheet = self.config_data['HOJAS_IMPORTANTES']['CADENA_CUSTODIA']

            self.basic_data = read_main_sheet(self.workbook, basic_sheet)

            self.basic_data["XX_REVISADO_POR_XX"] ="Claudia Calderón"
            self.basic_data["XX_ROL_REVISADOR_XX"] = "Profesional de proyectos"
            self.basic_data["XX_AUTORIZADO_POR_XX"] = "Claudia Calderón"
            self.basic_data["XX_AUTORIZADO_POR_ROL_XX"] = "Directora de poyectos"

            self.samples_data = read_chain_custody(
                self.workbook,
                custody_sheet,
                self.config_data['RUTA_EXCEL_CADENA'],
                self.config_data.get('CADENA_CUSTODIA_CONFIG', {}),
                self.config_data['HOJAS_MUESTRAS']
            )

            for index, sheet_info in self.config_data['HOJAS_MUESTRAS'].items():
                sheet_name = sheet_info["NOMBRE"]
                columns = sheet_info.get("COLUMNAS", {})
                initial_row = sheet_info.get("FILA_INICIAL")

                if columns and initial_row is not None:
                    ph_data = ph_reader(self.workbook, sheet_name, columns, initial_row)
                    #self.specific_data[sheet_name] = ph_data
                    print(f"PH DATA {ph_data}")

                    for code, sample in self.samples_data.items():

                        if code == 'OSI':
                            continue

                        sample_identification = sample.get("sample_identification", "").lower()
                        sheet_name_lower = sheet_name.lower()

                        if (sheet_name_lower in sample_identification or
                            any(word in sample_identification for word in sheet_name_lower.split())
                        ):


                            sample["mediciones"] = ph_data
                            #break

            return self.basic_data, self.samples_data, self.specific_data

        except Exception as ex:
            raise Exception(f"Error leyendo hojas: {ex}")