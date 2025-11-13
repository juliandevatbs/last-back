import io
import logging

from openpyxl.reader.excel import load_workbook

logger = logging.getLogger(__name__)



class ReadGraphicsData:



    def __init__(self):

        self.chain_custody_obj = None
        self.info_cod_muestras = None

        #Tables data(ph, solidos, caudales)
        self.ph_solidos_caudales_tbl = None

    def get_chain_custody(self, chain_custody_obj):

        try:

            self.chain_custody_obj = load_workbook(io.BytesIO(chain_custody_obj), data_only=True)

        except Exception as ex:

            logger.error(f"Invalid chain of custody excel -> {ex}")

    def get_cod_muestras(self, json_config):

        if not self.chain_custody_obj:

            logger.error("Chain of custody not loaded")

            return


        if not json_config:

            logger.error("Json config not loaded")

            return


        info_cod_muestras = {}


        for key, value in json_config["INSITU"]["INFO_TABLAS_LECTURA"].items():

            info_cod_muestras[key] = value


        self.info_cod_muestras = info_cod_muestras

        for key, value in info_cod_muestras.items():

            #Sheet chain of custody
            sheet_ch = self.chain_custody_obj[value["HOJA_COD_MUESTRA"]]
            cod_muestra_value = sheet_ch[value["COD_MUESTRA_UBI"]].value

            self.info_cod_muestras[key]["COD_MUESTRA"] = cod_muestra_value

    def complete_table_data(self):

        if self.ph_solidos_caudales_tbl:

            for key, registros in self.ph_solidos_caudales_tbl.items():

                for register in registros:

                    print(register)

                    formatted_ph = f"±{float(str(register.get('PH'))):.4f}".replace(".", ",")

                    register["INCERTIDUMBRE_PH"] = formatted_ph

                    if register.get("SOLIDOS SEDIMENTABLES") != None:

                        formatted_value = register.get('SOLIDOS SEDIMENTABLES')[1:].strip().replace(",", ".")

                        register["SOLIDOS SEDIMENTABLES"] = f"<{float(formatted_value):.3f}".replace(".", ",")

            return self.ph_solidos_caudales_tbl



    def read_table_ph_solidos_caudales(self):

        self.ph_solidos_caudales_tbl = {}

        try:

            for key_parent, value_parent in self.info_cod_muestras.items():

                sheet_ch = self.chain_custody_obj[value_parent["HOJA_EXCEL"]]

                # Coords where inits the data
                init_row = value_parent["FILA_INICIAL_DATOS"]

                #init_column = value["COLUMNA_INICIAL_DATOS"]


                # Registers per sheet
                rows_quantity = value_parent["CANTIDAD_REGISTROS"]
                columns_data = value_parent["COLUMNAS_REGISTROS"]

                self.ph_solidos_caudales_tbl[key_parent] = []

                if key_parent not in self.ph_solidos_caudales_tbl:
                    self.ph_solidos_caudales_tbl[key_parent] = {}

                for i in range(rows_quantity):

                    row = {}

                    for key, value in columns_data.items():

                        cell_cords = f"{value}{init_row + i}"

                        row[key] = sheet_ch[cell_cords].value

                    self.ph_solidos_caudales_tbl[key_parent].append(row)

        except Exception as ex:

            logger.error(f"Error reading the ph, caudales, solidos data {ex}")


        """for key, value in self.ph_solidos_caudales_tbl.items():

            print(f"{key}: {value}")"""

        return self.ph_solidos_caudales_tbl
























