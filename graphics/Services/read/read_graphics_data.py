

import io
import logging

from openpyxl.reader.excel import load_workbook

from core.database.SamplerDatabase import SamplerDatabase

logger = logging.getLogger(__name__)

class ReadGraphicsData:

    def __init__(self):

        self.chain_custody_obj = None
        self.results_sheet_obj = None
        self.info_cod_muestras = None
        self.json_config = None

        self.query = """

           SELECT
	
	MUESTRAS.FECHAHORAMUESTRAREAL,

     CASE
    WHEN MEDICIONESMUESTRA.VALORMEDIDOCADENA IS NOT NULL 
         AND LTRIM(RTRIM(MEDICIONESMUESTRA.VALORMEDIDOCADENA)) <> '' 
        THEN MEDICIONESMUESTRA.VALORMEDIDOCADENA

    WHEN PARAMETROS.INCERTIDUMBRE IS NULL OR PARAMETROS.INCERTIDUMBRE <= 0 
        THEN CAST(MEDICIONESMUESTRA.VALORMEDIDONUMERICO AS VARCHAR(50))
    
    WHEN PARAMETROS.IDCLASEPARAMETRO = 2 THEN
        CONCAT(
            MEDICIONESMUESTRA.VALORMEDIDONUMERICO, 
            '+/- (', 
            CAST(ROUND(MEDICIONESMUESTRA.VALORMEDIDONUMERICO / EXP(PARAMETROS.INCERTIDUMBRE / 100.0), 0) AS INT), 
            ' - ', 
            CAST(ROUND(MEDICIONESMUESTRA.VALORMEDIDONUMERICO * EXP(PARAMETROS.INCERTIDUMBRE / 100.0), 0) AS INT), 
            ')'
        )

    
    -- Para parámetros normales (clase 1), escalar decimales según magnitud
    ELSE
        CASE
            WHEN PARAMETROS.INCERTIDUMBRE * MEDICIONESMUESTRA.VALORMEDIDONUMERICO >= 100 THEN
                CONCAT(MEDICIONESMUESTRA.VALORMEDIDONUMERICO, '+/- ', REPLACE(CAST(ROUND(PARAMETROS.INCERTIDUMBRE * MEDICIONESMUESTRA.VALORMEDIDONUMERICO, 0) AS VARCHAR(50)), '.', ','))

            WHEN PARAMETROS.INCERTIDUMBRE * MEDICIONESMUESTRA.VALORMEDIDONUMERICO >= 10 THEN
                CONCAT(MEDICIONESMUESTRA.VALORMEDIDONUMERICO, '+/- ', REPLACE(CAST(ROUND(PARAMETROS.INCERTIDUMBRE * MEDICIONESMUESTRA.VALORMEDIDONUMERICO, 1) AS VARCHAR(50)), '.', ','))

            WHEN PARAMETROS.INCERTIDUMBRE * MEDICIONESMUESTRA.VALORMEDIDONUMERICO >= 1 THEN
                CONCAT(MEDICIONESMUESTRA.VALORMEDIDONUMERICO, '+/- ', REPLACE(CAST(ROUND(PARAMETROS.INCERTIDUMBRE * MEDICIONESMUESTRA.VALORMEDIDONUMERICO, 2) AS VARCHAR(50)), '.', ','))

            WHEN PARAMETROS.INCERTIDUMBRE * MEDICIONESMUESTRA.VALORMEDIDONUMERICO >= 0.1 THEN
                CONCAT(MEDICIONESMUESTRA.VALORMEDIDONUMERICO, '+/- ', 
                    REPLACE(
                        REPLACE(FORMAT(ROUND(PARAMETROS.INCERTIDUMBRE * MEDICIONESMUESTRA.VALORMEDIDONUMERICO, 3), 'N3'), ',', ''),
                        '.', 
                        ','
                    )
                )

            WHEN PARAMETROS.INCERTIDUMBRE * MEDICIONESMUESTRA.VALORMEDIDONUMERICO >= 0.01 THEN
                CONCAT(MEDICIONESMUESTRA.VALORMEDIDONUMERICO, '+/- ', 
                    REPLACE(
                        REPLACE(FORMAT(ROUND(PARAMETROS.INCERTIDUMBRE * MEDICIONESMUESTRA.VALORMEDIDONUMERICO, 4), 'N4'), ',', ''),
                        '.', 
                        ','
                    )
                )

            ELSE
                CONCAT(MEDICIONESMUESTRA.VALORMEDIDONUMERICO, '+/- ', 
                    REPLACE(
                        REPLACE(FORMAT(ROUND(PARAMETROS.INCERTIDUMBRE * MEDICIONESMUESTRA.VALORMEDIDONUMERICO, 5), 'N5'), ',', ''),
                        '.', 
                        ','
                    )
                )
        END
	END AS INCERTIDUMBRE

FROM MUESTRAS

	INNER JOIN MEDICIONESMUESTRA ON MUESTRAS.IDMUESTRA = MEDICIONESMUESTRA.IDMUESTRA
	INNER JOIN SUBMATRIZ ON MUESTRAS.IDSUBMATRIZ = SUBMATRIZ.IDSUBMATRIZ
	INNER JOIN PARAMETROS ON MEDICIONESMUESTRA.IDPARAMETRO = PARAMETROS.IDPARAMETRO
	INNER JOIN METODOS ON PARAMETROS.IDMETODO = METODOS.IDMETODO
	LEFT JOIN UNIDADMEDIDA ON MEDICIONESMUESTRA.IDUNIDADMEDIDA = UNIDADMEDIDA.IDUNIDADMEDIDA
    
    	
WHERE MUESTRAS.CODMUESTRA = ?
AND NOMBREREPORTE  = ? 
	ORDER BY NOMBREREPORTE;
        
        

        """

        # Tables data(ph, solidos, caudales)
        self.ph_solidos_caudales_tbl = None

        # Sampler parameters (to read to the database info non insitu graphics)
        self.sampler_parameters = None
        self.specific_parameters = None

    def get_chain_custody(self, chain_custody_obj):

            try:

                self.chain_custody_obj = load_workbook(io.BytesIO(chain_custody_obj), data_only=True)

            except Exception as ex:

                logger.error(f"Invalid chain of custody excel -> {ex}")

    def load_json_config(self, json_config_obj):

            if not json_config_obj:
                logger.error("Json config no loaded")
                return

            self.json_config = json_config_obj

    def get_cod_muestras(self):

            if not self.chain_custody_obj:
                logger.error("Chain of custody not loaded")

                return

            if not self.json_config:
                logger.error("Json config not loaded")

                return

            info_cod_muestras = {}

            for key, value in self.json_config["INSITU"]["INFO_TABLAS_LECTURA"].items():
                info_cod_muestras[key] = value

            self.info_cod_muestras = info_cod_muestras

            for key, value in info_cod_muestras.items():
                # Sheet chain of custody
                sheet_ch = self.chain_custody_obj[value["HOJA_COD_MUESTRA"]]
                cod_muestra_value = sheet_ch[value["COD_MUESTRA_UBI"]].value

                self.info_cod_muestras[key]["COD_MUESTRA"] = cod_muestra_value

    def complete_table_data(self):

            if self.ph_solidos_caudales_tbl:

                for key, registros in self.ph_solidos_caudales_tbl.items():

                    for register in registros:

                        ph_value = float(str(register.get('PH'))) * 0.01

                        formatted_ph = f"±{ph_value:.4f}".replace(".", ",")

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

                    # init_column = value["COLUMNA_INICIAL_DATOS"]

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

    def read_parameters_to_sampler(self, excel_graphics_obj):

        if not self.chain_custody_obj:
            logger.error("Not chain of custody loaded")
            return

        sampler_parameters = {}

        try:
            ch_sheet = self.chain_custody_obj["CADENA DE CUSTODIA"]

            sampler_parameters[ch_sheet["A23"].value] = {}
            sampler_parameters[ch_sheet["A25"].value] = {}
            sampler_parameters[ch_sheet["A27"].value] = {}

            sampler_parameters[ch_sheet["A23"].value]["OSI"] = ch_sheet["H49"].value
            sampler_parameters[ch_sheet["A23"].value]["PM"] = ch_sheet["AK10"].value

            sampler_parameters[ch_sheet["A25"].value]["OSI"] = ch_sheet["H49"].value
            sampler_parameters[ch_sheet["A25"].value]["PM"] = ch_sheet["AK10"].value

            sampler_parameters[ch_sheet["A27"].value]["OSI"] = ch_sheet["H49"].value
            sampler_parameters[ch_sheet["A27"].value]["PM"] = ch_sheet["AK10"].value

        except Exception as ex:
            logger.error(f"Error reading the sampler parameters {ex}")

        self.sampler_parameters = sampler_parameters

        try:
            self.results_sheet_obj = excel_graphics_obj[self.json_config["PARAMETROS"]["HOJA"]]

            parameters_column = self.json_config["PARAMETROS"]["COORD_PARAMETROS"][0]
            parameters_row = int(self.json_config["PARAMETROS"]["COORD_PARAMETROS"][1])
            parameters_last_row = self.json_config["PARAMETROS"]["CANT_PARAMETROS"] + parameters_row

            parameters_results = {}
            instance_db = SamplerDatabase()

            for key, muestra_config in self.json_config["PARAMETROS"]["COD_MUESTRAS"].items():

                ch_sheet = self.chain_custody_obj[muestra_config["HOJA_COD_MUESTRA"]]
                cod_muestra = ch_sheet[muestra_config["COD_MUESTRA_UBI"]].value

                if not cod_muestra:
                    logger.warning(f"No se encontró código de muestra en {muestra_config['COD_MUESTRA_UBI']}")
                    continue

                parameters_results[cod_muestra] = {}

                parametros_consultados = {}

                for i in range(parameters_row, parameters_last_row):
                    param = self.results_sheet_obj[f"{parameters_column}{i}"].value

                    if not param:
                        continue

                    if param in parametros_consultados:
                        logger.info(f"Parámetro '{param}' ya consultado, reutilizando resultado")
                        continue

                    try:
                        results = instance_db.execute_query(self.query, (str(cod_muestra), param))

                        if results and len(results) > 0:
                            resultado_bd = results[0][1]
                            fecha_bd = results[0][0]

                            if resultado_bd and '-' in str(resultado_bd) and not str(resultado_bd).startswith('+/-'):
                                valores_individuales = str(resultado_bd).split('-')

                                logger.info(
                                    f"Parámetro '{param}' tiene {len(valores_individuales)} valores: {valores_individuales}")

                                for idx, valor in enumerate(valores_individuales):
                                    param_key = f"{param}_VALOR_{idx + 1}"
                                    parameters_results[cod_muestra][param_key] = {
                                        "FECHA": fecha_bd,
                                        "RESULTADO": valor.strip(),
                                        "PARAMETRO_ORIGINAL": param,
                                        "ES_MULTIVALUE": True,
                                        "INDICE": idx
                                    }
                            else:
                                parameters_results[cod_muestra][param] = {
                                    "FECHA": fecha_bd,
                                    "RESULTADO": resultado_bd,
                                    "ES_MULTIVALUE": False
                                }

                            parametros_consultados[param] = True

                    except Exception as ex:
                        logger.error(f"Error querying parameters results for {cod_muestra} - {param}: {ex}")

            return sampler_parameters, parameters_results

        except Exception as ex:
            logger.error(f"Error reading specific parameters: {ex}")
            return None, None

    def graphics_post_calculus(self, excel_graphics_obj, json_config):


        # This read the results calculated by the excel to generate the final graphics
        if not json_config:

            logger.error("No json config loaded")

            return

        if not excel_graphics_obj:

            logger.error("No excel to read loaded")

            return

        # Sheet to read
        try:
            graphics_sheet = excel_graphics_obj[json_config["PARAMETROS"]["HOJA"]]
            cant_params = json_config["PARAMETROS"]["CANT_PARAMETROS"]
            # Quantity of params + excel row = limit

            final_results = {}


            # get the columns
            for key, values in json_config["PARAMETROS"]["COD_MUESTRAS"].items():

                final_results[key] = []

                coord_to_read = values["RESULTADO_FINAL_COORD"]
                limit_row = cant_params + int(coord_to_read[1])

                init_row = coord_to_read[1]


                for i in range(int(init_row), limit_row):

                    final_results[key].append(graphics_sheet[f"{coord_to_read[0]}{i}"].value)

            return final_results




        except KeyError as ex:

            logger.error(f"No sheet to read found {ex}")

            return None

        except Exception as ex:

            logger.error(f"Error reading the post graphics data {ex}")

            return None





















