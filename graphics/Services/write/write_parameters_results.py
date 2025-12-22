import logging
import math

logger = logging.getLogger(__name__)


class WriteParametersResults:

    def __init__(self):
        return

    def write_parameters_result(self, graphics_excel, json_config, results: dict):
        if not graphics_excel:
            logger.error(f"Not graphics excel loaded")
            return

        if not json_config:
            logger.error(f"Not json config loaded")
            return

        sheet_to_write = graphics_excel["RESULTADOS"]
        samples_config = json_config["PARAMETROS"]["COD_MUESTRAS"]

        for sample_code, parameters_data in results.items():
            sample_keys = list(samples_config.keys())
            result_keys = list(results.keys())

            sample_index = result_keys.index(sample_code)
            sample_key = sample_keys[sample_index]
            sample_config = samples_config[sample_key]

            date_col = sample_config["FECHA_COORD"][0]
            date_row_initial = int(sample_config["FECHA_COORD"][1])

            result_col = sample_config["INCERTIDUMBRE_COORD"][0]
            result_row_initial = int(sample_config["INCERTIDUMBRE_COORD"][1])

            skip_rows = sample_config.get("SKIP_FILA", [])

            parameters_column = json_config["PARAMETROS"]["COORD_PARAMETROS"][0]
            parameters_row = int(json_config["PARAMETROS"]["COORD_PARAMETROS"][1])
            parameters_last_row = json_config["PARAMETROS"]["CANT_PARAMETROS"] + parameters_row

            results_sheet = graphics_excel[json_config["PARAMETROS"]["HOJA"]]

            current_excel_row_date = date_row_initial
            current_excel_row_result = result_row_initial

            for i in range(parameters_row, parameters_last_row):
                if current_excel_row_result in skip_rows:
                    current_excel_row_date += 1
                    current_excel_row_result += 1
                    continue

                param_name = results_sheet[f"{parameters_column}{i}"].value

                if not param_name:
                    current_excel_row_date += 1
                    current_excel_row_result += 1
                    continue

                print(f"Row {i}: Searching parameter '{param_name}' to write at row {current_excel_row_result}")

                found_data = None

                for key, value in parameters_data.items():
                    if value.get("ES_MULTIVALUE") and value.get("PARAMETRO_ORIGINAL") == param_name:
                        if not value.get("USADO", False):
                            found_data = value
                            value["USADO"] = True
                            break

                if not found_data and param_name in parameters_data:
                    found_data = parameters_data[param_name]

                if found_data:
                    date = found_data["FECHA"]
                    result = found_data["RESULTADO"]

                    sheet_to_write[f"{date_col}{current_excel_row_date}"].value = date
                    sheet_to_write[f"{result_col}{current_excel_row_result}"].value = result

                    print(f"  ✓ Written: {result} at row {current_excel_row_result}")
                else:
                    logger.warning(f"Data not found for parameter '{param_name}' at row {i}")

                current_excel_row_date += 1
                current_excel_row_result += 1

    def write_final_results(self, final_results: dict, graphics_obj, json_config):
        try:
            sheet_to_write = graphics_obj["GRÁFICAS"]
            samples_code_config = json_config["PARAMETROS"]["COD_MUESTRAS"]

            for sample_code, data in final_results.items():
                if sample_code not in samples_code_config:
                    continue

                sample_config = samples_code_config[sample_code]
                result_coord = sample_config["FINAL_COORD"]

                col_coord = ''.join([c for c in result_coord if c.isalpha()])
                row_init = int(''.join([c for c in result_coord if c.isdigit()]))

                print(f"\n=== Procesando {sample_code} ===")
                print(f"Coordenada inicial: {result_coord} (fila {row_init})")
                print(f"Total datos a escribir: {len(data)}")

                for data_index, value in enumerate(data):
                    actual_row = row_init + data_index
                    cell = f"{col_coord}{actual_row}"

                    if value is None or (isinstance(value, str) and value.strip() == ''):
                        sheet_to_write[cell].value = None
                        print(f"  [EMPTY] {cell} = vacío (data[{data_index}])")
                        continue

                    numeric_value, has_symbol = self._extract_numeric_value_simple(value)

                    if numeric_value is not None:
                        sheet_to_write[cell].value = numeric_value

                        format_str = self._get_format_3_significant_figures(numeric_value, has_symbol)
                        sheet_to_write[cell].number_format = format_str

                        print(f"  [OK] {cell} = {has_symbol or ''}{numeric_value} (formato: {format_str})")
                    else:
                        sheet_to_write[cell].value = value
                        print(f"  [TEXT] {cell} = '{value}'")

                print(f"=== Fin {sample_code}: {len(data)} valores escritos ===\n")

        except Exception as ex:
            logger.error(f"Error writing the final data {ex}")
            raise

    def _extract_numeric_value_simple(self, value):
        if value is None:
            return None, None

        if isinstance(value, (int, float)):
            return float(value), None

        value_str = str(value).strip()
        symbol = None

        if value_str.startswith(('<', '>')):
            symbol = value_str[0]
            value_str = value_str[1:].strip()

        value_str = value_str.replace(' ', '')

        try:
            if ',' in value_str and '.' not in value_str:
                if value_str.count(',') == 1:
                    value_str = value_str.replace(',', '.')

            numeric_value = float(value_str)
            return numeric_value, symbol
        except (ValueError, AttributeError):
            logger.warning(f"Could not convert '{value}' to number")
            return None, None

    def _get_format_3_significant_figures(self, value, symbol=None):
        if value == 0:
            decimals = 2
        else:
            abs_value = abs(value)
            if abs_value >= 100:
                decimals = 0
            elif abs_value >= 10:
                decimals = 1
            elif abs_value >= 1:
                decimals = 2
            elif abs_value >= 0.1:
                decimals = 3
            elif abs_value >= 0.01:
                decimals = 4
            elif abs_value >= 0.001:
                decimals = 5
            else:
                decimals = 6

        format_pattern = "0"
        if decimals > 0:
            format_pattern += "." + "0" * decimals

        if symbol == '<':
            return '"<"' + format_pattern
        elif symbol == '>':
            return '">"' + format_pattern
        else:
            return format_pattern