from core.utils.data.hour_to_str import hour_to_str
from read_data.services.readers.read_punctual_sheet_data import read_punctual_sheet_data
from read_data.services.readers.read_specific_sheet import read_specific_sheet, read_specific_sheet_data


def read_chain_custody(workbook, sheet_name: str, file_path: str, config: dict, hojas_muestras: dict) -> dict:
    try:
        chain_custody_sheet = workbook[sheet_name]
        samples_data = {}

        descriptions = []
        environmental_data = []

        for index, hoja_info in hojas_muestras.items():
            nombre_hoja = hoja_info["NOMBRE"]
            params = hoja_info.get("PARAMETROS_AMBIENTALES", {})

            if not params:
                continue

            desc = read_specific_sheet(file_path, nombre_hoja)
            descriptions.append(desc)

            clima, temp, humedad, altitud = read_specific_sheet_data(
                workbook,
                nombre_hoja,
                day_type_start_row=params["FILA_INICIO_CLIMA"],
                nublado_col=params["COL_NUBLADO"],
                soleado_col=params["COL_SOLEADO"],
                lluvioso_col=params["COL_LLUVIOSO"],
                temp_col=params["COL_TEMPERATURA"],
                hum_col=params["COL_HUMEDAD"],
                alt_col=params["COL_ALTITUD"]
            )

            environmental_data.append({
                "clima": clima,
                "temperatura_ambiente": temp,
                "humedad_relativa": humedad,
                "altitud": altitud
            })

        hours = read_punctual_sheet_data(workbook, config["HOJA_HORAS"])

        cols = config["COLUMNAS"]
        fila_inicial = config["FILA_INICIAL"]
        fila_final = config["FILA_FINAL"]
        incremento = config["INCREMENTO_FILA"]

        for idx, row in enumerate(range(fila_inicial, fila_final, incremento)):
            codigo_col = cols["CODIGO_CHEMILAB"]

            if chain_custody_sheet[f"{codigo_col}{row}"].value not in [None, '']:
                sample = {}
                sample["chemilab_code"] = chain_custody_sheet[f"{codigo_col}{row}"].value
                sample["sample_identification"] = chain_custody_sheet[f"{cols['IDENTIFICACION_MUESTRA']}{row}"].value
                sample["sample_year"] = chain_custody_sheet[f"{cols['AÑO']}{row}"].value
                sample["sample_month"] = chain_custody_sheet[f"{cols['MES']}{row}"].value
                sample["sample_day"] = chain_custody_sheet[f"{cols['DIA']}{row}"].value
                sample["sample_description"] = descriptions[idx] if idx < len(descriptions) else None

                hour_obj = hours.get(str(idx + 1), None)
                sample["sampler_hour"] = hour_to_str(hour_obj)

                if idx < len(environmental_data):
                    sample["sample_weather"] = environmental_data[idx]["clima"]
                    sample["sample_temperature"] = environmental_data[idx]["temperatura_ambiente"]
                    sample["sample_humidity"] = environmental_data[idx]["humedad_relativa"]
                    sample["sample_altitude"] = environmental_data[idx]["altitud"]

                samples_data[sample["chemilab_code"]] = sample

        samples_data["OSI"] = chain_custody_sheet[config["CELDA_OSI"]].value

        return samples_data

    except KeyError as e:
        print(f"Sheet {sheet_name} not found: {e}")
        return {}
    except Exception as e:
        print(f"Error opening the sheet: {e}")
        import traceback
        traceback.print_exc()
        return {}