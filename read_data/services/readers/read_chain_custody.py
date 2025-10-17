from core.utils.data.hour_to_str import hour_to_str
from read_data.services.readers.read_punctual_sheet_data import read_punctual_sheet_data
from read_data.services.readers.read_specific_sheet import read_specific_sheet, read_specific_sheet_data


def read_chain_custody(workbook, sheet_name: str, file_path: str) -> dict:
    try:


        print(f"ABRIENDO HOJA -> {sheet_name}")

        chain_custody_sheet = workbook[sheet_name]
        samples_data = {}

        sample_description_afluente = read_specific_sheet(file_path, "AFLUENTE 1")
        sample_description_efluente = read_specific_sheet(file_path, "EFLUENTE 2")
        sample_description_3 = read_specific_sheet(file_path, "CADENA DE VIGILANCIA PUNTUAL")

        # Leer datos ambientales de cada hoja
        clima_afluente, temp_afluente, humedad_afluente, altitud_afluente = read_specific_sheet_data(workbook,
                                                                                                     "AFLUENTE 1")
        clima_efluente, temp_efluente, humedad_efluente, altitud_efluente = read_specific_sheet_data(workbook,
                                                                                                     "EFLUENTE 2")


        clima_descarga, temp_descarga, humedad_descarga, altitud_descarga = read_specific_sheet_data(workbook,
                                                                                                     "CADENA DE VIGILANCIA PUNTUAL", day_type_start_row=18, nublado_col='F', soleado_col='I', lluvioso_col='L', temp_col = "O", hum_col="S", alt_col="W")

        descriptions = [sample_description_afluente, sample_description_efluente, sample_description_3]

        # ✅ Lista de datos ambientales correspondientes a cada muestra
        environmental_data = [
            {
                "clima": clima_afluente,
                "temperatura_ambiente": temp_afluente,
                "humedad_relativa": humedad_afluente,
                "altitud": altitud_afluente
            },
            {
                "clima": clima_efluente,
                "temperatura_ambiente": temp_efluente,
                "humedad_relativa": humedad_efluente,
                "altitud": altitud_efluente
            },
            {
                "clima": clima_descarga,
                "temperatura_ambiente": temp_descarga,
                "humedad_relativa": humedad_descarga,
                "altitud": altitud_descarga
            }
        ]

        hours = read_punctual_sheet_data(workbook, "CADENA DE VIGILANCIA PUNTUAL")

        for idx, row in enumerate(range(23, 28, 2)):
            if chain_custody_sheet[f"A{row}"].value != '' and chain_custody_sheet[f"A{row}"].value != None:

                sample = {}
                sample["chemilab_code"] = chain_custody_sheet[f"A{row}"].value
                sample["sample_identification"] = chain_custody_sheet[f"C{row}"].value
                sample["sample_year"] = chain_custody_sheet[f"G{row}"].value
                sample["sample_month"] = chain_custody_sheet[f"H{row}"].value
                sample["sample_day"] = chain_custody_sheet[f"I{row}"].value
                sample["sample_description"] = descriptions[idx] if idx < len(descriptions) else None



                hour_obj = hours.get(str(idx + 1), None)
                print(hour_obj)
                sample["sampler_hour"] = hour_to_str(hour_obj)

                # ✅ AGREGAR DATOS AMBIENTALES A CADA MUESTRA
                if idx < len(environmental_data):
                    sample["sample_weather"] = environmental_data[idx]["clima"]
                    sample["sample_temperature"] = environmental_data[idx]["temperatura_ambiente"]
                    sample["sample_humidity"] = environmental_data[idx]["humedad_relativa"]
                    sample["sample_altitude"] = environmental_data[idx]["altitud"]

                samples_data[chain_custody_sheet[f"A{row}"].value] = sample

        samples_data["OSI"] = chain_custody_sheet["H49"].value

        print(samples_data)
        return samples_data

    except KeyError:
        print(f"Sheet {sheet_name} not found please review")
        return {}

    except Exception as e:
        print(f"Error opening the sheet: {e}")
        return {}