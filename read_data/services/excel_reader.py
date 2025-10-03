from openpyxl.reader.excel import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from core.utils.data.date_literal import date_literal
from core.utils.data.get_today_date import get_today_date


def read_main_sheet_excel(workbook) -> dict:
    main_sheet = workbook.worksheets[1]

    # print(main_sheet)

    # Dict to storage the data
    data_client = {}
    sampling_data = {}

    # Storage client and main report data
    data_client["XX_FECHA_ELABORACION_XX"] = get_today_date()
    data_client["XX_REVISADO_POR_XX"] = "Andrés Amado"
    data_client["XX_ROL_REVISADOR_XX"] = "Profesional de proyectos"
    data_client["XX_AUTORIZADO_POR_XX"] = "Claudia Calderón"
    data_client["XX_AUTORIZADO_POR_ROL_XX"] = "Directora de Proyectos"
    data_client["XX_INFORME_NUMERO_XX"] = 12701
    data_client["XX_FECHA_EMISION_INFORME_XX"] = str(get_today_date()).split(" ")[0]

    """data_client["client_name"] = main_sheet["B2"].value or "Not client found"
    data_client["contact_client_name"] = main_sheet["B4"].value or "Not client contact name found"
    data_client["client_contact"] = main_sheet["B6"].value or "Not client contact found"
    data_client["prepared_by"] = main_sheet["E2"].value or "No manufacturer found"
    data_client["report_type"] = "INFORME DE MONITOREO"

    data_client["municipality"] = main_sheet["B7"].value or "No municipality found"

    #Storage sampling importation data
    sampling_data["sampling_site"] = main_sheet["E4"].value or "Not sampling site"""

    sampling_data["XX_FECHA_MONITOREO_XX"] = str(main_sheet["E5"].value).split(" ")[0] or "Not sampling date"
    sampling_data["XX_FECHA_MONITOREO_LITERAL_XX"] = date_literal(str(main_sheet["E5"].value).split(" ")[0])
    sampling_data["XX_MES_LITERAL_XX"] = date_literal(str(main_sheet["E5"].value).split(" ")[0]).split(" ")[2]
    sampling_data["XX_PLAN_DE_MUESTREO_XX"] = main_sheet["E9"].value or "Not sampling plan"

    return data_client, sampling_data


def read_chain_of_custody(workbook) -> dict:
    # Dictionary: key = chemilab code, value = sample info dict
    samples = {}

    try:
        chain_of_custody_sheet = workbook["CADENA DE CUSTODIA"]
    except KeyError:
        print("Chain of custody sheet not found")
        return samples

    INITIAL_ROW = 23
    MAX_EMPTY_ROWS = 10  # Detenerse después de 10 filas vacías consecutivas
    empty_row_count = 0

    # Recorrer TODAS las filas hasta encontrar muchas vacías consecutivas
    while empty_row_count < MAX_EMPTY_ROWS:

        chemilab_code = chain_of_custody_sheet[f"A{INITIAL_ROW}"].value

        # Si la fila tiene datos
        if chemilab_code:
            sample = {}
            sample["chemilab_code"] = chemilab_code

            # Sample identification
            sample_identification = chain_of_custody_sheet[f"C{INITIAL_ROW}"].value
            sample["sample_identification"] = sample_identification

            # Date of the sample
            sample_year = chain_of_custody_sheet[f"G{INITIAL_ROW}"].value
            sample_month = chain_of_custody_sheet[f"H{INITIAL_ROW}"].value
            sample_day = chain_of_custody_sheet[f"I{INITIAL_ROW}"].value

            if all([sample_year, sample_month, sample_day]):
                sample["sample_date"] = f"{sample_year}-{sample_month}-{sample_day}"
            else:
                sample["sample_date"] = "No date found"

            # Add the sample to the dictionary
            samples[chemilab_code] = sample

            # Resetear contador de filas vacías
            empty_row_count = 0
        else:
            # Contar filas vacías consecutivas
            empty_row_count += 1

        # Avanzar de 1 en 1 para no saltarse filas
        INITIAL_ROW += 1

    print(f"✅ Se encontraron {len(samples)} muestras")

    # Leer las horas de muestreo de la cadena de vigilancia puntual
    try:
        punctual_sheet = workbook["CADENA DE VIGILANCIA PUNTUAL"]
        hour_1 = punctual_sheet["F71"].value
        hour_2 = punctual_sheet["R71"].value

        # Agregar las horas a las dos primeras muestras
        sample_keys = list(samples.keys())

        if len(sample_keys) >= 1 and hour_1:
            samples[sample_keys[0]]["sample_hour"] = str(hour_1)

        if len(sample_keys) >= 2 and hour_2:
            samples[sample_keys[1]]["sample_hour"] = str(hour_2)

        print(f"✅ Se agregaron las horas de muestreo a las primeras dos muestras")

    except KeyError:
        print("⚠️ No se pudo leer la hoja de vigilancia puntual para agregar las horas")
    except Exception as e:
        print(f"⚠️ Error al agregar las horas: {e}")

    return samples


def read_punctual_surveillance_chain(workbook) -> dict:
    punctual_data = {}

    try:
        punctual_surveillance_chain_sheet = workbook["CADENA DE VIGILANCIA PUNTUAL"]

        if punctual_surveillance_chain_sheet:

            print("SI EXISTE LA HOJA")

        else:

            print("NO EXISTE LA HOJA SURVEILLANCE")

        ROW_TYPE_WATER = 17

        dict_water_types = {

            "D": "A.R.I",
            "F": "A.R.Nd",
            "H": "A.R.D",
            "J": "Agua Superficial",
            "L": "Agua subterranea",
            "N": "A.POT",
            "P": "A.MAR"

        }

        columns_row_type_w = ["D", "F", "H", "J", "L", "N", "P"]

        for column in columns_row_type_w:

            cell_coord = f"{column}{ROW_TYPE_WATER}"
            cell_value = punctual_surveillance_chain_sheet[cell_coord].value

            if cell_value == 'X' or cell_value == 'x':
                punctual_data["water_type"] = dict_water_types[column]

    except KeyError:

        print("Punctual surveillance sheet sheet not found")
        return punctual_data

    return punctual_data


def data_constructor(workbook) -> dict:
    try:

        main_data, sampling_data = read_main_sheet_excel(workbook)
        samples_data = read_chain_of_custody(workbook)
        surveillance_data = read_punctual_surveillance_chain(workbook)

        return {

            "main_data": main_data,
            "sampling_data": sampling_data,
            "samples": samples_data,
            "surveillance_data": surveillance_data
        }

    except IndexError:

        raise ValueError("The file has no sheets available")

    except KeyError as e:

        raise ValueError(f"Invalid cell reference")

    except AttributeError:

        raise TypeError("The excel worksheet is not valid")

    except InvalidFileException:

        raise ValueError("The file is not a valid excel file")

        raise Exception