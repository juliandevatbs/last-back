from openpyxl.reader.excel import load_workbook
from openpyxl.utils.exceptions import InvalidFileException



def read_main_sheet_excel(workbook) -> dict :


        main_sheet = workbook.worksheets[0]

        #print(main_sheet)

        # Dict to storage the data
        data_client = {}
        sampling_data = {}

        # Storage client and main report data
        data_client["client_name"] = main_sheet["B2"].value or "Not client found"
        data_client["contact_client_name"] = main_sheet["B4"].value or "Not client contact name found"
        data_client["client_contact"] = main_sheet["B6"].value or "Not client contact found"
        data_client["prepared_by"] = main_sheet["E2"].value or "No manufacturer found"
        data_client["report_type"] = "INFORME DE MONITOREO"
        data_client["report_number"] = 12701
        data_client["municipality"] = main_sheet["B7"].value or "No municipality found"

        #Storage sampling importation data
        sampling_data["sampling_site"] = main_sheet["E4"].value or "Not sampling site"
        sampling_data["sampling_date"] = main_sheet["E5"].value or "Not sampling date"


        return data_client, sampling_data





def read_chain_of_custody(workbook) -> dict:

    # Dictionary the key is the chemilab code of the sample and the value is a dict with the sample info
    samples = {}


    try:

        chain_of_custody_sheet = workbook["CADENA DE CUSTODIA"]

    except KeyError:

        print("Chain of custody sheet not found")
        return samples

    INITIAL_ROW = 23

    # The while stops if the first column with the chemilab code doesn´t have info
    while True:

        sample = {}


        chemilab_code = chain_of_custody_sheet[f"A{INITIAL_ROW}"].value

        # Verify if the row has data
        if chemilab_code:

            sample["chemilab_code"] = chemilab_code

            #Sample idenification
            sample_identification  = chain_of_custody_sheet[f"C{INITIAL_ROW}"].value
            sample["sample_identification"] = sample_identification

            #Date of the sample
            sample_year = chain_of_custody_sheet[f"G{INITIAL_ROW}"].value
            sample_month = chain_of_custody_sheet[f"H{INITIAL_ROW}"].value
            sample_day = chain_of_custody_sheet[f"I{INITIAL_ROW}"].value

            if all([sample_year, sample_month, sample_day]):

                sample["sample_date"] = f"{sample_year}-{sample_month}-{sample_day}"

            else:

                sample["sample_date"] = "No date found"

        else:

            break

        # Add the sample to the samples dictionary
        samples[chemilab_code] = sample

        INITIAL_ROW += 1

    return samples


def read_punctual_surveillance_chain(workbook) -> dict:


    punctual_data = {}


    try:
        punctual_surveillance_chain_sheet = workbook["CADENA DE VIGILANCIA PUNTUAL"]

        ROW_TYPE_WATER = 17

        dict_water_types = {

            "D": "A.R.I",
            "F": "A.R.Nd",
            "H": "A.R.D",
            "J": "Agua Superficial",
            "L": "Agua subterranea",
            "N":  "A.POT",
            "P": "A.MAR"

        }

        columns_row_type_w = ["D", "F", "H", "J", "L", "N", "P"]


        for column in columns_row_type_w:

            cell_coord = f"{column}{ROW_TYPE_WATER}"
            cell_value = punctual_surveillance_chain_sheet[cell_coord].value

            if cell_value == 'X':

                punctual_data["water_type"] = dict_water_types[column]

    except KeyError:

        print("Punctual surveillance sheet sheet not found")
        return punctual_data

    return punctual_data




def data_constructor(workbook) -> dict:

    try:

        main_data, client_data = read_main_sheet_excel(workbook)
        samples_data = read_chain_of_custody(workbook)
        surveillance_data = read_punctual_surveillance_chain(workbook)

        return {

            "main_data" : main_data,
            "sampling_data": client_data,
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
































