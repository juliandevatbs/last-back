def read_main_sheet(workbook, sheet_name: str) -> dict:

    # read basic data sheet

    try:
        main_sheet = workbook[sheet_name]

        # Storage the sheet data
        basic_data = {}

        # Sub dicts for data
        client_data  = {}
        sampling_basic_data = {}

        # Read client data
        client_data["XX_RAZON_SOCIAL_XX"] = main_sheet["B2"].value
        client_data["XX_DIRECCION_XX"] = main_sheet["B3"].value
        client_data["XX_PERSONA_CONTACTO_XX"] = main_sheet["B4"].value
        client_data["XX_NIT_XX"] = main_sheet["B5"].value
        client_data["XX_TELEFONO_XX"] = str(main_sheet["B6"].value)
        client_data["XX_MUNICIPIO/DEPARTAMENTO_XX"] = main_sheet["B7"].value
        client_data["XX_COTIZACION_NUM_XX"] = main_sheet["B8"].value
        client_data["XX_ACTIVIDAD_ECONOMICA_XX"] = main_sheet["B9"].value
        client_data["XX_PLAN_MUESTRO_AGUAS_XX"] = main_sheet["B10"].value

        basic_data["client_data"] = client_data


        # Read sampling basic data
        sampling_basic_data["XX_RESPONSABLE_MUESTREO_XX"] = main_sheet["E2"].value
        sampling_basic_data["XX_MUNICIPIO/DEPARTAMENTO_MUESTREO_XX"] = main_sheet["E3"].value
        sampling_basic_data["XX_SITIO_MUESTREO_XX"] = main_sheet["E4"].value
        sampling_basic_data["XX_FECHA_MUESTREO_XX"] = main_sheet["E5"].value

        basic_data["sampling_basic_data"] = sampling_basic_data

        print(basic_data)

        return basic_data


    except KeyError:

        print(f"Sheet {sheet_name} not found please review")
        return {}

    except Exception as e:

        print("Error opening the sheet")
        return {}


