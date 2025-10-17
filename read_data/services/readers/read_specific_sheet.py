import xlwings as xw

def read_sample_information(file_path: str) -> dict:
    samples_information = {}

    try:
        app = xw.App(visible=False)
        wb = app.books.open(file_path)

        sheet_names = [s.name for s in wb.sheets]
        if "AFLUENTE 1" not in sheet_names or "EFLUENTE 2" not in sheet_names or "CADENA DE VIGILANCIA PUNTUAL" not in sheet_names:
            print("NO EXISTEN CADENA DE VIGILANCIA PUNTUAL")
            wb.close()
            app.quit()
            return samples_information

        sheet_one = wb.sheets["AFLUENTE 1"]
        sheet_two = wb.sheets["EFLUENTE 2"]
        sheet_three = wb.sheets["CADENA DE VIGILANCIA PUNTUAL"]


        sheet_one_texts = {}
        for shape in sheet_one.shapes:
            if shape.text:
                text = shape.text.strip()
                if ":" in text:
                    text = text.split(":", 1)[1].strip()
                sheet_one_texts["descripcion_punto"] = text

        print("Leyendo formas de EFLUENTE 2...")
        sheet_two_texts = {}
        for shape in sheet_two.shapes:
            if shape.text:
                text = shape.text.strip()
                if ":" in text:
                    text = text.split(":", 1)[1].strip()
                sheet_two_texts["descripcion_punto"] = text

        sheet_three_texts = {}
        shapes_found = 0
        for shape in sheet_three.shapes:
            shapes_found += 1
            print(f"Forma encontrada #{shapes_found}: texto='{shape.text}'")
            if shape.text:
                text = shape.text.strip()
                if ":" in text:
                    text = text.split(":", 1)[1].strip()
                    sheet_three_texts["descripcion_punto"] = text
                else:
                    # Si no tiene ":", guardar el texto completo
                    print(f"Advertencia: forma sin ':' - {text}")
                    sheet_three_texts["descripcion_punto"] = text

        print(f"Total de formas encontradas en CADENA: {shapes_found}")




        # Construir resultado
        samples_information = {
            "AFLUENTE_1": sheet_one_texts,
            "EFLUENTE_2": sheet_two_texts,
            "CADENA DE VIGILANCIA PUNTUAL": sheet_three_texts
        }



        wb.close()
        app.quit()

    except Exception as e:
        print(f"Error al leer shapes con xlwings: {e}")
        try:
            app.quit()
        except:
            pass

    return samples_information


def read_specific_sheet(file_path, sheet_name: str):
    text_description = ''
    app = None
    wb = None

    try:
        app = xw.App(visible=False)
        wb = app.books.open(file_path)

        sheet_names = [s.name for s in wb.sheets]

        print(sheet_names)

        if sheet_name not in sheet_names:
            return ''

        sheet_one = wb.sheets[sheet_name]

        for shape in sheet_one.shapes:
            if shape.text:
                text = shape.text.strip()
                if ":" in text:
                    text = text.split(":", 1)[1].strip()
                text_description += text

        return text_description

    except KeyError:
        print(f"Sheet {sheet_name} not found please review")
        return ''

    except Exception as e:
        print(f"Error opening the sheet: {e}")
        return ''

    finally:
        try:
            if wb:
                wb.close()
            if app:
                app.quit()
        except:
            pass


def read_specific_sheet_data(workbook, sheet_name: str, day_type_start_row=14, nublado_col='F', soleado_col='H', lluvioso_col='J', temp_col = "M", hum_col="O", alt_col="T"):


    try:

        specific_sheet = workbook[sheet_name]



        nublado = specific_sheet[f"{nublado_col}{day_type_start_row}"].value
        soleado = specific_sheet[f"{soleado_col}{day_type_start_row}"].value
        lluvioso = specific_sheet[f"{lluvioso_col}{day_type_start_row}"].value

        clima = None

        if nublado and str(nublado).strip().upper() == 'X':
            clima = 'Dia Nublado'
        elif soleado and str(soleado).strip().upper() == 'X':
            clima = 'Dia Soleado'
        elif lluvioso and str(lluvioso).strip().upper() == 'X':
            clima = 'Dia Lluvioso'
        else:
            clima = 'No especificado'

        temperatura_ambiente = f"{specific_sheet[f"{temp_col}{day_type_start_row}"].value}°C"
        humedad_relativa = f"{specific_sheet[f"{hum_col}{day_type_start_row}"].value}%"
        altitud = f"{specific_sheet[f"T72"].value} m.s.n.m"


        return clima, temperatura_ambiente, humedad_relativa, altitud





    except KeyError:

        print(f"Sheet {sheet_name} not found please review")
        return {}

    except Exception as e:

        print("Error opening the sheet")
        return {}









