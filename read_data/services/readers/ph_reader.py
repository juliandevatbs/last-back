from math import inf
from core.utils.data.incertidumbre_autom import incertidumbre_auto


def ph_reader(workbook, sheet_name: str, columns: dict, initial_row: int) -> dict:
    try:
        data_rows = {}
        sheet_to_read = workbook[sheet_name]
        current_row = initial_row
        sum_reported_values = 0
        sum_caudal_values = 0

        min_valor_reportado = float(inf)
        min_valor_reportado_solidos = '<0,1'
        min_valor_reportado_caudales = float(inf)

        max_valor_reportado = float(-inf)
        max_valor_reportado_solidos = '<0,1'
        max_valor_reportado_caudales = float(-inf)

        while True:
            init_cell_address = f"{columns.get('initial_row')}{current_row}"
            init_cell_value = sheet_to_read[init_cell_address].value

            if init_cell_value is None or str(init_cell_value).strip() == '':
                break

            data_row = {}

            ph_value = sheet_to_read[f"{columns.get('ph_column')}{current_row}"].value
            caudal_value = sheet_to_read[f"{columns.get('caudal_column')}{current_row}"].value
            solidos_value = sheet_to_read[f"{columns.get('solidos_sedimentables_column')}{current_row}"].value

            data_row["hour"] = sheet_to_read[f"{columns.get('hour_column')}{current_row}"].value
            data_row["ph"] = ph_value
            data_row["caudal"] = str(caudal_value)[:4]
            data_row["solidos_sedimentables"] = solidos_value
            data_row["incertidumbre"] = incertidumbre_auto(ph_value)

            sum_reported_values += ph_value
            sum_caudal_values += caudal_value

            if ph_value < min_valor_reportado:
                min_valor_reportado = ph_value

            if ph_value > max_valor_reportado:
                max_valor_reportado = ph_value

            if caudal_value < min_valor_reportado_caudales:
                min_valor_reportado_caudales = caudal_value

            if caudal_value > max_valor_reportado_caudales:
                max_valor_reportado_caudales = caudal_value

            data_rows[init_cell_value] = data_row
            current_row += 1

        num_registros = len(data_rows)

        if num_registros > 0:
            media_valores = sum_reported_values / num_registros
            media_caudal = sum_caudal_values / num_registros

            data_rows["_metadata"] = {
                "media_valores": f"{media_valores:.3}",
                "media_incertidumbre": incertidumbre_auto(media_valores),
                "media_caudal": f"{media_caudal:.3}",
                "valor_minimo_reportado": min_valor_reportado,
                "valor_maximo_reportado": max_valor_reportado,
                "min_valor_reportado_solidos": min_valor_reportado_solidos,
                "max_valor_reportado_solidos": max_valor_reportado_solidos,
                "suma_total": sum_reported_values,
                "num_registros": num_registros,
                "min_valor_reportado_caudales": str(min_valor_reportado_caudales)[:4],
                "max_valor_reportado_caudales": str(max_valor_reportado_caudales)[:4]
            }

        return data_rows

    except KeyError as e:
        print(f"Error: Sheet '{sheet_name}' not found in workbook. Details: {e}")
        return {}

    except Exception as e:
        print(f"Error reading sheet '{sheet_name}': {e}")
        import traceback
        traceback.print_exc()
        return {}