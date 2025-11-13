from core.utils.data.incertidumbre_autom import incertidumbre_auto


def ph_reader(workbook, sheet_name: str, columns: dict, initial_row: int) -> dict:
    try:
        sheet = workbook[sheet_name]
        data_rows = {}
        current_row = initial_row

        sum_ph = 0
        sum_caudal = 0
        min_ph = float('inf')
        max_ph = float('-inf')
        min_caudal = float('inf')
        max_caudal = float('-inf')

        while True:
            index_cell = f"{columns['COLUMNA_INDICE']}{current_row}"
            index_val = sheet[index_cell].value

            if index_val is None or str(index_val).strip() == '':
                break

            ph_val = sheet[f"{columns['PH']}{current_row}"].value
            caudal_val = sheet[f"{columns['CAUDAL']}{current_row}"].value
            solidos_val = sheet[f"{columns['SOLIDOS_SEDIMENTABLES']}{current_row}"].value
            hora_val = sheet[f"{columns['HORA']}{current_row}"].value

            if ph_val is not None:

                if hasattr(hora_val, 'strftime'):
                    hora_str = hora_val.strftime('%H:%M')
                else:
                    hora_str = str(hora_val) if hora_val else ""

                data_rows[index_val] = {
                    "hour": hora_str,
                    "ph": ph_val,
                    "caudal": str(caudal_val)[:4] if caudal_val else "",
                    "solidos_sedimentables": solidos_val,
                    "incertidumbre": incertidumbre_auto(ph_val)
                }

                sum_ph += ph_val
                sum_caudal += caudal_val if caudal_val else 0
                min_ph = min(min_ph, ph_val)
                max_ph = max(max_ph, ph_val)

                if caudal_val:
                    min_caudal = min(min_caudal, caudal_val)
                    max_caudal = max(max_caudal, caudal_val)

            current_row += 1

        if data_rows:
            count = len(data_rows)
            data_rows["_metadata"] = {
                "media_valores": round(sum_ph / count, 3) if count > 0 else 0,
                "media_incertidumbre": incertidumbre_auto(sum_ph / count) if count > 0 else "",
                "media_caudal": round(sum_caudal / count, 3) if count > 0 else 0,
                "valor_minimo_reportado": min_ph if min_ph != float('inf') else 0,
                "valor_maximo_reportado": max_ph if max_ph != float('-inf') else 0,
                "min_valor_reportado_caudales": str(round(min_caudal, 2))[:4] if min_caudal != float('inf') else "0",
                "max_valor_reportado_caudales": str(round(max_caudal, 2))[:4] if max_caudal != float('-inf') else "0"
            }

        return data_rows

    except Exception as e:
        print(f"Error en ph_reader para {sheet_name}: {e}")
        return {}