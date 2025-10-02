from datetime import datetime


def date_literal(fecha_monitoreo) -> str:

    # 2025-20-11 -> 20 de noviembre de 2025

    month_mapping = {

        '01': 'Enero',
        '02': 'Febrero',
        '03': 'Marzo',
        '04': 'Abril',
        '05': 'Mayo',
        '06': 'Junio',
        '07': 'Julio',
        '08': 'Agosto',
        '09': 'Septiembre',
        '10': 'Octubre',
        '11': 'Noviembre',
        '12': 'Diciembre'

    }

    if isinstance(fecha_monitoreo, datetime):

        month = month_mapping.get(f"{fecha_monitoreo.month:02d}")
        year = fecha_monitoreo.year
        day = fecha_monitoreo.day

        literal_date = f"{day} de {month} de {year}"
        return literal_date

    elif isinstance(fecha_monitoreo, str):

        fecha_monitoreo = fecha_monitoreo.split('-')
        month =month_mapping.get(fecha_monitoreo[1])
        year = fecha_monitoreo[0]
        day = fecha_monitoreo[-1]


        literal_date = f"{day} de {month} de {year}"

        return literal_date

    else:

        return ''



