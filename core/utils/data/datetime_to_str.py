from datetime import datetime


def datetime_to_string(fecha: datetime, formato: str = '%Y-%m-%d') -> str:

    return fecha.strftime(formato)


