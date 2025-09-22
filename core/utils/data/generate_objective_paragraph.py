def generate_objective_paragraph(sampling_site, client_name, analysis_type, water_type, license_number, polygon_name,
                                 municipality, department, facility_type):
    """
    Genera párrafo de objetivo limpio y bien formado
    """

    # Limpiar y preparar datos
    sampling_site = sampling_site.title() if sampling_site else ""
    client_name = client_name.upper() if client_name else ""
    analysis_type = analysis_type.lower() if analysis_type else "fisicoquímica"
    municipality = municipality.title() if municipality else ""
    department = department.title() if department else ""

    # Construir acción principal - CORREGIR LA DUPLICACIÓN
    if "caracterización" in analysis_type:
        # Si ya viene "caracterización fisicoquímica", usar solo "calidad fisicoquímica"
        action = "Determinar calidad fisicoquímica"
    elif "microbiológica" in analysis_type:
        action = f"Determinar la calidad {analysis_type}"
    else:
        action = f"Determinar calidad {analysis_type}"

    # Determinar tipo de agua y fuente
    if "superficial" in water_type or "agua superficial" in sampling_site.lower():
        water_description = f"de agua superficial del {sampling_site}"
    elif "residual" in water_type:
        water_description = f"del agua residual del {sampling_site}"
    else:
        water_description = f"de agua del {sampling_site}"

    # Construir párrafo
    parts = [action, water_description]

    # Agregar licencia si existe
    if license_number and license_number.strip():
        license_info = f"bajo licencia {license_number}"
        if polygon_name:
            license_info += f" ({polygon_name})"
        parts.append(license_info)

    # Agregar operador
    if client_name:
        parts.append(f"operada por {client_name}")

    # Agregar ubicación
    location_parts = []
    if municipality and municipality.strip():
        location_parts.append(f"en el municipio de {municipality}")
    if department and department.strip() and department != municipality:
        location_parts.append(f"en el departamento de {department}")

    if location_parts:
        parts.append(f"localizada {', '.join(location_parts)}")

    # Unir y limpiar
    objective = " ".join(parts) + "."

    # Limpiar espacios dobles y otros problemas
    objective = " ".join(objective.split())

    return objective