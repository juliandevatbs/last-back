def sampling_site_fixing(text_to_fix: str):
    if not text_to_fix:
        return False

    # Buscar la posición de "agua" (sin asignar a fixed_text)
    agua_position = text_to_fix.lower().find("agua")

    if agua_position != -1:
        # Si encuentra "agua", tomar el texto antes de esa posición
        fixed_text = text_to_fix[:agua_position].rstrip("- \n").strip()
    else:
        # Si no encuentra "agua", usar el texto completo o manejar según tu lógica
        fixed_text = text_to_fix.strip()  # o return False, según lo que necesites

    print(f"TEXTO ARREGLADO -> {fixed_text}")
    return fixed_text