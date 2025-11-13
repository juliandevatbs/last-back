def incertidumbre_auto(value):

    try:
        # Convertir a float para no perder decimales
        val = float(value)
    except (ValueError, TypeError):
        return "±0"

    incertidumbre = val / 100

    return f"±{incertidumbre:.4f}"
