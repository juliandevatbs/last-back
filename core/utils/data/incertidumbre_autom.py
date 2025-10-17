def incertidumbre_auto(value):

    print(f"VALOR QUE LLEGA {value}")

    try:
        # Convertir a float para no perder decimales
        val = float(value)
    except (ValueError, TypeError):
        return "±0"

    incertidumbre = val / 100

    return f"±{incertidumbre:.4f}"
