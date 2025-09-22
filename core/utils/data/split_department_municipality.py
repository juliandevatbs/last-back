def split_department_municipality(full_text: str, separators=['-', '/'], part='second'):


    for separator in separators:
        if separator in full_text:
            parts = full_text.split(separator, 1)  # Split solo en el primer separador

            if part.lower() == 'first':
                return parts[0].strip()
            elif part.lower() == 'second' and len(parts) > 1:
                return parts[1].strip()

    return full_text