def sampling_site_fixing(text_to_fix: str):

    fixed_text = ""


    if not text_to_fix:

        return False

    fixed_text = text_to_fix.lower().find("agua")

    if fixed_text != -1:

        fixed_text = text_to_fix[:fixed_text].rstrip("- \n").strip()


    print(f"TEXTO ARREGLADO -> {fixed_text}")
    return fixed_text



