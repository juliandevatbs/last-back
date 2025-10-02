from num2words import num2words


def int_to_string_relative(number_int: int) -> str:


    # Convert int to strings relative ex 4 -> cuatro


    return num2words(number_int, lang='es')