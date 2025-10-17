from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import glob


def write_monitoring_table(doc, font: str, size: int, bold: bool, samples_data: dict, basic_data: dict):
    """
    Escribe los datos de muestras en la tabla de monitoreo

    Args:
        doc: Documento de python-docx
        font: Nombre de la fuente
        size: Tamaño de fuente en puntos
        bold: Si el texto debe ser en negrita
        samples_data: Diccionario con los datos de las muestras

    Returns:
        bool: True si fue exitoso, False si hubo error
    """

    # Buscar la tabla de monitoreo por título
    target_title = "Puntos de monitoreo"
    monitoring_tables = _find_monitoring_tables(doc, target_title)

    if not monitoring_tables:
        print("❌ No se encontraron tablas de monitoreo")
        return False

    print(f"✅ Encontradas {len(monitoring_tables)} tablas de monitoreo")

    # Estructura de la tabla
    FIRST_DATA_ROW = 3
    ROWS_PER_SAMPLE = 2

    # Posiciones de columnas
    COL_SAMPLE_ID = 0
    COL_DATE = 1
    COL_HOUR = 2
    COL_NAME = 3
    COL_EAST = 4
    COL_NORTH = 5
    COL_PHOTO = 6
    COL_DESCRIPTION = 0

    # Obtener lista de muestras (excluyendo 'OSI')
    sample_list = [(k, v) for k, v in samples_data.items() if k != 'OSI']

    # Calcular cuántas muestras caben en cada tabla
    available_rows_first_table = len(monitoring_tables[0].rows) - FIRST_DATA_ROW
    samples_in_first_table = available_rows_first_table // ROWS_PER_SAMPLE

    print(f"📄 Primera tabla puede contener {samples_in_first_table} muestras")

    # Si hay segunda tabla, calcular su capacidad
    samples_in_second_table = 0
    if len(monitoring_tables) > 1:
        available_rows_second_table = len(monitoring_tables[1].rows) - FIRST_DATA_ROW
        samples_in_second_table = available_rows_second_table // ROWS_PER_SAMPLE
        print(f"📄 Segunda tabla puede contener {samples_in_second_table} muestras")

    # Escribir cada muestra
    for i, (sample_id, sample_data) in enumerate(sample_list):

        # Determinar en qué tabla escribir
        if i < samples_in_first_table:
            current_table = monitoring_tables[0]
            sample_index_in_table = i
            table_type = "PRIMERA"
        elif len(monitoring_tables) > 1 and i < samples_in_first_table + samples_in_second_table:
            current_table = monitoring_tables[1]
            sample_index_in_table = i - samples_in_first_table
            table_type = "SEGUNDA"
        else:
            print(f"❌ No hay suficiente espacio para la muestra {i + 1}")
            break

        # Calcular índices de filas
        data_row_idx = FIRST_DATA_ROW + (sample_index_in_table * ROWS_PER_SAMPLE)
        desc_row_idx = data_row_idx + 1

        print(f"\nMuestra {i + 1} ({sample_id}):")
        print(f"  Escribiendo en fila {data_row_idx} de la tabla {table_type}")

        # ✅ Solo escribir PLAN DE MUESTREO si es la primera muestra de cada tabla
        if sample_index_in_table == 0:
            prev_row_idx = data_row_idx - 1

            if prev_row_idx >= FIRST_DATA_ROW - 1:
                prev_row = current_table.rows[prev_row_idx]

                write_cell_safe(
                    prev_row.cells[0],
                    [(f"PLAN DE MUESTREO: {basic_data['client_data']['XX_PLAN_MUESTRO_AGUAS_XX']}", True)],
                    font, size, False, WD_ALIGN_PARAGRAPH.LEFT
                )

        if data_row_idx >= len(current_table.rows):
            print(f"❌ No hay suficientes filas para muestra {i + 1}")
            break

        # Obtener la fila de datos
        data_row = current_table.rows[data_row_idx]

        # 🔍 DEBUG: Verificar número de columnas
        num_columns = len(data_row.cells)
        print(f"  📊 Número de columnas en la fila: {num_columns}")

        # Formatear la fecha
        sample_date = f"{sample_data.get('sample_day', ''):02d}/{sample_data.get('sample_month', ''):02d}/{sample_data.get('sample_year', ''):02d}"

        # Escribir código de muestra
        write_cell_safe(
            data_row.cells[COL_SAMPLE_ID],
            sample_data.get('chemilab_code', ''),
            font, size, False, WD_ALIGN_PARAGRAPH.CENTER
        )

        # Escribir fecha
        write_cell_safe(
            data_row.cells[COL_DATE],
            sample_date,
            font, size, False, WD_ALIGN_PARAGRAPH.CENTER
        )

        # Escribir hora
        write_cell_safe(
            data_row.cells[COL_HOUR],
            sample_data.get('sampler_hour', ''),
            font, size, False, WD_ALIGN_PARAGRAPH.CENTER
        )

        # Escribir nombre del punto
        write_cell_safe(
            data_row.cells[COL_NAME],
            [(f"\n{sample_data.get('sample_identification', '')}\n", False)],
            font, size, False, WD_ALIGN_PARAGRAPH.CENTER
        )

        # ✅ Insertar imagen en la columna de fotografía
        if COL_PHOTO < num_columns:  # ✅ Verificar que la columna existe
            image_path = _get_sample_image(i)
            if image_path and os.path.exists(image_path):
                try:
                    print(f"  🔧 Intentando acceder a celda #{COL_PHOTO}...")
                    photo_cell = data_row.cells[COL_PHOTO]
                    print(f"  ✅ Celda obtenida correctamente")

                    # Limpiar la celda de foto
                    clear_cell_safe(photo_cell)
                    print(f"  ✅ Celda limpiada")

                    # Agregar la imagen
                    if not photo_cell.paragraphs:
                        photo_cell.add_paragraph()

                    paragraph = photo_cell.paragraphs[0]
                    run = paragraph.add_run()
                    run.add_picture(image_path, width=Inches(1.5))
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    print(f"  ✅ Imagen insertada: {os.path.basename(image_path)}")
                except IndexError as e:
                    print(f"  ⚠️ Error de índice al insertar imagen: {e}")
                    print(f"  🔍 Celdas disponibles: {len(data_row.cells)}")
                    print(f"  🔍 Intentando acceder a índice: {COL_PHOTO}")
                except Exception as e:
                    print(f"  ⚠️ Error al insertar imagen: {e}")
                    import traceback
                    traceback.print_exc()
            else:
                print(f"  ⚠️ No se encontró imagen en: assets/images/{_get_folder_name(i)}")
        else:
            print(f"  ⚠️ Columna de foto (#{COL_PHOTO}) no existe. La tabla solo tiene {num_columns} columnas")

        # Escribir descripción (en la siguiente fila)
        if desc_row_idx < len(current_table.rows):
            desc_row = current_table.rows[desc_row_idx]

            description_text = sample_data.get('sample_description', '')

            write_cell_safe(
                desc_row.cells[COL_DESCRIPTION],
                [('Descripción del punto: ', True),
                 (f"{description_text[2:]}\n", False),
                 ("\nCondiciones ambientales: ", True),
                 (
                     f"{sample_data.get('sample_weather')} - Temperatura ambiente: {sample_data.get('sample_temperature')} - Humedad relativa: {sample_data.get('sample_humidity')} - Altitud: {sample_data.get('sample_altitude')}")
                 ],
                font, size, False, WD_ALIGN_PARAGRAPH.LEFT
            )

        print(f"  ✅ Muestra {sample_id} escrita exitosamente")

    return True


def _find_monitoring_tables(doc, target_title):
    """
    Encuentra todas las tablas de monitoreo basándose en el título
    """
    monitoring_tables = []

    # Buscar por el título específico
    for element in doc.element.body:
        if element.tag.endswith('p'):
            para_text = ""
            for run in element.findall('.//w:t',
                                       {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                if run.text:
                    para_text += run.text

            if target_title.lower() in para_text.lower():
                print(f"✅ Título encontrado: '{para_text}'")

                # Buscar todas las tablas que siguen a este título
                next_element = element.getnext()
                tables_found = 0

                while next_element is not None and tables_found < 3:
                    if next_element.tag.endswith('tbl'):
                        for table in doc.tables:
                            if table._element == next_element:
                                if _is_monitoring_table(table):
                                    monitoring_tables.append(table)
                                    tables_found += 1
                                    print(f"✅ Tabla #{tables_found} encontrada con {len(table.rows)} filas")
                                break

                    next_element = next_element.getnext()

                break

    # Si no encontramos por título, buscar por estructura
    if not monitoring_tables:
        monitoring_tables = _find_tables_by_structure(doc)

    return monitoring_tables


def _is_monitoring_table(table):
    """
    Verifica si una tabla tiene la estructura de tabla de monitoreo
    """
    if len(table.rows) < 4:
        return False

    first_row_text = ""
    for cell in table.rows[0].cells:
        first_row_text += cell.text.upper() + " "

    keywords = ['CÓDIGO', 'FECHA', 'HORA', 'IDENTIFICACIÓN', 'COORDENADAS', 'FOTOGRAFÍA']
    found_keywords = sum(1 for keyword in keywords if keyword in first_row_text)

    return found_keywords >= 3


def _find_tables_by_structure(doc):
    """
    Busca tablas por su estructura cuando no se encuentra el título
    """
    monitoring_tables = []

    for i, table in enumerate(doc.tables):
        if _is_monitoring_table(table):
            monitoring_tables.append(table)
            print(f"✅ Tabla #{i + 1} identificada por estructura con {len(table.rows)} filas")

    return monitoring_tables


def write_cell_safe(cell, text, font_name, font_size, bold, align, space_after=0):
    """
    Escribe texto en una celda con formato específico
    """
    # Limpiar celda de forma segura
    clear_cell_safe(cell)

    # Agregar nuevo contenido
    p = cell.add_paragraph()

    if isinstance(text, list):
        # Si text es una lista de tuplas (text_part, is_bold)
        for item in text:
            if isinstance(item, (tuple, list)) and len(item) == 2:
                text_part, is_bold = item
                run = p.add_run(text_part)
                run.font.name = font_name
                run.font.size = Pt(font_size)
                run.bold = is_bold
            else:
                run = p.add_run(str(item))
                run.font.name = font_name
                run.font.size = Pt(font_size)
                run.bold = bold
    else:
        # Si text es un string simple
        run = p.add_run(str(text))
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.bold = bold

    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1


def clear_cell_safe(cell):
    """
    Limpia el contenido de una celda preservando imágenes
    """
    images = []

    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            if run._element.xpath('.//w:drawing'):
                images.extend(run._element.xpath('.//w:drawing'))

    # Eliminar párrafos
    paragraphs_to_remove = list(cell.paragraphs)
    for p in paragraphs_to_remove:
        p._element.getparent().remove(p._element)

    # Restaurar imágenes
    if images:
        p = cell.add_paragraph()
        for img in images:
            p._element.append(img)


def _get_sample_image(sample_index):
    """
    Obtiene la ruta de la primera imagen según el índice de la muestra

    Args:
        sample_index: Índice de la muestra (0, 1, 2)

    Returns:
        str: Ruta completa de la imagen o None si no se encuentra
    """
    folder_name = _get_folder_name(sample_index)
    if not folder_name:
        return None

    # Obtener la ruta absoluta del proyecto (donde está BackEnd)
    current_dir = os.path.dirname(os.path.abspath(__file__))
    # Subir hasta llegar a BackEnd
    while os.path.basename(current_dir) != "BackEnd" and current_dir != os.path.dirname(current_dir):
        current_dir = os.path.dirname(current_dir)

    # Ruta base de las imágenes
    base_path = os.path.join(current_dir, "assets", "images", folder_name)

    print(f"  🔍 Buscando imagen en: {base_path}")

    # Verificar que la carpeta existe
    if not os.path.exists(base_path):
        print(f"  ❌ La carpeta no existe: {base_path}")
        return None

    # Buscar la primera imagen (extensiones comunes)
    image_extensions = ['*.jpg', '*.jpeg', '*.png', '*.JPG', '*.JPEG', '*.PNG']

    for ext in image_extensions:
        images = glob.glob(os.path.join(base_path, ext))
        if images:
            # Ordenar y retornar la primera
            images.sort()
            print(f"  ✅ Imagen encontrada: {images[0]}")
            return images[0]

    print(f"  ❌ No se encontraron imágenes en: {base_path}")
    return None


def _get_folder_name(sample_index):
    """
    Obtiene el nombre de la carpeta según el índice de la muestra
    """
    folder_map = {
        0: "AFLUENTE",
        1: "EFLUENTE",
        2: "PUNTO DE DESCARGA"
    }
    return folder_map.get(sample_index)