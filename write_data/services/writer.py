import copy
import json
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.table import Table

from core.exceptions import NoJson
from read_data.services.excel_reader import data_constructor


class Writer():

    def __init__(self):
        self.json_config = None
        self.json_config_fields = None
        self.word_template = None
        self.main_data =None
        self.font = "Century Gothic"


    def load_exce_data(self, wk):

        self.main_data = data_constructor(wk)





    # GET THE JSON CONFIG
    def load_json_config(self):
        with open('fields_config/fields.json', 'r', encoding='utf-8') as json_file:
            self.json_config = json.load(json_file)

        self.json_config_fields = self.json_config["fields"]

    # GET THE WORD TEMPLATE TO WRITE
    def load_word_template(self):
        self.word_template = Document("templates/PLANTILLA INF_CPF_CUPIAGUA_ACEITOSAS_ARI_ACBB.docx")

    # SEARCH A LABEL AND REPLACE IT
    def search_and_replace(self, label: str, to_write: str) -> int:

        replacements = 0

        # 1. Search in normal paragraphs
        for paragraph in self.word_template.paragraphs:
            if label in paragraph.text:
                print(f"label reemplazado -> {label}")
                replacements += self._replace_in_paragraph(paragraph, label, to_write)

        # 2. Search in all file tables
        for table in self.word_template.tables:
                print(f"label reemplazado -> {label}")
                replacements += self._search_in_table(table, label, to_write)

        # 3. Search in the header tables
        for section in self.word_template.sections:
                replaced = self._search_in_header(section.header, label, to_write)

                if replaced > 0:


                    print(f"label reemplazado -> {label}")
                    replacements += replaced

        return replacements

    def _search_in_header(self, header, search_text, replace_text):

        replacements = 0

        for paragraph in header.paragraphs:
            if search_text in paragraph.text:
                replacements += self._replace_in_paragraph(paragraph, search_text, replace_text)

        for table in header.tables:
            replacements += self._search_in_table(table, search_text, replace_text)  # ✅ Usa _search_in_table

        return replacements

    # REPLACE A LABEL INTO A PARAGRAPH
    def _replace_in_paragraph(self, parrafo, texto_buscar, texto_reemplazo):

        replacements = 0

        for run in parrafo.runs:
            if texto_buscar in run.text:
                times = run.text.count(texto_buscar)
                if type(texto_reemplazo) != str:
                    texto_reemplazo = str(texto_reemplazo)
                run.text = run.text.replace(texto_buscar, texto_reemplazo)
                replacements += times

        return replacements

    # SEARCH ANY LABEL INTO TABLES
    def _search_in_table(self, table, texto_buscar, texto_reemplazo):

        replacements = 0

        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if texto_buscar in paragraph.text:
                        replacements += self._replace_in_paragraph(paragraph, texto_buscar, texto_reemplazo)

                if cell.tables:
                    for table_child in cell.tables:
                        replacements += self._search_in_table(table_child, texto_buscar, texto_reemplazo)

        return replacements

    def write_cell_safe(self, cell, text, size=8, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=0):
        """
        Versión SEGURA de write_cell que NO borra contenido existente
        Solo limpia si la celda está realmente vacía
        """

        # Verificar si la celda tiene contenido importante
        current_text = cell.text.strip()



        # Solo limpiar si es seguro
        self.clear_cell_safe(cell)

        # Agregar el nuevo contenido
        p = cell.add_paragraph()

        if isinstance(text, list):
            for text_part, is_bold in text:
                run = p.add_run(text_part)
                run.font.name = self.font
                run.font.size = Pt(size)
                run.bold = is_bold
        else:
            run = p.add_run(text)
            run.font.name = self.font
            run.font.size = Pt(size)
            run.bold = bold

        p.alignment = align
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1

        return True

    def clear_cell_safe(self, cell):
        """
        Versión SEGURA de clear_cell que preserva imágenes Y títulos importantes
        """

        # Guardar imágenes
        images = []
        important_content = []

        for paragraph in cell.paragraphs:
            para_text = paragraph.text.strip()

            # Preservar contenido importante (títulos, headers, etc.)

            # Guardar imágenes
            for run in paragraph.runs:
                if run._element.xpath('.//w:drawing'):
                    images.extend(run._element.xpath('.//w:drawing'))

        # Solo eliminar párrafos que NO son importantes
        paragraphs_to_remove = []
        for p in cell.paragraphs:
            if p not in important_content:
                paragraphs_to_remove.append(p)

        # Eliminar solo párrafos no importantes
        for p in paragraphs_to_remove:
            p._element.getparent().remove(p._element)

        # Restaurar imágenes si había
        if images:
            if not cell.paragraphs:  # Si no quedan párrafos
                p = cell.add_paragraph()
            else:
                p = cell.paragraphs[0]

            for img in images:
                p._element.append(img)

    def fill_monitoring_table(self, data):
        """
        Versión con DEBUG para ver estructura de tabla
        """
        table_found = None
        target_title = "Puntos de monitoreo"
        self.samples = data.get("samples")

        for element in self.word_template.element.body:
            if element.tag.endswith('p'):
                para_text = ""
                for run in element.findall('.//w:t',
                                           {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                    if run.text:
                        para_text += run.text

                if target_title.lower() in para_text.lower():
                    next_element = element.getnext()
                    while next_element is not None:
                        if next_element.tag.endswith('tbl'):
                            for i, table in enumerate(self.word_template.tables):
                                if table._element == next_element:
                                    table_found = table
                                    break
                            break
                        next_element = next_element.getnext()
                    break

        if table_found is None:
            print("❌ No se encontró la tabla")
            return False

        print(f"✅ Tabla encontrada con {len(table_found.rows)} filas")

        # 🔍 DEBUG: Ver estructura de las primeras filas
        print("\n=== ESTRUCTURA DE LA TABLA ===")
        for row_idx in range(min(3, len(table_found.rows))):
            row = table_found.rows[row_idx]
            print(f"\nFila {row_idx + 1} tiene {len(row.cells)} columnas:")
            for col_idx, cell in enumerate(row.cells):
                cell_text = cell.text.strip()
                print(f"  Columna {col_idx}: '{cell_text[:50]}'")  # Primeros 50 caracteres
        print("=" * 50 + "\n")

        sample_list = list(self.samples.items())

        first_data_row = 3  # Ajusta si es diferente
        rows_per_sample = 2  # 1 fila de datos + 1 fila vacía

        data_rows = []
        for i in range(len(sample_list)):
            data_row_index = first_data_row + (i * rows_per_sample)
            if data_row_index < len(table_found.rows):
                data_rows.append(data_row_index)

        print(f"Filas calculadas para datos (patrón cada {rows_per_sample} filas): {data_rows}")

        # Escribir datos
        for i, (sample_id, sample_data) in enumerate(sample_list):
            if i >= len(data_rows):
                print(f"❌ No hay suficientes filas de datos para todas las muestras")
                break

            target_row = data_rows[i]
            print(f"\nMuestra {i + 1}: Escribiendo en fila {target_row + 1}")

            if target_row >= len(table_found.rows):
                print(f"❌ Fila {target_row + 1} no existe")
                break

            row = table_found.rows[target_row]

            print("=== DATOS DE LA MUESTRA ===")
            print(f"  chemilab_code: {sample_data.get('chemilab_code', 'NO EXISTE')}")
            print(f"  sample_date: {sample_data.get('sample_date', 'NO EXISTE')}")
            print(f"  sample_hour: {sample_data.get('sample_hour', 'NO EXISTE')}")
            print(f"  sample_identification: {sample_data.get('sample_identification', 'NO EXISTE')}")

            # 🔍 Mapeo de columnas - AJUSTA ESTOS ÍNDICES según tu tabla
            column_mapping = {
                'chemilab_code': 0,  # Primera columna
                'sample_date': 1,  # Segunda columna
                'sample_hour': 2,  # Tercera columna (¿o va en otra?)
                'sample_identification': 3  # Cuarta columna
            }

            # Escribir cada dato en su columna correspondiente
            for field_name, col_idx in column_mapping.items():
                data_value = sample_data.get(field_name, '')

                if col_idx < len(row.cells):
                    try:
                        cell = row.cells[col_idx]
                        self.write_cell_safe(
                            cell,
                            str(data_value).upper() if data_value else "",
                            size=10,
                            bold=False,
                            align=WD_ALIGN_PARAGRAPH.CENTER
                        )
                        print(f"  ✓ Columna {col_idx} ({field_name}): '{data_value}'")
                    except Exception as e:
                        print(f"  ❌ Error en columna {col_idx}: {e}")
                else:
                    print(f"  ⚠️ Columna {col_idx} no existe (la fila solo tiene {len(row.cells)} columnas)")

            print(f"  ✅ Muestra {i + 1} procesada")

        print("🎉 PROCESO COMPLETADO")
        return True


    """def samples_table(self):
        target_title = "Descripción del punto de monitoreo"
        table_found = None

        for element in self.word_template.element.body:
            if element.tag.endswith('p'):
                para_text = ""
                for run in element.findall('.//w:t',
                                           {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                    if run.text:
                        para_text += run.text

                if target_title.lower() in para_text.lower():
                    next_element = element.getnext()
                    while next_element is not None:
                        if next_element.tag.endswith('tbl'):
                            for i, table in enumerate(self.word_template.tables):
                                if table._element == next_element:
                                    table_found = table
                                    break
                            break
                        next_element = next_element.getnext()
                    break

        if table_found is None:
            print(f"No se encontró tabla después del título '{target_title}'")
            return False

        original_rows = len(table_found.rows)
        template_block_size = 4

        print(f"Tabla encontrada con {original_rows} filas")

        if original_rows < template_block_size:
            print(f"Error: La tabla original tiene menos de {template_block_size} filas")
            return False

        # Guardar template del primer bloque
        template_rows = []
        for i in range(template_block_size):
            if i < len(table_found.rows):
                row_element = table_found.rows[i]._element
                template_rows.append(copy.deepcopy(row_element))

        # Calcular bloques adicionales
        samples_count = len(self.json_config["fields"]["CANTIDAD_MUESTRAS"])
        additional_blocks_needed = samples_count - 1

        print(f"Necesitamos {additional_blocks_needed} bloques adicionales para {samples_count} muestras")

        # Agregar bloques adicionales
        for block_num in range(additional_blocks_needed):
            print(f"Creando bloque adicional {block_num + 2} de {samples_count}")

            for template_row in template_rows:
                new_row_element = copy.deepcopy(template_row)

                # LIMPIAR SOLO DATOS, NO TÍTULOS
                for cell in new_row_element.findall('.//w:tc',
                                                    {
                                                        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                    for para in cell.findall('.//w:p',
                                             {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                        para_text = ""
                        for text_elem in para.findall('.//w:t',
                                                      {
                                                          'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                            if text_elem.text:
                                para_text += text_elem.text

                        # SOLO limpiar si NO es un título importante
                        if (para_text and
                                not any(keyword in para_text.upper() for keyword in
                                        ['PLAN DE MUESTREO', 'DESCRIPCIÓN DEL PUNTO', 'CONDICIONES AMBIENTALES',
                                         'CÓDIGO', 'FECHA', 'HORA', 'IDENTIFICACIÓN', 'COORDENADAS', 'FOTOGRAFÍA',
                                         'ESTE', 'NORTE', 'MUESTRA']) and
                                len(para_text.strip()) < 10):  # Solo limpiar textos cortos

                            for text_elem in para.findall('.//w:t',
                                                          {
                                                              'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                                text_elem.text = ""

                # Agregar la fila a la tabla
                table_found._element.append(new_row_element)

        print(f"Bloques duplicados exitosamente. Tabla ahora tiene {len(table_found.rows)} filas")
        return True"""

    def main_writer(self):

        for key, value in self.json_config_fields.items():

            self.search_and_replace(key, value)



    # SAVE DOCUMENT WITH CHANGES
    def save_document(self, output_path):
        self.word_template.save(output_path)






