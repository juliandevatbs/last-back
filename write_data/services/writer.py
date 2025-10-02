import copy
import json
from docx import Document
from docx.table import Table

from core.exceptions import NoJson


class Writer():

    def __init__(self):
        self.json_config = None
        self.json_config_fields = None
        self.word_template = None


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
                print(f"label reemplazado -> {label}")
                replacements += self._search_in_header(section.header, label, to_write)

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


    def samples_table(self):
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
        return True

    def main_writer(self):

        for key, value in self.json_config_fields.items():

            self.search_and_replace(key, value)



    # SAVE DOCUMENT WITH CHANGES
    def save_document(self, output_path):
        self.word_template.save(output_path)






