import json
from docx import Document


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
            replacements += self._search_in_table(table, label, to_write)

        # 3. Search in the header tables
        for section in self.word_template.sections:
            replaced = self._search_in_header(section.header, label, to_write)
            if replaced > 0:
                print(f"label reemplazado en header -> {label}")
                replacements += replaced

        return replacements

    def _search_in_header(self, header, search_text, replace_text):
        replacements = 0

        for paragraph in header.paragraphs:
            if search_text in paragraph.text:
                replacements += self._replace_in_paragraph(paragraph, search_text, replace_text)

        for table in header.tables:
            replacements += self._search_in_table(table, search_text, replace_text)

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

    def main_writer(self):
        """
        Recorre todas las claves del JSON y las reemplaza en el documento
        """
        for key, value in self.json_config_fields.items():
            self.search_and_replace(key, value)

    # SAVE DOCUMENT WITH CHANGES
    def save_document(self, output_path):
        self.word_template.save(output_path)