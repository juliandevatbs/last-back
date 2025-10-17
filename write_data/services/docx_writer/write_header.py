from datetime import datetime
from docx.shared import Pt, RGBColor



def write_header(doc, font: str, size: int, bold: bool):
    # Access to the header
    section = doc.sections[0]
    header = section.header

    header_tables = header.tables

    if len(header_tables) == 0:
        return False

    table = header_tables[0]
    cell = table.rows[2].cells[3]

    today_date = datetime.now().strftime('%Y-%m-%d')

    replaced = False

    for idx, paragraph in enumerate(cell.paragraphs):

        if 'XX_FECHA_ELABORACION_XX' in paragraph.text:
            old_text = paragraph.text
            paragraph.text = paragraph.text.replace('XX_FECHA_ELABORACION_XX', today_date)

            for run in paragraph.runs:

                run.font.name = font
                run.font.size = Pt(size)
                run.font.bold = bold

            replaced = True



    if not replaced:
        return False

    return True


