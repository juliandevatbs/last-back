from io import BytesIO

from docx import Document


def extract_text_docx(file_obj):


    # extract all text from the report

    file_ob = BytesIO(file_obj)

    doc = Document(file_ob)

    return "\n".join([p.text for p in doc.paragraphs if p.text.strip() != ""])