import os
from docx import Document
from write_data.services.docx_writer.ph_table_writer import ph_table_writer
from write_data.services.docx_writer.write_first_page import write_first_page
from write_data.services.docx_writer.write_header import write_header
from write_data.services.docx_writer.write_monitoring_tabe import write_monitoring_table


class DocxWriterMain:

    def __init__(self):
        self.docx = None
        self.basic_data = None
        self.samples_data = None
        self.ph_data_afluente = None
        self.ph_data_efluente = None
        self.template_path = r"C:\Code\automatizacion_informes\BackEnd\templates\PLANTILLA INF_CPF_CUPIAGUA_ACEITOSAS_ARI_ACBB.docx"
        self.output_path = r"C:\Code\automatizacion_informes\BackEnd\templates"
        self.save_filename = "final.docx"
        self.font = 'Century Gothic'

        self.font_size_header = 8
        self.font_bold_header = True

        self.font_size_first_page = 10
        self.font_bold_first_page = False

        self.font_size_monitoring_table = 8
        self.font_bold_monitoring_table = False

        self.title_afluente = "Resultados In Situ del AFLUENTE SISTEMA AGUAS LLUVIAS OCASIONALMENTE ACEITOSAS DE CPF CUPIAGUA"
        self.title_efluente = "Resultados In Situ del EFLUENTE SISTEMA AGUAS LLUVIAS OCASIONALMENTE ACEITOSAS DE CPF CUPIAGUA"

    def load_docx(self):
        try:
            self.docx = Document(self.template_path)
        except FileNotFoundError:
            print("File docx not found")
        except Exception as ex:
            print(f"Error opening the docx: {ex}")

    def load_data(self, basic_data: dict, samples_data: dict, ph_data_afluente: dict, ph_data_efluente: dict):
        self.basic_data = basic_data
        self.samples_data = samples_data
        self.ph_data_afluente = ph_data_afluente
        self.ph_data_efluente = ph_data_efluente

    def caller(self):
        write_header(self.docx, self.font, self.font_size_header, self.font_bold_header)

        write_first_page(self.docx, self.font, self.font_size_first_page, self.font_bold_first_page,
                         self.basic_data["sampling_basic_data"]["XX_FECHA_MUESTREO_XX"])

        write_monitoring_table(self.docx, self.font, self.font_size_monitoring_table,
                               self.font_bold_monitoring_table, self.samples_data, self.basic_data)

        print("\n" + "=" * 80)
        print("Writing AFLUENTE table")
        print("=" * 80)
        ph_table_writer(self.docx, self.font, self.font_size_monitoring_table,
                        self.font_bold_monitoring_table, self.ph_data_afluente, self.title_afluente)

        print("\n" + "=" * 80)
        print("Writing EFLUENTE table")
        print("=" * 80)
        ph_table_writer(self.docx, self.font, self.font_size_monitoring_table,
                        self.font_bold_monitoring_table, self.ph_data_efluente, self.title_efluente)



        return True

    def save_doc(self):
        filename = os.path.join(self.output_path, self.save_filename)
        os.makedirs(self.output_path, exist_ok=True)
        self.docx.save(filename)
        print(f"\nDocument saved successfully: {filename}")