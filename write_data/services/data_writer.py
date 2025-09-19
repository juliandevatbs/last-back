from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.shared import Pt
from docx.text.paragraph import Paragraph

# Import custom exceptions
from core.exceptions import NoDataError
from core.services.server_service import ServerService
from core.utils.data.int_to_string_relative import int_to_string_relative
from write_data.services.text_fixing import sampling_site_fixing
from datetime import datetime


# Service for write template
class WordService:


    def __init__(self, template_path, data_to_write: dict):

        self.template_path = template_path
        self.doc = None
        self.data_to_write = data_to_write

        # Server instance
        self.server_templates = ServerService()

        #IMPORTANT CONSTRAINTS
        self.font = "Century Gothic"



        # Get data from the dict
        self.main_data = self.data_to_write["main_data"]
        self.sampling_data = self.data_to_write["sampling_data"]
        self.samples = self.data_to_write["samples"]
        self.surveillance_data = self.data_to_write["surveillance_data"]


        # General data for pages
        self.sampling_site = self.sampling_data["sampling_site"]
        self.prefix_water_title = "CARACTERIZACIÓN FISICOQUÍMICA DE"
        self.water_type = self.sampling_data["sampling_site"]
        self.report_name = f"{self.prefix_water_title} {self.water_type}"
        self.fixed_sampling_site = sampling_site_fixing(self.sampling_site)
        self.client_contact_name = self.main_data["contact_client_name"]
        self.client_name = self.main_data["client_name"]
        self.report_number = self.main_data["report_number"]
        self.sampling_date = self.sampling_data["sampling_date"]
        self.sampling_date_fixed = self.sampling_date.strftime("%Y-%m-%d")
        self.samples_quantity = len(self.samples)




    def clear_cell(self, cell):

        for p in cell.paragraphs:

            p._element.getparent().remove(p._element)


    def load_template(self):

        try:

            # Open the docx
            self.doc = Document(self.template_path)
            return True


        except Exception as ex:

            # If error opening the file
            print(f"Error while charging the document {ex}")
            return None

    #Function to write the header data
    def write_header_table(self):

        if not self.doc:

            return False

        # Get the header
        header = self.doc.sections[0].header

        # Get the header table
        table_header = header.tables[0]

        report_type = self.main_data["report_type"]

        first_cell = table_header.cell(1, 1)

        self.clear_cell(first_cell)

        # clean previous text
        first_cell.text = ""


        # Lines to write in the first section
        first_lines = [
            report_type,
            self.report_name,
            self.fixed_sampling_site
        ]

        for line in first_lines:
            p = first_cell.add_paragraph()
            run = p.add_run(line)
            run.font.name = self.font
            run.font.size = Pt(8)
            run.bold = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Line to write in 4 section
        elaborate_day = datetime.today().strftime("%Y-%m-%d")

        fourth_cell_lines = [
            "ELABORADO",
            elaborate_day
        ]

        fourth_cell = table_header.cell(2, 3)

        self.clear_cell(fourth_cell)

        for line in fourth_cell_lines:

            p = fourth_cell.add_paragraph()
            run = p.add_run(line)
            run.font.name = self.font
            run.font.size = Pt(8)
            run.bold = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Write data into the first page table
    def write_info_table_f_page(self):

        if not self.doc:

            return False


        first_table = self.doc.tables[0]
        first_cell = first_table.cell(0, 0)

        # clean previous text
        first_cell.text = ""

        first_cell_lines = [
            "Presentado a:",
            self.client_contact_name,
            self.client_name

        ]

        for index, line in enumerate(first_cell_lines):
            p = first_cell.add_paragraph()
            run = p.add_run(line)
            run.font.name = self.font
            run.font.size = Pt(10)
            run.bold = True if index == 0 else False
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        second_cell = first_table.cell(0, 1)

        second_cell.text = ""

        second_cell_lines = [

            "Revisado por:"

        ]

        for index, line in enumerate(second_cell_lines):
            p = second_cell.add_paragraph()
            run = p.add_run(line)
            run.font.name = self.font
            run.font.size = Pt(10)
            run.bold = True if index == 0 else False
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        third_cell = first_table.cell(0, 2)
        third_cell.text = ""

        third_cell_lines = [

            "Autorizado por:",
            "Claudia Calderón",
            "Directora de Proyectos"

        ]

        for index, line in enumerate(third_cell_lines):
            p = third_cell.add_paragraph()
            run = p.add_run(line)
            run.font.name = self.font
            run.font.size = Pt(10)
            run.bold = True if index == 0 else False
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER




    def first_page(self):

        if not self.doc:

            return False

        len_data = len(self.data_to_write)

        if len_data == 0 :

            # Return a custom exception to the parent function
            raise NoDataError


        report_type =  self.main_data["report_type"]
        water_type = self.surveillance_data["water_type"]


        # First paragraph (This is the type of report like INFORME DE MONITOREO)
        p = self.doc.paragraphs[0].insert_paragraph_before(report_type)
        p.paragraph_format.space_before = Pt(20)

        # First paragraph styles
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run1 = p.runs[0]
        run1.bold = True
        run1.font.name = self.font
        run1.font.size = Pt(20)


        """
        
            Second paragraph (This is the second title with the location and water location)
        
            Possible types:
            
                Agua subterranea
                Agua superficial
                
        
        """


        water_type = self.sampling_data["sampling_site"]
        sampling_site = self.sampling_data["sampling_site"]



        p_two = self.doc.paragraphs[1].insert_paragraph_before(f"{self.prefix_water_title} {water_type}")
        p_two.paragraph_format.space_before = Pt(80)
        p_two.paragraph_format.space_after = Pt(80)
        p_two.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run2 = p_two.runs[0]
        run2.bold = True
        run2.font.name = self.font
        run2.font.size = Pt(20)

        # Location
        p_location = self.doc.paragraphs[2].insert_paragraph_before(self.fixed_sampling_site)
        p_location.paragraph_format.space_before = Pt(80)
        p_location.paragraph_format.space_after = Pt(10)
        p_location.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run3 = p_location.runs[0]
        run3.bold = True
        run3.font.name = self.font
        run3.font.size = Pt(20)



        return True


    def second_part_page(self):

        if not self.doc:

            return False

        # First paragraph (This is the type of report like INFORME DE MONITOREO)
        first_table = self.doc.tables[0]

        tbl_element = first_table._element
        new_paragraph = OxmlElement("w:p")
        tbl_element.addnext(new_paragraph)

        p1 = Paragraph(new_paragraph, first_table._parent)
        p1.paragraph_format.space_before = Pt(20)
        p1.paragraph_format.space_after = Pt(20)
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER

        run1_bold = p1.add_run("INFORME NÚMERO: ")
        run1_bold.font.name = self.font
        run1_bold.font.size = Pt(10)
        run1_bold.bold = True

        run1_normal = p1.add_run(f"{self.report_number}")
        run1_normal.font.name = self.font
        run1_normal.font.size = Pt(10)
        run1_normal.bold = False

        new_paragraph2 = OxmlElement("w:p")
        p1._element.addnext(new_paragraph2)

        p2 = Paragraph(new_paragraph2, first_table._parent)
        p2.paragraph_format.space_before = Pt(5)
        p2.paragraph_format.space_after = Pt(5)
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER

        run2_bold = p2.add_run("Fechas de muestreo: ")
        run2_bold.font.name = self.font
        run2_bold.font.size = Pt(11)
        run2_bold.bold = True


        run2_normal = p2.add_run(f"{self.sampling_date_fixed}")
        run2_normal.font.name = self.font
        run2_normal.font.size = Pt(10)
        run2_normal.bold = False




        print(self.samples_quantity)


        templates = self.server_templates.get_avaible_templates()

        print(templates)


    def write_goals(self):


        relative_quantity_samples = int_to_string_relative(self.samples_quantity)

        # Number of samples plural o singular
        number_samples_p_s = 'punto'

        if self.samples_quantity > 1:

            number_samples_p_s = "puntos"





        return True













    def save(self, output_path):

        if self.doc:
            self.doc.save(output_path)
            print("Document saved")