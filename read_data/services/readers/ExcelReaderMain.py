from read_data.services.readers.ph_reader import ph_reader
from read_data.services.readers.read_chain_custody import read_chain_custody
from read_data.services.readers.read_main_sheet import read_main_sheet


class ExcelReaderMain:



    def __init__(self):


        #SHEET NAMES
        self.BASIC_SHEET_NAME = "DATOS BASICOS"
        self.CHAIN_OF_CUSTODY_SHEET_NAME = "CADENA DE CUSTODIA"

        self.AFLUENTE_SHEET_NAME = "AFLUENTE 1"
        self.EFLUENTE_SHEET_NAME = "EFLUENTE 2"
        self.PUNTO_DESCARGUE_SHEET_NAME = "CADENA DE VIGILANCIA PUNTUAL"


        self.ph_columns_afluente = {

            "initial_row": "B",
            "hour_column": "C",
            "ph_column": "D",
            "caudal_column": "T",
            "solidos_sedimentables_column": "L",
        }


        self.workbook = None

        #COLUMNS TO READ FROM CHAIN OF CUSTODY
        self.chain_of_custody_columns = ("A", "C", "G", "H", "I", "G")



        #Storage the readed data
        self.basic_data = None
        self.samples_data = None


    def load_work_book(self, workbook):

        if workbook:

            self.workbook = workbook


    def caller(self):


        try:

            self.basic_data = read_main_sheet(self.workbook, self.BASIC_SHEET_NAME)

            self.samples_data = read_chain_custody(self.workbook, self.CHAIN_OF_CUSTODY_SHEET_NAME, r"C:\Code\automatizacion_informes\BackEnd\templates\2 PM 20164 (2025-04-11)ACEITOSAS CPF CUPIAGUA_Muestreos Barranca.xlsx")

            self.ph_data = ph_reader(self.workbook, self.AFLUENTE_SHEET_NAME, self.ph_columns_afluente, 75)

            self.ph_data_2 = ph_reader(self.workbook, self.EFLUENTE_SHEET_NAME, self.ph_columns_afluente, 75)



            return self.basic_data, self.samples_data, self.ph_data, self.ph_data_2

        except Exception as ex:

            print("Error reading the sheets")