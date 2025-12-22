import json
import logging
import os
import traceback
import time
import gc

import win32com.client as win32
import pythoncom  # ← NUEVO: Import pythoncom
from docx import Document
from openpyxl.reader.excel import load_workbook

from core.exceptions import NotJsonConfigFound
from graphics.Services.copy.copy_graphics import CopyGraphics
from graphics.Services.read.read_graphics_data import ReadGraphicsData
from graphics.Services.write.write_insitu_data import WriteInsituData
from graphics.Services.write.write_parameters_results import WriteParametersResults

logger = logging.getLogger(__name__)

class GraphicGenerator:

    def __init__(self):

        self.graphics_config = None
        self.chain_of_custody = None

        # Constants
        self.JSON_FOLDER_URL = "fields_config/"
        self.word_path = "templates/informe_final.docx"

        # Excels
        self.graphics_obj = None
        self.chain_custody_obj = None
        self.json_config = None
        self.word_obj = None

        # COM objects
        self.graphics_excel_com = None
        self.word_com_app = None  # ← NUEVO: Guardar referencia a Word Application
        self.excel_com_app = None  # ← NUEVO: Guardar referencia a Excel Application

        # Instances
        self.reader_graphics_data = ReadGraphicsData()
        self.writer_graphics_data = WriteInsituData()
        self.writer_parameters_data = WriteParametersResults()
        self.copy_graphics = None

    def search_graphics_config(self, report_name: str):

        report_name = "GRAFICAS_" + report_name + ".json"

        if report_name not in os.listdir(self.JSON_FOLDER_URL):
            raise NotJsonConfigFound(f"The configuration file for the {report_name} graphics does not exist")

        config_path = os.path.join(self.JSON_FOLDER_URL, report_name)

        try:
            with open(config_path, 'r', encoding='utf-8') as file:
                self.graphics_config = json.load(file)

            logger.info(f"Graphics config loaded successfully {report_name}")

        except json.JSONDecodeError as e:
            logger.error(f"Invalid json in {report_name}: {e}")
            raise

        except Exception as e:
            logger.error(f"Error reading {report_name} graphics config: {e}")
            raise

    def load_graphics_excel(self, url_graphics_config, read_only):

        if not url_graphics_config:
            logger.error(f"No excel url for the charts is provided")

        graphics_excel = load_workbook(url_graphics_config, data_only=read_only)
        self.graphics_obj = graphics_excel
        logger.info("Graphics excel open successfully")

    def load_graphics_excel_com(self, excel_path):
        """
        Carga Excel con COM
        ★★★ DEBE ejecutarse después de pythoncom.CoInitialize() ★★★
        """
        try:
            logger.info("Creando Excel COM Application...")
            excel = win32.gencache.EnsureDispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            
            logger.info(f"Abriendo archivo Excel: {excel_path}")
            wb = excel.Workbooks.Open(os.path.abspath(excel_path))
            
            self.graphics_excel_com = wb
            self.excel_com_app = excel  # ← NUEVO: Guardar referencia a Application
            
            logger.info("Excel COM loaded successfully for graphics")
            
        except Exception as e:
            logger.error(f"Error loading Excel COM: {e}", exc_info=True)
            raise

    def load_json_config(self, report_name: str):

        path_config = f"{self.JSON_FOLDER_URL}GRAFICAS_{report_name}.json"

        try:
            with open(path_config, 'r', encoding='utf-8') as archivo:
                self.json_config = json.load(archivo)

        except Exception as ex:
            logger.error(f"Error opening json {ex}")

        return

    def load_chain_of_custody(self, chain_custody):

        if not chain_custody:
            logger.error("Not chain of custody excel")
            return

        try:
            self.chain_of_custody = chain_custody

        except Exception as ex:
            logger.error(f"Error opening the chain of custody excel")

    def load_word_obj(self):
        """
        Carga Word con python-docx (para escritura normal de texto)
        NO usar esta función si vas a pegar gráficas COM
        """
        if not self.word_path:
            logger.error("No se suministro ruta del word")
            raise ValueError("Word path is None")
            
        try:
            self.word_obj = Document(self.word_path)

        except Exception as ex:
            logger.error(f"Error cargando documento {ex}", exc_info=True)
            raise

    def load_word_obj_com(self):
        """
        Abre Word con COM para poder pegar gráficos
        ★★★ DEBE ejecutarse después de pythoncom.CoInitialize() ★★★
        """
        if not self.word_path:
            logger.error("No se suministro ruta del word")
            raise ValueError("Word path is None")

        try:
            logger.info("Creando Word COM Application...")
            word = win32.gencache.EnsureDispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            word.ScreenUpdating = False
            
            logger.info(f"Abriendo documento Word: {self.word_path}")
            doc = word.Documents.Open(os.path.abspath(self.word_path))
            
            self.word_obj = doc
            self.word_com_app = word  # ← NUEVO: Guardar referencia a Application
            
            logger.info("Word COM loaded successfully")
            
        except Exception as ex:
            logger.error(f"Error cargando documento Word con COM: {ex}", exc_info=True)
            raise

    def copy_graphics_proc(self):
        """
        Ejecuta el proceso de copia de gráficas
        ★★★ DEBE ejecutarse en el MISMO thread que creó los objetos COM ★★★
        """
        try:
            if self.copy_graphics is None:
                self.copy_graphics = CopyGraphics(
                    self.graphics_excel_com,
                    self.json_config,
                    self.word_obj,
                    self.word_path
                )

            logger.info("Iniciando copy_and_insert_editable_charts...")
            self.copy_graphics.copy_and_insert_editable_charts()
            logger.info("copy_and_insert_editable_charts completado")
            
            # paste_graphics() parece estar vacío en tu código, pero lo mantengo
            self.copy_graphics.paste_graphics()

        except Exception as ex:
            logger.error(f"Error copiando graficas", exc_info=True)
            raise

    def cleanup_com_objects(self):
        """
        ═══════════════════════════════════════════════════════════════
        NUEVO: Limpieza exhaustiva de objetos COM en el orden correcto
        ═══════════════════════════════════════════════════════════════
        """
        logger.info("Iniciando limpieza de objetos COM...")
        
        # 1. Cerrar documentos
        try:
            if self.word_obj and hasattr(self.word_obj, 'Close'):
                logger.info("Cerrando documento Word...")
                self.word_obj.Close(SaveChanges=True)
                self.word_obj = None
        except Exception as e:
            logger.warning(f"Error cerrando Word doc: {e}")
        
        try:
            if self.graphics_excel_com and hasattr(self.graphics_excel_com, 'Close'):
                logger.info("Cerrando workbook Excel...")
                self.graphics_excel_com.Close(SaveChanges=False)
                self.graphics_excel_com = None
        except Exception as e:
            logger.warning(f"Error cerrando Excel workbook: {e}")
        
        # 2. Cerrar aplicaciones
        try:
            if self.word_com_app:
                logger.info("Cerrando Word Application...")
                self.word_com_app.Quit()
                self.word_com_app = None
        except Exception as e:
            logger.warning(f"Error cerrando Word app: {e}")
        
        try:
            if self.excel_com_app:
                logger.info("Cerrando Excel Application...")
                self.excel_com_app.Quit()
                self.excel_com_app = None
        except Exception as e:
            logger.warning(f"Error cerrando Excel app: {e}")
        
        # 3. Pump messages y garbage collection
        try:
            pythoncom.PumpWaitingMessages()
            gc.collect()
            logger.info("Recursos COM liberados")
        except Exception as e:
            logger.warning(f"Error en limpieza final: {e}")

    def load_objs(self, report_name, chain_custody, read_only: False):
        try:
            self.load_json_config(report_name)
            self.load_graphics_excel(self.json_config["RUTA_EXCEL_GRAFICAS"], read_only)

            if self.graphics_obj and not read_only:
                self.graphics_obj.close()
                logger.info("Closed openpyxl workbook before COM access")

            self.load_graphics_excel_com(self.json_config["RUTA_EXCEL_GRAFICAS"])
            self.load_chain_of_custody(chain_custody)
            self.load_word_obj()

        except Exception as ex:
            logger.error(f"Error in the load_objs {ex}")
            logger.error(f"Traceback: {traceback.format_exc()}")

    def main(self, report_name: str, chain_custody):
        """
        Método principal que ejecuta todo el proceso
        ★★★ DEBE ejecutarse en el MAIN THREAD ★★★
        """
        
        # ═══════════════════════════════════════════════════════════════
        # NUEVO: Inicializar COM al inicio del proceso
        # ═══════════════════════════════════════════════════════════════
        logger.info("═══════════════════════════════════════════════════════")
        logger.info("INICIANDO PROCESO - Inicializando COM...")
        logger.info("═══════════════════════════════════════════════════════")
        pythoncom.CoInitialize()
        
        try:
            # ═══════════════════════════════════════════════════════════
            # FASE 1: Cargar datos y escribir en Excel con openpyxl
            # ═══════════════════════════════════════════════════════════
            logger.info("FASE 1: Procesamiento de datos con openpyxl...")
            
            self.load_json_config(report_name)
            self.load_graphics_excel(self.json_config["RUTA_EXCEL_GRAFICAS"], False)
            self.load_chain_of_custody(chain_custody)

            self.reader_graphics_data.get_chain_custody(chain_custody)
            self.reader_graphics_data.load_json_config(self.json_config)
            self.reader_graphics_data.get_cod_muestras()
            
            sampler_parameters, parameters_results = self.reader_graphics_data.read_parameters_to_sampler(
                self.graphics_obj)

            tables_data = self.reader_graphics_data.read_table_ph_solidos_caudales()
            formatted_tables_data = self.reader_graphics_data.complete_table_data()
            
            self.writer_graphics_data.load_insitu_data(formatted_tables_data)
            self.writer_graphics_data.load_json_config(self.json_config)
            self.writer_graphics_data.graphics_excel = self.graphics_obj
            self.writer_graphics_data.write_insitu_data()
            
            self.writer_parameters_data.write_parameters_result(
                self.graphics_obj, self.json_config, parameters_results)

            self.graphics_obj.save(self.json_config["RUTA_EXCEL_GRAFICAS"])
            self.graphics_obj.close()
            self.graphics_obj = None
            logger.info("✓ FASE 1 completada: Excel guardado y cerrado")

            # ═══════════════════════════════════════════════════════════
            # FASE 2: Cálculos post-procesamiento
            # ═══════════════════════════════════════════════════════════
            logger.info("FASE 2: Cálculos post-procesamiento...")
            
            self.load_graphics_excel(self.json_config["RUTA_EXCEL_GRAFICAS"], True)
            final_results = self.reader_graphics_data.graphics_post_calculus(
                self.graphics_obj, self.json_config)
            self.graphics_obj.close()
            self.graphics_obj = None
            logger.info("✓ FASE 2 completada: Cálculos finalizados")

            # ═══════════════════════════════════════════════════════════
            # FASE 3: Escribir resultados finales
            # ═══════════════════════════════════════════════════════════
            logger.info("FASE 3: Escribiendo resultados finales...")
            
            self.load_graphics_excel(self.json_config["RUTA_EXCEL_GRAFICAS"], False)
            self.writer_parameters_data.write_final_results(
                final_results, self.graphics_obj, self.json_config)

            self.graphics_obj.save(self.json_config["RUTA_EXCEL_GRAFICAS"])
            self.graphics_obj.close()
            self.graphics_obj = None
            logger.info("✓ FASE 3 completada: Resultados finales guardados")

            # ═══════════════════════════════════════════════════════════
            # PAUSA: Asegurar que Excel termine de escribir
            # ═══════════════════════════════════════════════════════════
            logger.info("Esperando 2 segundos para asegurar escritura de Excel...")
            time.sleep(2)
            pythoncom.PumpWaitingMessages()

            # ═══════════════════════════════════════════════════════════
            # FASE 4: Copiar gráficas con COM (CRÍTICO - mismo thread)
            # ═══════════════════════════════════════════════════════════
            logger.info("═══════════════════════════════════════════════════════")
            logger.info("FASE 4: Copiando gráficas con COM...")
            logger.info("═══════════════════════════════════════════════════════")
            
            # Abrir Word con COM
            self.load_word_obj_com()
            pythoncom.PumpWaitingMessages()
            
            # Abrir Excel con COM
            self.load_graphics_excel_com(self.json_config["RUTA_EXCEL_GRAFICAS"])
            pythoncom.PumpWaitingMessages()
            
            # Ejecutar copia de gráficas EN EL MISMO THREAD
            self.copy_graphics_proc()
            
            logger.info("✓ FASE 4 completada: Gráficas copiadas exitosamente")

            # ═══════════════════════════════════════════════════════════
            # FASE 5: Guardar y cerrar documentos COM
            # ═══════════════════════════════════════════════════════════
            logger.info("FASE 5: Guardando y cerrando documentos...")
            
            # Pump messages antes de guardar
            pythoncom.PumpWaitingMessages()
            
            # Guardar Word
            if self.word_obj:
                logger.info("Guardando documento Word...")
                self.word_obj.Save()
                logger.info("✓ Word guardado exitosamente")
            
            # Limpieza de objetos COM
            self.cleanup_com_objects()
            
            logger.info("═══════════════════════════════════════════════════════")
            logger.info("✓✓✓ PROCESO COMPLETADO EXITOSAMENTE ✓✓✓")
            logger.info("═══════════════════════════════════════════════════════")

        except Exception as ex:
            logger.error("═══════════════════════════════════════════════════════")
            logger.error(f"✗✗✗ ERROR EN EL PROCESO: {ex}")
            logger.error(f"Traceback: {traceback.format_exc()}")
            logger.error("═══════════════════════════════════════════════════════")

            # Intentar limpieza en caso de error
            try:
                if self.graphics_obj:
                    self.graphics_obj.close()
                    self.graphics_obj = None
            except:
                pass
            
            # Limpieza COM
            try:
                self.cleanup_com_objects()
            except:
                pass
            
            raise

        finally:
            # ═══════════════════════════════════════════════════════════
            # LIMPIEZA FINAL: Uninitialize COM
            # ═══════════════════════════════════════════════════════════
            try:
                pythoncom.PumpWaitingMessages()
                pythoncom.CoUninitialize()
                logger.info("COM uninitializado correctamente")
            except Exception as e:
                logger.warning(f"Error en CoUninitialize: {e}")