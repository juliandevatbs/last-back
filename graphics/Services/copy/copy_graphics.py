import logging
import win32com.client as win32
import pythoncom
from pathlib import Path
import time
import gc
import threading
from functools import wraps

from core.exceptions import KeyNotFound

logger = logging.getLogger(__name__)


# ==============================================================================
# DECORATOR FOR COM OPERATION RETRIES
# ==============================================================================
def com_operation_retry(max_attempts=3, delay=1.0):
    """
    Decorator to retry COM operations that may fail due to timing issues
    
    Args:
        max_attempts: Maximum number of retry attempts
        delay: Delay in seconds between attempts
    """
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            last_exception = None
            for attempt in range(max_attempts):
                try:
                    pythoncom.PumpWaitingMessages()
                    result = func(*args, **kwargs)
                    return result
                except Exception as e:
                    last_exception = e
                    if attempt < max_attempts - 1:
                        logger.warning(f"Attempt {attempt + 1} failed: {e}. Retrying...")
                        time.sleep(delay)
                        pythoncom.PumpWaitingMessages()
                    else:
                        logger.error(f"Operation failed after {max_attempts} attempts")
            raise last_exception
        return wrapper
    return decorator


class CopyGraphics:

    def __init__(self, excel_obj, json_config, word_to_write, word_path):
        """
        Initialize the CopyGraphics class
        
        Args:
            excel_obj: Excel workbook COM object
            json_config: Configuration dictionary with paths and settings
            word_to_write: Word document COM object
            word_path: Path to the Word document
        """
        self.excel_graph = excel_obj
        self.json_config = json_config
        self.excel_path = self.json_config["RUTA_EXCEL_GRAFICAS"]
        self.word_obj = word_to_write
        self.word_path = word_path
        
        # Thread verification
        self.main_thread_id = threading.current_thread().ident
        logger.info(f"CopyGraphics initialized in thread: {self.main_thread_id}")


    def copy_and_insert_editable_charts(self):
        """
        Main method to copy and insert editable charts from Excel to Word
        Uses Word's Find API instead of iterating through paragraphs for better performance
        """
        logger.info("=" * 60)
        logger.info("START: copy_and_insert_editable_charts")
        logger.info("=" * 60)
        
        # Validate required objects
        if self.excel_graph is None:
            logger.error("Excel graphics object not provided", exc_info=True)
            raise ValueError("Excel object is None")

        if self.json_config is None:
            logger.error("JSON configuration not provided", exc_info=True)
            raise ValueError("JSON config is None")

        if self.word_obj is None:
            logger.error("Word object not provided", exc_info=True)
            raise ValueError("Word object is None")

        original_screen_updating = None
        word_app = None
        
        try:
            # Thread verification
            current_thread = threading.current_thread().ident
            if current_thread != self.main_thread_id:
                logger.warning("WARNING: Running in different thread!")
                logger.warning(f"   Original thread: {self.main_thread_id}")
                logger.warning(f"   Current thread: {current_thread}")
            else:
                logger.info(f"Running in correct thread: {current_thread}")
            
            # Get Word Application object
            logger.info("Getting Word Application object...")
            try:
                word_app = self.word_obj.Application
                logger.info("Application obtained successfully")
                
            except Exception as app_error:
                logger.error(f"Error obtaining Application: {app_error}")
                raise
            
            # Disable screen updating for better performance
            if word_app is not None:
                try:
                    original_screen_updating = word_app.ScreenUpdating
                    word_app.ScreenUpdating = False
                    logger.info("ScreenUpdating disabled")
                    
                except Exception as screen_error:
                    logger.warning(f"Could not disable ScreenUpdating: {screen_error}")
            
            # Pump COM messages before heavy operations
            pythoncom.PumpWaitingMessages()
            
            # Build chart registry
            logger.info("Building chart registry...")
            chart_registry = self._build_chart_registry(self.excel_graph)
            logger.info(f"Registry built: {sum(len(charts) for charts in chart_registry.values())} total charts")

            pythoncom.PumpWaitingMessages()

            # Insert charts into Word document
            logger.info("Inserting charts into Word...")
            self._insert_charts_in_word(self.word_obj, chart_registry)

            pythoncom.PumpWaitingMessages()

            # Save Word document
            logger.info("Saving Word document...")
            self.word_obj.Save()
            logger.info("Document saved successfully")

            logger.info("=" * 60)
            logger.info("SUCCESS: All charts inserted correctly")
            logger.info("=" * 60)

        except KeyNotFound:
            logger.error("Chart sheets do not exist in the provided Excel file", exc_info=True)
            raise

        except Exception as ex:
            logger.error(f"Error copying charts from Excel: {ex}", exc_info=True)
            raise
            
        finally:
            logger.info("Executing final cleanup...")
            try:
                if word_app is not None and original_screen_updating is not None:
                    word_app.ScreenUpdating = original_screen_updating
                    logger.info("ScreenUpdating restored")
                    
            except Exception as e:
                logger.warning(f"Error restoring ScreenUpdating: {e}")
                
            self._cleanup_resources()
            logger.info("Cleanup completed")


    @com_operation_retry(max_attempts=3, delay=1.0)
    def _build_chart_registry(self, excel_workbook):
        """
        Build a registry of all charts in the specified sheets
        
        Args:
            excel_workbook: Excel workbook COM object
            
        Returns:
            dict: Dictionary mapping sheet names to chart objects
        """
        logger.info("--- Starting registry construction ---")
        chart_registry = {}
        sheets = self.json_config["HOJAS_GRAFICAS"].values()
        
        logger.info(f"Sheets to process: {list(sheets)}")

        for sheet_name in sheets:
            logger.info(f"Processing sheet: {sheet_name}")
            
            try:
                pythoncom.PumpWaitingMessages()
                sheet = excel_workbook.Sheets(sheet_name)
                logger.info(f"  Sheet {sheet_name} found")
                
            except Exception as e:
                logger.error(f"  Sheet {sheet_name} not found: {e}")
                continue

            try:
                chart_objects = sheet.ChartObjects()
                chart_count = chart_objects.Count
                logger.info(f"  Charts in {sheet_name}: {chart_count}")
                
            except Exception as e:
                logger.error(f"  Error accessing ChartObjects in {sheet_name}: {e}")
                continue

            if chart_count == 0:
                logger.warning(f"  No charts found in {sheet_name}")
                continue

            if sheet_name not in chart_registry:
                chart_registry[sheet_name] = {}

            for i in range(1, chart_count + 1):
                try:
                    chart_obj = chart_objects(i)
                    chart_registry[sheet_name][i] = chart_obj
                    logger.info(f"    Registered chart {i}")
                    
                except Exception as chart_error:
                    logger.error(f"    Error getting chart {i} from sheet {sheet_name}: {chart_error}")

        logger.info(f"--- Registry completed: {len(chart_registry)} sheets ---")
        return chart_registry


    def _insert_charts_in_word(self, word_doc, chart_registry):
        """
        Insert charts into the Word document
        
        Args:
            word_doc: Word document COM object
            chart_registry: Dictionary of chart objects from _build_chart_registry
        """
        logger.info("--- Starting chart insertion ---")
        graphic_mappings = self._get_graphic_mappings()
        chart_counter = 0
        success_counter = 0
        fail_counter = 0

        for sheet_name, charts in graphic_mappings.items():
            logger.info(f"Processing sheet: {sheet_name} with {len(charts)} charts")
            
            if sheet_name not in chart_registry:
                logger.warning(f"Sheet {sheet_name} not found in registry")
                continue

            for chart_index, config in charts.items():
                chart_counter += 1
                logger.info(f"[{chart_counter}/28] Processing chart {chart_index} from {sheet_name}")
                
                # Pump COM messages before each chart operation
                pythoncom.PumpWaitingMessages()
                
                try:
                    if chart_index not in chart_registry[sheet_name]:
                        logger.warning(f"Chart {chart_index} not found in sheet {sheet_name}")
                        fail_counter += 1
                        continue

                    chart_obj = chart_registry[sheet_name][chart_index]
                    search_text = config["text"]
                    target_occurrence = config["occurrence"]

                    if self._insert_chart_at_location(
                        word_doc, 
                        chart_obj, 
                        search_text, 
                        target_occurrence, 
                        chart_index,
                        sheet_name
                    ):
                        success_counter += 1
                        logger.info(f"[{chart_counter}/28] SUCCESS - Total successful: {success_counter}")
                        
                    else:
                        fail_counter += 1
                        logger.warning(f"[{chart_counter}/28] FAIL - Location not found")

                except Exception as ex:
                    fail_counter += 1
                    logger.error(f"[{chart_counter}/28] ERROR: {ex}", exc_info=True)
                    self._clear_clipboard()
                    continue
                
                # Pump COM messages after each chart
                pythoncom.PumpWaitingMessages()
                
                # Maintenance pause every 5 charts to give COM a break
                if chart_counter % 5 == 0:
                    logger.info(f"Maintenance pause after {chart_counter} charts...")
                    time.sleep(1.0)
                    pythoncom.PumpWaitingMessages()

        logger.info("=" * 60)
        logger.info("--- Insertion completed ---")
        logger.info(f"    Successful: {success_counter}")
        logger.info(f"    Failed: {fail_counter}")
        logger.info("=" * 60)


    def _insert_chart_at_location(self, word_doc, chart_obj, search_text, target_occurrence, chart_index, sheet_name):
        """
        Insert a chart at a specific location in the Word document
        Uses Word's Find API for efficient searching
        
        Args:
            word_doc: Word document COM object
            chart_obj: Excel chart object to insert
            search_text: Text to search for in the document
            target_occurrence: Which occurrence of the text to use
            chart_index: Index of the chart (for logging)
            sheet_name: Name of the Excel sheet (for logging)
            
        Returns:
            bool: True if successful, False otherwise
        """
        logger.info(f"Searching for: '{search_text}' (occurrence {target_occurrence})")
        
        try:
            # Get Word selection and Find object
            selection = word_doc.Application.Selection
            find_obj = selection.Find
            
            # Configure search
            find_obj.ClearFormatting()
            find_obj.Text = search_text
            find_obj.Forward = True
            find_obj.Wrap = 0  # wdFindStop - don't wrap around document
            find_obj.Format = False
            find_obj.MatchCase = False
            find_obj.MatchWholeWord = False
            
            # Go to document start
            selection.HomeKey(6)  # wdStory
            found_count = 0
            
            # Search for all occurrences until we find the correct one
            while find_obj.Execute():
                found_count += 1
                logger.info(f"  Found occurrence {found_count}")
                
                if found_count == target_occurrence:
                    logger.info(f"  Correct location found, inserting chart...")
                    
                    # Multiple paste attempts with different methods
                    max_paste_attempts = 3
                    paste_success = False
                    
                    for paste_attempt in range(max_paste_attempts):
                        try:
                            logger.info(f"    Paste attempt {paste_attempt + 1}/{max_paste_attempts}")
                            
                            # Step 1: Clear clipboard
                            pythoncom.PumpWaitingMessages()
                            self._clear_clipboard()
                            time.sleep(0.5)

                            # Step 2: Copy chart from Excel
                            logger.info("    Copying chart from Excel...")
                            chart_obj.Copy()
                            
                            # Wait longer for Excel to complete the copy operation
                            time.sleep(2.0)  # Increased from 0.8 to 2.0
                            pythoncom.PumpWaitingMessages()
                            
                            # Step 3: Verify clipboard has data
                            if not self._verify_clipboard_has_data():
                                logger.warning("    Clipboard empty, waiting more...")
                                time.sleep(1.5)
                                pythoncom.PumpWaitingMessages()
                                
                                if not self._verify_clipboard_has_data():
                                    logger.error("    Clipboard still empty after additional wait")
                                    if paste_attempt < max_paste_attempts - 1:
                                        continue
                                    else:
                                        return False

                            # Step 4: Prepare position in Word
                            logger.info("    Preparing position in Word...")
                            selection.Collapse(0)  # wdCollapseEnd
                            selection.TypeParagraph()
                            
                            # Step 5: Try different paste methods
                            logger.info("    Attempting to paste chart...")
                            paste_methods = [
                                # Method 1: PasteSpecial without DataType (let Word decide)
                                {
                                    'name': 'PasteSpecial (auto)',
                                    'func': lambda: selection.PasteSpecial(
                                        Link=False,
                                        Placement=0,
                                        DisplayAsIcon=False
                                    )
                                },
                                # Method 2: PasteSpecial with wdPasteChart
                                {
                                    'name': 'PasteSpecial (wdPasteChart)',
                                    'func': lambda: selection.PasteSpecial(
                                        Link=False,
                                        DataType=10,  # wdPasteChart
                                        Placement=0,
                                        DisplayAsIcon=False
                                    )
                                },
                                # Method 3: PasteAndFormat
                                {
                                    'name': 'PasteAndFormat',
                                    'func': lambda: selection.PasteAndFormat(16)  # wdChart
                                },
                                # Method 4: Simple Paste
                                {
                                    'name': 'Paste (simple)',
                                    'func': lambda: selection.Paste()
                                }
                            ]
                            
                            pasted = False
                            last_paste_error = None
                            
                            for method in paste_methods:
                                try:
                                    logger.info(f"    Trying paste method: {method['name']}")
                                    method['func']()
                                    pasted = True
                                    logger.info(f"    Paste successful with method: {method['name']}")
                                    break
                                    
                                except Exception as method_error:
                                    last_paste_error = method_error
                                    logger.warning(f"    Method {method['name']} failed: {method_error}")
                                    
                                    # Try to undo and reset for next method
                                    try:
                                        word_doc.Application.Selection.Delete()
                                        selection.Collapse(0)
                                        selection.TypeParagraph()
                                    except:
                                        pass
                                    
                                    continue
                            
                            if not pasted:
                                logger.error(f"    All paste methods failed. Last error: {last_paste_error}")
                                if paste_attempt < max_paste_attempts - 1:
                                    logger.info("    Will retry entire copy-paste operation...")
                                    continue
                                else:
                                    return False
                            
                            # Wait for paste operation to complete
                            time.sleep(1.5)
                            pythoncom.PumpWaitingMessages()

                            # Step 6: Adjust size
                            logger.info("    Adjusting size...")
                            if selection.InlineShapes.Count > 0:
                                shape = selection.InlineShapes(1)
                                shape.Width = 425.2
                                shape.Height = 283.5
                                shape.Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
                                logger.info("    Size adjusted successfully")
                            else:
                                logger.warning("    No InlineShape found after paste")

                            # Step 7: Final cleanup
                            self._clear_clipboard()
                            pythoncom.PumpWaitingMessages()
                            time.sleep(0.5)

                            logger.info(f"  SUCCESS: Chart {chart_index} inserted correctly")
                            paste_success = True
                            break  # Exit paste attempt loop
                            
                        except Exception as paste_error:
                            logger.error(f"    Error in paste attempt {paste_attempt + 1}: {paste_error}")
                            self._clear_clipboard()
                            
                            if paste_attempt < max_paste_attempts - 1:
                                logger.info(f"    Retrying in 2 seconds...")
                                time.sleep(2.0)
                                pythoncom.PumpWaitingMessages()
                            else:
                                logger.error(f"    All {max_paste_attempts} paste attempts failed")
                            
                            continue
                    
                    return paste_success
                
                # Continue searching for next occurrence
                if not find_obj.Found:
                    break
            
            logger.warning(f"  '{search_text}' (occurrence {target_occurrence}) not found")
            logger.warning(f"     Occurrences found: {found_count}")
            return False

        except Exception as ex:
            logger.error(f"  Error in search: {ex}", exc_info=True)
            return False


    def _verify_clipboard_has_data(self):
        """
        Verify that the clipboard contains data before attempting to paste
        
        Returns:
            bool: True if clipboard has data, False otherwise
        """
        try:
            import win32clipboard
            
            win32clipboard.OpenClipboard()
            try:
                # Enumerate available formats
                available_formats = []
                format_id = 0
                while True:
                    format_id = win32clipboard.EnumClipboardFormats(format_id)
                    if format_id == 0:
                        break
                    available_formats.append(format_id)
                
                has_data = len(available_formats) > 0
                
                if has_data:
                    logger.info(f"    Clipboard has {len(available_formats)} formats available")
                    
                    # Common chart formats: CF_METAFILEPICT=3, CF_DIB=8, CF_ENHMETAFILE=14
                    chart_formats = [3, 8, 14]
                    has_chart = any(fmt in available_formats for fmt in chart_formats)
                    
                    if has_chart:
                        logger.info("    Clipboard contains chart format")
                    else:
                        logger.warning(f"    Available formats: {available_formats}")
                else:
                    logger.warning("    Clipboard is empty")
                
                return has_data
                
            finally:
                win32clipboard.CloseClipboard()
                
        except Exception as e:
            logger.warning(f"Could not verify clipboard: {e}")
            return True  # Assume it has data if we can't verify


    def _clear_clipboard(self):
        """
        Clear the Windows clipboard
        Uses multiple retry attempts for reliability
        """
        try:
            import win32clipboard
            max_attempts = 3
            
            for attempt in range(max_attempts):
                try:
                    win32clipboard.OpenClipboard()
                    win32clipboard.EmptyClipboard()
                    win32clipboard.CloseClipboard()
                    return
                    
                except Exception as e:
                    if attempt < max_attempts - 1:
                        time.sleep(0.2)
                        pythoncom.PumpWaitingMessages()
                    else:
                        logger.warning(f"Could not clear clipboard after {max_attempts} attempts: {e}")
                        
        except Exception as e:
            logger.warning(f"Error clearing clipboard: {e}")


    def _cleanup_resources(self):
        """
        Clean up system resources
        """
        try:
            self._clear_clipboard()
            pythoncom.PumpWaitingMessages()
            gc.collect()
            logger.info("Resources released")
            
        except Exception as e:
            logger.warning(f"Error releasing resources: {e}")


    def _get_graphic_mappings(self):
        """
        Returns the mapping of charts: which chart goes to which document location
        
        Returns:
            dict: Dictionary mapping sheet names to chart configurations
        """
        return {
            "IN SITU": {
                2: {
                    "text": "Gráfica 2. Comportamiento del Caudal en el Afluente y el Efluente",
                    "occurrence": 2
                },
                1: {
                    "text": "Gráfica 4. Comportamiento del pH en el Afluente y Efluente",
                    "occurrence": 2
                }
            },
            "GRÁFICAS": {
                3: {
                    "text": "Gráfica 3. Comportamiento del Caudal",
                    "occurrence": 2
                },
                4: {
                    "text": "Gráfica 5. Comportamiento del pH",
                    "occurrence": 2
                },
                6: {
                    "text": "Gráfica 6. Comportamiento de los Sólidos Sedimentables",
                    "occurrence": 2
                },
                5: {
                    "text": "Gráfica 7. Comportamiento de la Acidez y la Alcalinidad Total",
                    "occurrence": 2
                },
                7: {
                    "text": "Gráfica 8. Comportamiento del Cianuro Total",
                    "occurrence": 2
                },
                8: {
                    "text": "Gráfica 9. Comportamiento de los Cloruros",
                    "occurrence": 2
                },
                9: {
                    "text": "Gráfica 10. Comportamiento de los Sulfatos",
                    "occurrence": 2
                },
                10: {
                    "text": "Gráfica 11. Comportamiento de la DBO5 y DQO",
                    "occurrence": 2
                },
                27: {
                    "text": "Gráfica 12. Comportamiento de la Dureza Cálcica y Dureza Total",
                    "occurrence": 2
                },
                11: {
                    "text": "Gráfica 13. Comportamiento de los Fenoles",
                    "occurrence": 2
                },
                12: {
                    "text": "Gráfica 14. Comportamiento de los Sulfuros",
                    "occurrence": 2
                },
                13: {
                    "text": "Gráfica 15. Comportamiento de las Grasas y Aceites e Hidrocarburos",
                    "occurrence": 2
                },
                17: {
                    "text": "Gráfica 16. Comportamiento del Arsénico Total",
                    "occurrence": 2
                },
                22: {
                    "text": "Gráfica 17. Comportamiento del Cadmio Total",
                    "occurrence": 2
                },
                24: {
                    "text": "Gráfica 18. Comportamiento del Mercurio total.",
                    "occurrence": 2
                },
                16: {
                    "text": "Gráfica 19. Comportamiento del Níquel Total.",
                    "occurrence": 2
                },
                25: {
                    "text": "Gráfica 20. Comportamiento del Plomo Total.",
                    "occurrence": 2
                },
                18: {
                    "text": "Gráfica 21. Comportamiento del Selenio Total.",
                    "occurrence": 2
                },
                23: {
                    "text": "Gráfica 22. Comportamiento del Vanadio Total.",
                    "occurrence": 2
                },
                26: {
                    "text": "Gráfica 23. Comportamiento del Zinc Total.",
                    "occurrence": 2
                },
                19: {
                    "text": "Gráfica 24. Comportamiento del Cobre Total.",
                    "occurrence": 2
                },
                21: {
                    "text": "Gráfica 25. Comportamiento del Cromo Total",
                    "occurrence": 2
                },
                20: {
                    "text": "Gráfica 26. Comportamiento del Hierro Total.",
                    "occurrence": 2
                },
                14: {
                    "text": "Gráfica 27. Comportamiento del Nitrógeno Total",
                    "occurrence": 2
                },
                15: {
                    "text": "Gráfica 28. Comportamiento de los Sólidos Suspendidos Totales",
                    "occurrence": 2
                }
            }
        }


    def paste_graphics(self):
        """
        Legacy method - kept for compatibility
        """
        pass