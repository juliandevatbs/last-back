import logging
import win32com.client as win32
from pathlib import Path
import time

from core.exceptions import KeyNotFound

logger = logging.getLogger(__name__)


class CopyGraphics:

    def __init__(self, excel_obj, json_config, word_to_write, word_path):
        self.excel_graph = excel_obj
        self.json_config = json_config
        self.excel_path = self.json_config["RUTA_EXCEL_GRAFICAS"]
        self.word_obj = word_to_write
        self.word_path = word_path

    def copy_and_insert_editable_charts(self):
        """if not self.excel_graph:
            logger.error("No se suministro excel de graficas", exc_info=True)
            raise ValueError("Excel object is None")
        """

        if self.excel_graph is None:
            logger.error("No se suministro excel de graficas", exc_info=True)
            raise ValueError("Excel object is None")

        if self.json_config is None:
            logger.error("No se suministro json con configuración de graficas", exc_info=True)
            raise ValueError("JSON config is None")

        if self.word_obj is None:
            logger.error("No se suministro objeto Word", exc_info=True)
            raise ValueError("Word object is None")

        try:
            # NO cerramos el Word, trabajamos directamente con el que ya está abierto
            logger.info("Construyendo registro de gráficas...")
            chart_registry = self._build_chart_registry(self.excel_graph)

            # Insertar gráficas en el Word actual
            logger.info("Insertando gráficas en Word...")
            self._insert_charts_in_word(self.word_obj, chart_registry)

            # Guardar el documento
            logger.info("Guardando documento Word...")
            self.word_obj.Save()

            logger.info("Todas las graficas insertadas como objetos editables correctamente")

        except KeyNotFound:
            logger.error("Hojas de graficas no existen en el excel suministrado", exc_info=True)
            raise

        except Exception as ex:
            logger.error(f"Error copiando graficas del excel: {ex}", exc_info=True)
            raise

    def _build_chart_registry(self, excel_workbook):
        chart_registry = {}
        sheets = self.json_config["HOJAS_GRAFICAS"].values()

        for sheet_name in sheets:
            try:
                # Acceso correcto con win32com
                sheet = excel_workbook.Sheets(sheet_name)
            except Exception as e:
                logger.error(f"Hoja {sheet_name} no encontrada: {e}")
                continue

            # Contar gráficos en la hoja usando ChartObjects (win32com)
            try:
                chart_objects = sheet.ChartObjects()
                chart_count = chart_objects.Count
            except Exception as e:
                logger.error(f"Error accediendo a ChartObjects en {sheet_name}: {e}")
                continue

            if chart_count == 0:
                logger.warning(f"No se encontraron graficas en {sheet_name}")
                continue

            if sheet_name not in chart_registry:
                chart_registry[sheet_name] = {}

            # Registrar cada gráfico (índice base 1 en win32com)
            for i in range(1, chart_count + 1):
                try:
                    chart_obj = chart_objects(i)
                    chart_registry[sheet_name][i] = chart_obj
                    logger.info(f"Registrada grafica {i} de hoja {sheet_name}")
                except Exception as chart_error:
                    logger.error(f"Error obteniendo grafica {i} de hoja {sheet_name}: {chart_error}")

        return chart_registry

    def _insert_charts_in_word(self, word_doc, chart_registry):
        graphic_mappings = self._get_graphic_mappings()

        for sheet_name, charts in graphic_mappings.items():
            if sheet_name not in chart_registry:
                logger.warning(f"Hoja {sheet_name} no encontrada en registro")
                continue

            for chart_index, config in charts.items():
                try:
                    if chart_index not in chart_registry[sheet_name]:
                        logger.warning(f"Grafica {chart_index} no encontrada en hoja {sheet_name}")
                        continue

                    logger.info(f"Procesando grafica {chart_index} de {sheet_name}")

                    chart_obj = chart_registry[sheet_name][chart_index]

                    search_text = config["text"]  # text en minúscula
                    target_occurrence = config["occurrence"]

                    if self._insert_chart_at_location(word_doc, chart_obj, search_text, target_occurrence, chart_index,
                                                      sheet_name):
                        logger.info(f"Grafica {chart_index} de {sheet_name} insertada correctamente")
                    else:
                        logger.warning(f"No se encontro ubicación para grafica {chart_index} de {sheet_name}")

                except Exception as ex:
                    logger.error(f"Error insertando grafica {chart_index} de {sheet_name}: {ex}", exc_info=True)
                    continue

    def _insert_chart_at_location(self, word_doc, chart_obj, search_text, target_occurrence, chart_index, sheet_name):
        found_count = 0

        try:
            paragraphs = word_doc.Paragraphs
            total_paragraphs = paragraphs.Count

            logger.info(f"Buscando '{search_text}' en {total_paragraphs} párrafos")

            for i in range(1, total_paragraphs + 1):
                para = paragraphs(i)
                para_text = para.Range.Text

                if search_text.lower() in para_text.lower():
                    found_count += 1
                    logger.info(f"Encontrado '{search_text}' - ocurrencia {found_count}")

                    if found_count == target_occurrence:
                        try:
                            logger.info(f"Insertando gráfica en ocurrencia {target_occurrence}")

                            # Copiar el gráfico
                            chart_obj.Copy()
                            time.sleep(0.5)

                            # Obtener el rango del párrafo
                            para_range = para.Range

                            # Posicionar después del párrafo
                            para_range.Collapse(0)  # 0 = wdCollapseStart
                            para_range.InsertParagraphAfter()
                            para_range.Move(1, 1)  # 1 = wdParagraph
                            para_range.Select()

                            # Pegar como objeto de Excel (editable)
                            word_doc.Application.Selection.PasteSpecial(
                                Link=False,
                                DataType=10,  # wdPasteChart
                                Placement=0,  # wdInLine
                                DisplayAsIcon=False
                            )

                            time.sleep(0.5)

                            # Ajustar tamaño y centrar
                            if word_doc.Application.Selection.InlineShapes.Count > 0:
                                shape = word_doc.Application.Selection.InlineShapes(1)
                                shape.Width = 425.2
                                shape.Height = 283.5
                                shape.Range.ParagraphFormat.Alignment = 1

                            logger.info(f"✓ Grafica {chart_index} de {sheet_name} pegada exitosamente")
                            return True

                        except Exception as paste_error:
                            logger.error(f"Error pegando grafica {chart_index}: {paste_error}", exc_info=True)
                            return False

        except Exception as ex:
            logger.error(f"Error buscando ubicación: {ex}", exc_info=True)
            return False

            logger.warning(f"✗ No se encontró '{search_text}' (ocurrencia {target_occurrence}) en el documento")
            return False

        except Exception as ex:
            logger.error(f"Error buscando ubicación: {ex}", exc_info=True)
            return False

            logger.warning(f"No se encontró '{search_Text}' (ocurrencia {target_occurrence})")
            return False

    def _get_graphic_mappings(self):

        return {
            "IN SITU": {
                2: {
                    "Text": "Gráfica 2. Comportamiento del Caudal en el Afluente y el Efluente",
                    "occurrence": 2
                },
                1: {
                    "Text": "Gráfica 4. Comportamiento del pH en el Afluente y Efluente",
                    "occurrence": 2
                }
            },
            "GRÁFICAS": {
                5: {
                    "Text": "Gráfica 3. Comportamiento del Caudal",
                    "occurrence": 2
                },
                4: {
                    "Text": "Gráfica 5. Comportamiento del pH",
                    "occurrence": 2
                },
                6: {
                    "Text": "Gráfica 6. Comportamiento de los Sólidos Sedimentables",
                    "occurrence": 2
                },
                3: {
                    "Text": "Gráfica 7. Comportamiento de la Acidez y la Alcalinidad Total",
                    "occurrence": 2
                },
                7: {
                    "Text": "Gráfica 8. Comportamiento del Cianuro Total",
                    "occurrence": 2
                },
                8: {
                    "Text": "Gráfica 9. Comportamiento de los Cloruros",
                    "occurrence": 2
                },
                9: {
                    "Text": "Gráfica 10. Comportamiento de los Sulfatos",
                    "occurrence": 2
                },
                10: {
                    "Text": "Gráfica 11. Comportamiento de la DBO5 y DQO",
                    "occurrence": 2
                },
                27: {
                    "Text": "Gráfica 12. Comportamiento de la Dureza Cálcica y Dureza Total",
                    "occurrence": 2
                },
                11: {
                    "Text": "Gráfica 13. Comportamiento de los Fenoles",
                    "occurrence": 2
                },
                12: {
                    "Text": "Gráfica 14. Comportamiento de los Sulfuros",
                    "occurrence": 2
                },
                13: {
                    "Text": "Gráfica 15. Comportamiento de las Grasas y Aceites e Hidrocarburos",
                    "occurrence": 2
                },
                17: {
                    "Text": "Gráfica 16. Comportamiento del Arsénico Total",
                    "occurrence": 2
                },
                22: {
                    "Text": "Gráfica 17. Comportamiento del Cadmio Total",
                    "occurrence": 2
                },
                24: {
                    "Text": "Gráfica 18. Comportamiento del Mercurio total.",
                    "occurrence": 2
                },
                16: {
                    "Text": "Gráfica 19. Comportamiento del Níquel Total.",
                    "occurrence": 2
                },
                25: {
                    "Text": "Gráfica 20. Comportamiento del Plomo Total.",
                    "occurrence": 2
                },
                18: {
                    "Text": "Gráfica 21. Comportamiento del Selenio Total.",
                    "occurrence": 2
                },
                23: {
                    "Text": "Gráfica 22. Comportamiento del Vanadio Total.",
                    "occurrence": 2
                },
                26: {
                    "Text": "Gráfica 23. Comportamiento del Zinc Total.",
                    "occurrence": 2
                },
                19: {
                    "Text": "Gráfica 24. Comportamiento del Cobre Total.",
                    "occurrence": 2
                },
                21: {
                    "Text": "Gráfica 25. Comportamiento del Cromo Total",
                    "occurrence": 2
                },
                20: {
                    "Text": "Gráfica 26. Comportamiento del Hierro Total.",
                    "occurrence": 2
                },
                14: {
                    "Text": "Gráfica 27. Comportamiento del Nitrógeno Total",
                    "occurrence": 2
                },
                15: {
                    "Text": "Gráfica 28. Comportamiento de los Sólidos Suspendidos Totales",
                    "occurrence": 2
                }
            }
        }

    def paste_graphics(self):
        pass