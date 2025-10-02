import io
import json
from typing import Optional, Dict, Any

from openpyxl.reader.excel import load_workbook

from read_data.services.excel_reader import data_constructor


class JsonBuilder():
    """Manages excel data extraction and JSON configuration updates"""

    def __init__(self, file_bytes):
        """
        Initialize JsonBuilder with excel file bytes

        args:
            file_bytes = Excel file content as bytes
        """
        self.config_path = 'fields_config/fields.json'

        # Load and parse excel data
        self.excel_object = load_workbook(io.BytesIO(file_bytes), data_only=True)
        self.super_data = data_constructor(self.excel_object)

        # Extract specific data sections
        self.main_data = self.super_data.get("main_data", {})
        self.sampling_data = self.super_data.get("sampling_data", {})
        self.samples = self.super_data.get("samples", {})
        self.surveillance_data = self.super_data.get("surveillance_data", {})

        # JSON configuration (loaded lazily)
        self._json_config: Optional[Dict[str, Any]] = None
        self._json_config_labels: Optional[Dict[str, Any]] = None

    @property
    def json_config(self) -> Dict[str, Any]:
        """Get the JSON config, loading it if not already loaded"""
        if self._json_config is None:
            self.load_json()
        return self._json_config

    @property
    def json_config_labels(self) -> Dict[str, Any]:

        """Get the fields section of the json config"""
        if self._json_config_labels is None:
            # Ensure json_config is loaded first
            _ = self.json_config
        return self._json_config_labels

    def load_json(self) -> None:

        """Load json configuration from file"""

        try:
            with open(self.config_path, 'r', encoding='utf-8') as json_file:
                self._json_config = json.load(json_file)
                self._json_config_labels = self._json_config.get("fields", {})
        except FileNotFoundError:
            raise FileNotFoundError(
                f"Configuration file not found: {self.config_path}"
            )
        except json.JSONDecodeError as e:
            raise ValueError(
                f"Invalid JSON in configuration file: {e}"
            )

    def update_json(self):

        """Write the excel data to the JSON config"""
        # Ensure json is loaded
        _ = self.json_config

        # WRITE THE EXCEL DATA TO THE CONFIG JSON
        for excel_key, excel_value in self.main_data.items() | self.sampling_data.items():
            self._json_config_labels[excel_key] = excel_value

        # SAVE THE JSON
        with open(self.config_path, 'w', encoding='utf-8') as json_file:
            json.dump(self._json_config, json_file, indent=4, ensure_ascii=False)

    def clean_json(self):

        """Reset all fields in the JSON config to None"""
        # Ensure json is loaded
        _ = self.json_config

        for key in self._json_config_labels.keys():

            self._json_config_labels[key] = None

        # SAVE THE JSON
        with open(self.config_path, 'w', encoding='utf-8') as json_file:

            json.dump(self._json_config, json_file, indent=4, ensure_ascii=False)