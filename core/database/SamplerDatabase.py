import pyodbc
import os

class SamplerDatabase:
    def __init__(self):
        self.connection = pyodbc.connect(
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={os.getenv('SAMPLER_DATABASE_HOST')};"
            f"DATABASE={os.getenv('SAMPLER_DATABASE_NAME')};"
            f"UID={os.getenv('SAMPLER_DATABASE_USER')};"
            f"PWD={os.getenv('SAMPLER_DATABASE_PASSWORD')};"
        )
        self.cursor = self.connection.cursor()

    def execute_query(self, query, params=None):
        self.cursor.execute(query, params or [])
        try:
            result = self.cursor.fetchall()
            return result
        except pyodbc.ProgrammingError:
            self.connection.commit()
            return None

    def close_connection(self):
        self.cursor.close()
        self.connection.close()
