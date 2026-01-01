import pandas as pd
from typing import Generator

class Exporter:
    def __init__(self, max_rows_per_sheet: int = 1000000):
        self.max_rows_per_sheet = max_rows_per_sheet

    def export_to_excel(self, df: pd.DataFrame, output_path: str):
        """
        Exports a DataFrame to Excel, splitting into multiple sheets 
        if row count exceeds max_rows_per_sheet.
        """
        num_rows = len(df)
        
        if num_rows <= self.max_rows_per_sheet:
            df.to_excel(output_path, index=False)
        else:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for i in range(0, num_rows, self.max_rows_per_sheet):
                    sheet_name = f"Sheet_{i // self.max_rows_per_sheet + 1}"
                    chunk = df.iloc[i : i + self.max_rows_per_sheet]
                    chunk.to_excel(writer, sheet_name=sheet_name, index=False)
                    
    def export_from_generator(self, data_iterator, columns, output_path: str):
        """
        (Optional) For even better RAM efficiency, we could stream from SQLite 
        directly to Excel if needed, but pd.read_sql_query for the filtered set 
        is usually acceptable for modern machines unless the filtered set itself 
        is >10GB.
        """
        # Logic for streaming can be added if memory pressure is too high
        pass
