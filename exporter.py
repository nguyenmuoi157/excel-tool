import pandas as pd


class Exporter:
    def __init__(self, max_rows_per_sheet: int = 1000000):
        self.max_rows_per_sheet = max_rows_per_sheet

    def export_to_excel(self, df: pd.DataFrame, output_path: str, progress_callback=None):
        """Exports a DataFrame to Excel using xlsxwriter (faster than openpyxl)."""
        num_rows = len(df)

        writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
        try:
            for i in range(0, num_rows, self.max_rows_per_sheet):
                sheet_name = f"Sheet_{i // self.max_rows_per_sheet + 1}"
                chunk = df.iloc[i : i + self.max_rows_per_sheet]
                chunk.to_excel(writer, sheet_name=sheet_name, index=False)
                if progress_callback:
                    progress_callback(min(i + self.max_rows_per_sheet, num_rows))
        finally:
            writer.close()

    def export_from_generator(self, data_iterator, columns, output_path: str, progress_callback=None):
        """Stream chunks from SQLite to Excel without loading all data into RAM.

        Writes each SQL chunk directly to the Excel sheet using startrow,
        so only one chunk (100K rows) is held in memory at a time.
        """
        writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
        try:
            sheet_idx = 1
            current_row_in_sheet = 0
            total_rows = 0
            total_from_generator = 0
            sheet_name = f"Sheet_{sheet_idx}"

            for chunk in data_iterator:
                chunk_len = len(chunk)
                total_from_generator += chunk_len
                remaining_in_sheet = self.max_rows_per_sheet - current_row_in_sheet

                while len(chunk) > 0:
                    rows_to_write = min(len(chunk), remaining_in_sheet)
                    batch = chunk.iloc[:rows_to_write]

                    if current_row_in_sheet == 0:
                        # header=True: xlsxwriter writes header at startrow=0, data at rows 1 to len(batch).
                        # Next write position = len(batch) + 1 (accounting for header at row 0).
                        current_row_in_sheet = rows_to_write + 1
                        batch.to_excel(writer, sheet_name=sheet_name, index=False,
                                       startrow=0)
                    else:
                        batch.to_excel(writer, sheet_name=sheet_name, index=False,
                                       header=False, startrow=current_row_in_sheet)
                        current_row_in_sheet += rows_to_write

                    total_rows += rows_to_write

                    if progress_callback:
                        progress_callback(total_rows)

                    if current_row_in_sheet >= self.max_rows_per_sheet:
                        sheet_idx += 1
                        sheet_name = f"Sheet_{sheet_idx}"
                        current_row_in_sheet = 0
                        remaining_in_sheet = self.max_rows_per_sheet

                    if len(chunk) > rows_to_write:
                        chunk = chunk.iloc[rows_to_write:]
                        remaining_in_sheet = self.max_rows_per_sheet - current_row_in_sheet
                    else:
                        chunk = pd.DataFrame()

            # Empty result set
            if total_rows == 0 and columns:
                empty_df = pd.DataFrame(columns=columns)
                empty_df.to_excel(writer, sheet_name="Sheet_1", index=False)

            if progress_callback:
                progress_callback(total_rows)
        finally:
            writer.close()
            del writer
        return total_from_generator, total_rows