import sqlite3
import pandas as pd
import os
from typing import List, Tuple, Optional, Any

class DBManager:
    def __init__(self, db_name: str = "temp_data.db"):
        # Use a safe, writable user directory
        user_dir = os.path.join(os.path.expanduser("~"), ".large_data_processor")
        if not os.path.exists(user_dir):
            os.makedirs(user_dir)
        
        self.db_path = os.path.join(user_dir, db_name)
        self.table_name = "data_table"
        self.columns = []

    def _get_connection(self):
        return sqlite3.connect(self.db_path)

    def create_indexes(self):
        """Creates indexes on all columns. Call after import_file() completes."""
        if not self.columns:
            return
        with self._get_connection() as conn:
            for col in self.columns:
                conn.execute(f'CREATE INDEX IF NOT EXISTS "idx_{col}" ON "{self.table_name}" ("{col}")')

    def check_existing_db(self) -> bool:
        """Checks if a DB file exists and loads column info if it does."""
        if os.path.exists(self.db_path):
            try:
                self.get_columns()
                return True
            except Exception:
                return False
        return False

    def import_file(self, file_path: str, chunk_size: int = 50000, progress_callback=None, total_rows_est=None):
        """Imports a CSV or Excel file into SQLite using chunks."""
        # Only remove old DB if we are importing a NEW file. 
        # The caller decides when to call this.
        if os.path.exists(self.db_path):
            os.remove(self.db_path)

        extension = os.path.splitext(file_path)[1].lower()
        processed_rows = 0
        
        # Determine how to read the file
        if extension == '.csv':
            reader = pd.read_csv(file_path, chunksize=chunk_size, low_memory=False)
            
            with self._get_connection() as conn:
                first_chunk = True
                for chunk in reader:
                    # Clean column names to be SQL-safe (simple replace)
                    chunk.columns = [c.strip() for c in chunk.columns]
                    
                    if first_chunk:
                        self.columns = chunk.columns.tolist()
                        chunk.to_sql(self.table_name, conn, if_exists='replace', index=False)
                        first_chunk = False
                    else:
                        chunk.to_sql(self.table_name, conn, if_exists='append', index=False)
                    
                    processed_rows += len(chunk)
                    if progress_callback:
                        progress_callback(processed_rows)

            self.create_indexes()

        elif extension in ['.xlsx', '.xls']:
            import openpyxl
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            ws = wb.active

            columns = None
            rows_buffer_values = []
            first_chunk = True

            try:
                for row in ws.iter_rows(values_only=True):
                    if columns is None:
                        columns = [str(c).strip() if c is not None else f"col_{j}"
                                   for j, c in enumerate(row)]
                        self.columns = columns
                    else:
                        rows_buffer_values.append(row)

                    if len(rows_buffer_values) >= chunk_size:
                        temp_df = pd.DataFrame(rows_buffer_values, columns=columns)
                        rows_buffer_values = []

                        with self._get_connection() as conn:
                            temp_df.to_sql(self.table_name, conn,
                                          if_exists='replace' if first_chunk else 'append',
                                          index=False)
                            first_chunk = False

                        processed_rows += len(temp_df)
                        if progress_callback:
                            progress_callback(processed_rows)

                # Flush remaining rows
                if rows_buffer_values:
                    temp_df = pd.DataFrame(rows_buffer_values, columns=columns)
                    with self._get_connection() as conn:
                        temp_df.to_sql(self.table_name, conn,
                                      if_exists='replace' if first_chunk else 'append',
                                      index=False)
                    processed_rows += len(temp_df)
                    if progress_callback:
                        progress_callback(processed_rows)
            finally:
                wb.close()
                del wb

            self.create_indexes()
        else:
            raise ValueError("Unsupported file format")

    def _build_where_clause(self, filters: List[dict]) -> Tuple[str, List[Any]]:
        """
        Builds SQL WHERE clause from a list of filtered dicts.
        Logic:
        - Conditions for the SAME column are joined by OR.
        - Conditions for DIFFERENT columns are joined by AND.
        """
        if not filters:
            return "", []
        
        from collections import defaultdict
        col_groups = defaultdict(list)
        
        for f in filters:
            col = f.get('col')
            op = f.get('op')
            val = f.get('val')
            
            if not col or not val:
                continue
            
            condition = ""
            param = None
            
            if op == '=':
                condition = f'"{col}" = ?'
                param = val
            elif op in ['>', '<', '>=', '<=']:
                condition = f'"{col}" {op} ?'
                param = val
            elif op == 'contains':
                condition = f'"{col}" LIKE ?'
                param = f'%{val}%'
            elif op == 'starts_with':
                condition = f'"{col}" LIKE ?'
                param = f'{val}%'
            elif op == 'ends_with':
                condition = f'"{col}" LIKE ?'
                param = f'%{val}'
            elif op == '!=':
                condition = f'"{col}" != ?'
                param = val
            
            if condition:
                col_groups[col].append((condition, param))

        if not col_groups:
            return "", []

        final_conditions = []
        final_params = []

        for col, group in col_groups.items():
            # Join all conditions for this column with OR
            or_conditions = [item[0] for item in group]
            group_params = [item[1] for item in group]
            
            # Wrap in parentheses
            clause = "(" + " OR ".join(or_conditions) + ")"
            final_conditions.append(clause)
            final_params.extend(group_params)

        return " WHERE " + " AND ".join(final_conditions), final_params

    def get_data_paginated(self, page: int, page_size: int, filters: List[dict] = None) -> Tuple[List, int]:
        """Retrieves paginated data based on complex filters."""
        query = f'SELECT * FROM "{self.table_name}"'
        count_query = f'SELECT COUNT(*) FROM "{self.table_name}"'
        
        where_clause, params = self._build_where_clause(filters)
        
        query += where_clause
        count_query += where_clause

        offset = (page - 1) * page_size
        query += f" LIMIT {page_size} OFFSET {offset}"

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(count_query, params)
            total_count = cursor.fetchone()[0]
            
            cursor.execute(query, params)
            rows = cursor.fetchall()
            
        return rows, total_count

    def get_columns(self):
        # Always fetch from DB if self.columns is empty (e.g. startup)
        if not self.columns:
            try:
                with self._get_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{self.table_name}'")
                    if cursor.fetchone():
                        cursor.execute(f"PRAGMA table_info({self.table_name})")
                        self.columns = [row[1] for row in cursor.fetchall()]
            except Exception:
                pass
        return self.columns

    def export_filtered_data(self, filters: List[dict] = None):
        """Returns dataframe for the filtered set (full in-memory)."""
        query = f'SELECT * FROM "{self.table_name}"'
        where_clause, params = self._build_where_clause(filters)
        query += where_clause

        with self._get_connection() as conn:
            return pd.read_sql_query(query, conn, params=params)

    def export_filtered_data_streaming(self, filters: List[dict] = None, chunk_size: int = 100000):
        """Yields DataFrames chunk-by-chunk from SQLite for memory-efficient export."""
        query = f'SELECT * FROM "{self.table_name}"'
        where_clause, params = self._build_where_clause(filters)
        query += where_clause

        with self._get_connection() as conn:
            for chunk in pd.read_sql_query(query, conn, params=params, chunksize=chunk_size):
                yield chunk
