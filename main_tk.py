import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
from db_manager import DBManager
from exporter import Exporter

class DataProcessorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Project X: Large Data Filter & Export (Tkinter Version)")
        self.geometry("1400x900")
        
        self.db = DBManager()
        self.exporter = Exporter()
        
        self.current_page = 1
        self.page_size = 50
        self.total_records = 0
        self.filter_rows = [] # List of tuples (col_var, op_var, val_var, frame)
        self.current_columns = None

        self._init_ui()
        self._check_existing_data()

    def _init_ui(self):
        # Header
        header_frame = ttk.Frame(self, padding=10)
        header_frame.pack(fill=tk.X)
        
        ttk.Label(header_frame, text="Large Data Processor", font=("Helvetica", 16, "bold")).pack(side=tk.LEFT)
        
        btn_frame = ttk.Frame(header_frame)
        btn_frame.pack(side=tk.RIGHT)
        
        ttk.Button(btn_frame, text="Chọn file Excel/CSV", command=self.pick_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Xuất Excel", command=self.export_file).pack(side=tk.LEFT, padx=5)

        # Status & Progress
        self.status_var = tk.StringVar(value="Sẵn sàng")
        self.progress_var = tk.DoubleVar(value=0)
        
        status_frame = ttk.Frame(self, padding=(10, 0))
        status_frame.pack(fill=tk.X)
        ttk.Label(status_frame, textvariable=self.status_var).pack(side=tk.LEFT)
        self.progress_bar = ttk.Progressbar(status_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=10)

        # Filters Area (Dynamic)
        self.filter_frame_container = ttk.LabelFrame(self, text="Bộ lọc nâng cao", padding=10)
        self.filter_frame_container.pack(fill=tk.X, padx=10, pady=5)
        
        # Scrollable area for filters if needed, but let's keep it simple first
        self.filter_list_frame = ttk.Frame(self.filter_frame_container)
        self.filter_list_frame.pack(fill=tk.X)
        
        filter_actions = ttk.Frame(self.filter_frame_container)
        filter_actions.pack(fill=tk.X, pady=5)
        ttk.Button(filter_actions, text="+ Thêm điều kiện", command=self._add_filter_row).pack(side=tk.LEFT)
        ttk.Button(filter_actions, text="Áp dụng bộ lọc", command=self._apply_filters).pack(side=tk.LEFT, padx=10)

        # Data Treeview
        tree_frame = ttk.Frame(self, padding=10)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        self.tree = ttk.Treeview(tree_frame, columns=("info"), show="headings")
        self.tree.heading("info", text="Thông tin")
        
        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.grid(column=0, row=0, sticky='nsew')
        vsb.grid(column=1, row=0, sticky='ns')
        hsb.grid(column=0, row=1, sticky='ew')
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)

        # Pagination
        nav_frame = ttk.Frame(self, padding=10)
        nav_frame.pack(fill=tk.X)
        
        center_nav = ttk.Frame(nav_frame)
        center_nav.pack(anchor=tk.CENTER)
        
        ttk.Button(center_nav, text="< Trước", command=self.prev_page).pack(side=tk.LEFT)
        self.page_label = ttk.Label(center_nav, text="Trang 1")
        self.page_label.pack(side=tk.LEFT, padx=10)
        ttk.Button(center_nav, text="Sau >", command=self.next_page).pack(side=tk.LEFT)

    def _check_existing_data(self):
        if self.db.check_existing_db():
            self.status_var.set("Đã tải dữ liệu từ phiên trước.")
            self._apply_filters() # Load data
        else:
            self.status_var.set("Chưa có dữ liệu. Vui lòng chọn file.")

    def pick_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel/CSV Files", "*.csv *.xlsx *.xls")])
        if file_path:
            self.status_var.set(f"Đang phân tích: {os.path.basename(file_path)}...")
            self.progress_bar.start(10)
            threading.Thread(target=self._process_file_thread, args=(file_path,), daemon=True).start()

    def _count_lines(self, filename):
        f = open(filename, 'rb')
        lines = 0
        buf_size = 1024 * 1024
        read_f = f.raw.read
        buf = read_f(buf_size)
        while buf:
            lines += buf.count(b'\n')
            buf = read_f(buf_size)
        f.close()
        return lines

    def _process_file_thread(self, file_path):
        try:
            self.progress_bar.stop()
            self.progress_bar['mode'] = 'determinate'
            self.progress_var.set(0)
            
            # Estimate total rows
            total_rows = 0
            ext = os.path.splitext(file_path)[1].lower()
            if ext == '.csv':
                self.after(0, lambda: self.status_var.set("Đang đếm số dòng..."))
                total_rows = self._count_lines(file_path)
            
            self.after(0, lambda: self.status_var.set(f"Đang import... (Tổng: {total_rows} dòng)"))

            def progress_cb(processed):
                if total_rows > 0:
                    pct = (processed / total_rows) * 100
                    self.after(0, lambda p=pct: self.progress_var.set(p))
                    self.after(0, lambda p=processed: self.status_var.set(f"Đã import {p}/{total_rows} dòng ({int((p/total_rows)*100)}%)"))
                else:
                    self.after(0, lambda p=processed: self.status_var.set(f"Đã import {p} dòng"))
            
            self.db.import_file(file_path, progress_callback=progress_cb)
            self.after(0, self._on_import_success)
            
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Lỗi", str(e)))
            self.after(0, lambda: self.status_var.set("Lỗi import"))
        finally:
            self.after(0, self.progress_bar.stop)

    def _on_import_success(self):
        self.status_var.set("Import thành công!")
        # Clear filters on new import? Maybe keep them? Let's clear to avoid stale cols
        self._clear_filters()
        self.current_page = 1
        self._update_table()

    def _clear_filters(self):
        for *_, frame in self.filter_rows:
            frame.destroy()
        self.filter_rows.clear()
        self._add_filter_row() # Add one empty row

    def _add_filter_row(self):
        cols = self.db.get_columns()
        if not cols: return

        frame = ttk.Frame(self.filter_list_frame)
        frame.pack(fill=tk.X, pady=2)
        
        col_var = tk.StringVar()
        col_cb = ttk.Combobox(frame, textvariable=col_var, values=cols, width=15, state="readonly")
        col_cb.set(cols[0])
        col_cb.pack(side=tk.LEFT, padx=2)
        
        op_var = tk.StringVar(value="contains")
        ops = ["=", ">", "<", ">=", "<=", "contains", "starts_with", "ends_with", "!="]
        op_cb = ttk.Combobox(frame, textvariable=op_var, values=ops, width=10, state="readonly")
        op_cb.pack(side=tk.LEFT, padx=2)
        
        val_var = tk.StringVar()
        entry = ttk.Entry(frame, textvariable=val_var, width=20)
        entry.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        
        def delete_row(f=frame):
            f.destroy()
            self.filter_rows = [r for r in self.filter_rows if r[3] != f]

        ttk.Button(frame, text="X", width=3, command=delete_row).pack(side=tk.LEFT, padx=2)
        
        self.filter_rows.append((col_var, op_var, val_var, frame))

    def _apply_filters(self):
        self.current_page = 1
        self._update_table()

    def _get_current_filters(self):
        filters = []
        for col_var, op_var, val_var, _ in self.filter_rows:
            col = col_var.get()
            val = val_var.get().strip()
            if col and val:
                filters.append({
                    'col': col,
                    'op': op_var.get(),
                    'val': val
                })
        return filters

    def _update_table(self):
        # Determine columns
        cols = self.db.get_columns()
        
        # Only reset columns if they have changed (e.g. new file)
        # This preserves user's manual resizing during filtering
        if cols != self.current_columns:
            self.tree["columns"] = cols
            self.current_columns = cols
            
            # Smart auto-width
            # Simple heuristic: header length * 10, capped between 100 and 400
            for col in cols:
                self.tree.heading(col, text=col)
                # Estimate width roughly based on header
                header_width = len(col) * 10 + 20
                width = max(100, min(header_width, 400))
                self.tree.column(col, width=width, anchor='w')
        
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        filters = self._get_current_filters()
        rows, total = self.db.get_data_paginated(self.current_page, self.page_size, filters)
        self.total_records = total
        
        for row in rows:
            self.tree.insert("", "end", values=row)
            
        total_pages = (total // self.page_size) + 1 if total > 0 else 1
        self.page_label.config(text=f"Trang {self.current_page} / {total_pages} (Tổng: {total})")

    def prev_page(self):
        if self.current_page > 1:
            self.current_page -= 1
            self._update_table()

    def next_page(self):
        if self.current_page * self.page_size < self.total_records:
            self.current_page += 1
            self._update_table()
            
    def export_file(self):
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Lưu file Excel"
        )
        if not output_path:
            return

        self.status_var.set("Đang xuất file...")
        self.progress_bar['mode'] = 'indeterminate'
        self.progress_bar.start(10)
        threading.Thread(target=self._export_thread, args=(output_path,), daemon=True).start()
        
    def _export_thread(self, output_path):
        try:
            filters = self._get_current_filters()
            df = self.db.export_filtered_data(filters)
            self.exporter.export_to_excel(df, output_path)
            self.after(0, lambda: messagebox.showinfo("Thành công", f"Đã xuất file: {output_path}"))
            self.after(0, lambda: self.status_var.set("Xuất file hoàn tất"))
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Lỗi xuất file", str(e)))
            self.after(0, lambda: self.status_var.set("Lỗi xuất file"))
        finally:
             self.after(0, self.progress_bar.stop)

if __name__ == "__main__":
    app = DataProcessorApp()
    app.mainloop()
