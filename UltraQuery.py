import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import pandas as pd
import requests
import io

# Ganti dengan API key OpenRouter kamu
OPENROUTER_API_KEY = "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"

class PowerQueryPivot:
    def __init__(self, root, show_ai=True):
        self.root = root
        self.root.title("Power Query Lite with Pivot (Excel)")
        self.df = None
        self.pivot_df = None
        self.excel_file = None
        self.show_ai = show_ai


        # --- Frame utama dengan panel kanan ---
        main_container = tk.Frame(self.root)
        main_container.pack(fill="both", expand=True)

        # Frame kiri: scrollable utama
        outer_frame = tk.Frame(main_container)
        outer_frame.pack(side="left", fill="both", expand=True)

        canvas = tk.Canvas(outer_frame)
        canvas.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(outer_frame, orient="vertical", command=canvas.yview)
        scrollbar.pack(side="right", fill="y")

        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        self.main_frame = tk.Frame(canvas)
        main_frame_id = canvas.create_window((0, 0), window=self.main_frame, anchor="nw")
        
        self.main_frame.bind("<Configure>", lambda event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfig(main_frame_id, width=event.width))

        # Frame kanan: panel catatan RL 4 & RL 5
        note_panel = tk.Frame(main_container, width=260, bg="#f8f8e7", bd=2, relief="groove")
        note_panel.pack(side="right", fill="y", padx=5, pady=5)
        note_panel.pack_propagate(False)
        note_title = tk.Label(note_panel, text="Catatan Laporan RL 4 & RL 5", font=("Segoe UI", 11, "bold"), bg="#f8f8e7")
        note_title.pack(pady=(10, 5))
        note_text = (
            "RL 4:\n"
            "- SEX\n"
            "- DISCHARGE STATUS\n"
            "- NAMA_PASIEN\n"
            "- MRN\n"
            "- UMUR TAHUN\n"
            "- UMUR HARI\n"
            "- SEP\n"
            "- DIAGLIST 1\n\n"
            "RL 5:\n"
            "- SEX\n"
            "- DISCHARGE STATUS\n"
            "- NAMA_PASIEN\n"
            "- MRN\n"
            "- UMUR TAHUN\n"
            "- UMUR HARI\n"
            "- SEP\n"
            "- DIAGLIST 1\n"
            "- DIAGLIST 2"
        )
        note_label = tk.Label(note_panel, text=note_text, justify="left", anchor="nw", bg="#f8f8e7", font=("Segoe UI", 10))
        note_label.pack(fill="both", expand=True, padx=10, pady=5)

        # --- Fitur AI (opsional) ---
        if show_ai:
            btn_ai = tk.Button(self.main_frame, text="🤖 Tanya AI tentang Data", command=self.ask_ai)
            btn_ai.pack(pady=5)
            self.ai_result_box = scrolledtext.ScrolledText(self.main_frame, height=7)
            self.ai_result_box.pack(padx=7, pady=5, fill="both", expand=True)
        else:
            self.ai_result_box = None

        # Tombol load file (bisa pilih hingga 5 file .txt/.xlsx/.xls)
        btn_load = tk.Button(self.main_frame, text="Load Files (<=5)", command=self.load_excel)
        btn_load.pack(pady=5)

        # --- Frame pilih sheet ---
        frame_sheet = tk.Frame(self.main_frame)
        frame_sheet.pack(pady=5)
        tk.Label(frame_sheet, text="Pilih Sheet:").grid(row=0, column=0)
        self.sheet_cb = ttk.Combobox(frame_sheet, state="readonly", width=25)
        self.sheet_cb.grid(row=0, column=1)
        self.sheet_cb.bind("<<ComboboxSelected>>", self.on_sheet_selected)

        # --- Frame pilih baris header ---
        frame_header = tk.Frame(self.main_frame)
        frame_header.pack(pady=5)
        tk.Label(frame_header, text="Baris Header (angka, mulai dari 1):").grid(row=0, column=0)
        self.header_entry = tk.Entry(frame_header, width=5)
        self.header_entry.grid(row=0, column=1)
        self.header_entry.insert(0, "1")  # Default header di baris 1
        self.header_entry.bind("<Return>", self.on_sheet_selected)
        btn_header_ok = tk.Button(frame_header, text="OK", command=self.on_sheet_selected)
        btn_header_ok.grid(row=0, column=2, padx=5)

        # --- Frame untuk Split Kolom ---
        frame_split = tk.Frame(self.main_frame)
        frame_split.pack(pady=5)
        tk.Label(frame_split, text="Split Kolom:").grid(row=0, column=0)
        self.split_cb = ttk.Combobox(frame_split, state="readonly", width=15)
        self.split_cb.grid(row=0, column=1)
        btn_split = tk.Button(frame_split, text="Split by ;", command=self.split_column)
        btn_split.grid(row=0, column=2, padx=5)

        # --- Frame untuk Pilih Kolom Aktif ---
        frame_select = tk.Frame(self.main_frame)
        frame_select.pack(pady=5)
        tk.Label(frame_select, text="Pilih Kolom yang Ditampilkan:").grid(row=0, column=0)
        self.select_lb = tk.Listbox(frame_select, selectmode="multiple", width=25, height=6, exportselection=False)
        self.select_lb.grid(row=0, column=1)
        btn_use = tk.Button(frame_select, text="Gunakan Kolom Terpilih", command=self.use_selected_columns)
        btn_use.grid(row=0, column=2, padx=5)
        btn_show_all = tk.Button(frame_select, text="Tampilkan Semua Kolom", command=self.show_all_columns)
        btn_show_all.grid(row=1, column=2, padx=5, pady=2)

        # Pilihan kolom pivot
        frame_pivot = tk.Frame(self.main_frame)
        frame_pivot.pack(pady=5)
        tk.Label(frame_pivot, text="Rows:").grid(row=0, column=0)
        tk.Label(frame_pivot, text="Columns:").grid(row=0, column=1)
        tk.Label(frame_pivot, text="Values:").grid(row=0, column=2)
        tk.Label(frame_pivot, text="Aggfunc:").grid(row=0, column=3)

        self.rows_cb = ttk.Combobox(frame_pivot, state="readonly", width=15)
        self.rows_cb.grid(row=1, column=0)
        self.cols_cb = ttk.Combobox(frame_pivot, state="readonly", width=15)
        self.cols_cb.grid(row=1, column=1)
        self.vals_cb = ttk.Combobox(frame_pivot, state="readonly", width=15)
        self.vals_cb.grid(row=1, column=2)
        self.agg_cb = ttk.Combobox(frame_pivot, state="readonly", width=10, values=["sum", "mean", "count", "min", "max"])
        self.agg_cb.grid(row=1, column=3)
        self.agg_cb.set("sum")

        btn_pivot = tk.Button(frame_pivot, text="Pivot", command=self.do_pivot)
        btn_pivot.grid(row=1, column=4, padx=5)

        btn_sort = tk.Button(frame_pivot, text="Sort Descending", command=self.sort_descending)
        btn_sort.grid(row=1, column=5, padx=5)

        btn_reset = tk.Button(self.main_frame, text="Reset", command=self.reset_data)
        btn_reset.pack(pady=5)

        btn_load_ref = tk.Button(self.main_frame, text="Load File Referensi XLOOKUP", command=self.load_xlookup_reference_file)
        btn_load_ref.pack(pady=5)
        btn_xlookup = tk.Button(self.main_frame, text="XLOOKUP dari File Referensi", command=self.xlookup_dialog)
        btn_xlookup.pack(pady=5)

        btn_export = tk.Button(self.main_frame, text="Export to Excel", command=self.export_excel)
        btn_export.pack(pady=5)

        # --- Filter Kolom dan Nilai (Dinamis) ---
        frame_filter = tk.Frame(self.main_frame)
        frame_filter.pack(pady=5)
        tk.Label(frame_filter, text="Filter Kolom:").grid(row=0, column=0)
        self.filter_col_lb = tk.Listbox(frame_filter, selectmode="multiple", width=30, height=5, exportselection=False)
        self.filter_col_lb.grid(row=0, column=1)
        self.filter_col_lb.delete(0, tk.END)
        self.filter_col_lb.bind("<<ListboxSelect>>", self.update_filter_value_entries)
        self.filter_val_frame = tk.Frame(frame_filter)
        self.filter_val_frame.grid(row=1, column=1, columnspan=4, sticky="w")
        self.filter_val_entries = []

        btn_filter = tk.Button(frame_filter, text="Filter", command=self.apply_filter)
        btn_filter.grid(row=0, column=4, padx=5)

        # Tabel hasil
        frame_tree = tk.Frame(self.main_frame)
        frame_tree.pack(fill="both", expand=True)

        # Scrollbar vertikal
        tree_scroll_y = tk.Scrollbar(frame_tree, orient="vertical")
        tree_scroll_y.pack(side="right", fill="y")

        # Scrollbar horizontal
        tree_scroll_x = tk.Scrollbar(frame_tree, orient="horizontal")
        tree_scroll_x.pack(side="bottom", fill="x")

        self.tree = ttk.Treeview(
            frame_tree,
            yscrollcommand=tree_scroll_y.set,
            xscrollcommand=tree_scroll_x.set
        )
        self.tree.pack(fill="both", expand=True)

        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)


    def update_filter_value_entries(self, event=None):
        # Hapus entry lama
        for widget in self.filter_val_frame.winfo_children():
            widget.destroy()
        self.filter_val_entries = []
        selected_indices = self.filter_col_lb.curselection()
        for idx in selected_indices:
            col = self.filter_col_lb.get(idx)
            lbl = tk.Label(self.filter_val_frame, text=f"Nilai untuk '{col}':")
            lbl.pack(side=tk.LEFT, padx=2)
            # Ambil nilai unik dari kolom yang dipilih
            if self.df is not None and col in self.df.columns:
                unique_vals = sorted(self.df[col].dropna().astype(str).unique())
            else:
                unique_vals = []
            # Gunakan Combobox, bukan Entry
            val_cb = ttk.Combobox(self.filter_val_frame, values=unique_vals, width=12, state="readonly")
            val_cb.pack(side=tk.LEFT, padx=2)
            self.filter_val_entries.append((col, val_cb))

    def split_column(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Load Excel file dulu!")
            return
        col = self.split_cb.get()
        if not col:
            messagebox.showwarning("Warning", "Pilih kolom yang ingin di-split!")
            return

        import re
        try:
            def smart_split(val):
                if pd.isna(val):
                    return []
                s = str(val)
                match = re.search(r"\[(.*?)\]", s)
                if match:
                    return match.group(1).split(";")
                return s.split(";")

            new_cols = self.df[col].apply(lambda x: pd.Series(smart_split(x)))
            new_cols.columns = [f"{col}_{i+1}" for i in range(new_cols.shape[1])]
            self.df = self.df.drop(columns=[col]).join(new_cols)
            self.show_data(self.df)
            cols = list(self.df.columns)
            self.rows_cb["values"] = cols
            self.cols_cb["values"] = cols
            self.vals_cb["values"] = cols
            self.split_cb["values"] = cols
            self.select_lb.delete(0, tk.END)
            for col in cols:
                self.select_lb.insert(tk.END, col)
            # Update filter listbox & entry
            self.filter_col_lb.delete(0, tk.END)
            for col in cols:
                self.filter_col_lb.insert(tk.END, col)
            for widget in self.filter_val_frame.winfo_children():
                widget.destroy()
            self.filter_val_entries = []
        except Exception as e:
            messagebox.showerror("Error", f"Split gagal: {e}")

    def export_excel(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Tidak ada data untuk diexport!")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        try:
            with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                # Sheet 1: Data kolom yang sedang ditampilkan
                self.df.to_excel(writer, index=False, sheet_name="Sheet1")
                # Sheet 2: Data hasil pivot (jika ada)
                if self.pivot_df is not None:
                    self.pivot_df.to_excel(writer, index=False, sheet_name="Sheet2")
            messagebox.showinfo("Export", "Data berhasil diexport ke 2 sheet!")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal export: {e}")


    def load_xlookup_reference_file(self):
        file_path = filedialog.askopenfilename(
            title="Pilih file referensi XLOOKUP",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            return
        try:
            import os
            ext = os.path.splitext(file_path)[1].lower()
            if ext == ".xls":
                try:
                    self.ref_xl = pd.ExcelFile(file_path, engine="xlrd")
                except Exception as e1:
                    try:
                        self.ref_xl = pd.ExcelFile(file_path, engine="openpyxl")
                        messagebox.showwarning("Peringatan", "File .xls gagal dibuka dengan xlrd, dicoba dengan openpyxl. Jika data tidak sesuai, pastikan file benar-benar format Excel.")
                    except Exception as e2:
                        messagebox.showerror("Error", f"File .xls gagal dibuka dengan xlrd maupun openpyxl.\nPesan error:\n{e1}\n{e2}")
                        return
            else:
                self.ref_xl = pd.ExcelFile(file_path, engine="openpyxl")
            sheets = self.ref_xl.sheet_names
            # Pilih sheet & header
            sheet_win = tk.Toplevel(self.root)
            sheet_win.title("Pilih Sheet Referensi")
            tk.Label(sheet_win, text="Pilih sheet referensi:").grid(row=0, column=0, padx=10, pady=5)
            sheet_cb = ttk.Combobox(sheet_win, values=sheets, state="readonly")
            sheet_cb.grid(row=0, column=1, padx=10, pady=5)
            sheet_cb.set(sheets[0])

            tk.Label(sheet_win, text="Baris Header (angka, mulai dari 1):").grid(row=1, column=0, padx=10, pady=5)
            header_entry = tk.Entry(sheet_win, width=5)
            header_entry.grid(row=1, column=1, padx=10, pady=5)
            header_entry.insert(0, "1")

            def next_step():
                try:
                    header_row = int(header_entry.get())
                    if header_row < 1:
                        header_row = 1
                except Exception:
                    header_row = 1
                self.ref_df = self.ref_xl.parse(sheet_cb.get(), header=header_row-1)
                sheet_win.destroy()
                messagebox.showinfo("Sukses", "File referensi berhasil dimuat. Silakan lanjutkan XLOOKUP.")

            tk.Button(sheet_win, text="Lanjut", command=next_step).grid(row=2, column=0, columnspan=2, pady=10)
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membaca file referensi: {e}")

    def xlookup_dialog(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Load Excel file dulu!")
            return
        if not hasattr(self, "ref_df") or self.ref_df is None:
            messagebox.showwarning("Warning", "Load file referensi XLOOKUP dulu!")
            return
        self.xlookup_column_selector(self.ref_df)

    def xlookup_column_selector(self, ref_df):
        win = tk.Toplevel(self.root)
        win.title("XLOOKUP - Pilih Kolom")

        tk.Label(win, text="Kolom kunci di data utama:").grid(row=0, column=0, padx=5, pady=5)
        main_key_cb = ttk.Combobox(win, values=list(self.df.columns), state="readonly", width=25)
        main_key_cb.grid(row=0, column=1, padx=5, pady=5)
        main_key_cb.set(self.df.columns[0])

        tk.Label(win, text="Kolom kunci di referensi:").grid(row=1, column=0, padx=5, pady=5)
        ref_key_cb = ttk.Combobox(win, values=list(ref_df.columns), state="readonly", width=25)
        ref_key_cb.grid(row=1, column=1, padx=5, pady=5)
        ref_key_cb.set(ref_df.columns[0])

        tk.Label(win, text="Kolom hasil dari referensi (bisa pilih lebih dari satu):").grid(row=2, column=0, padx=5, pady=5)
        ref_val_lb = tk.Listbox(win, selectmode="multiple", width=25, height=8, exportselection=False)
        for col in ref_df.columns:
            ref_val_lb.insert(tk.END, col)
        ref_val_lb.grid(row=2, column=1, padx=5, pady=5)

        def do_xlookup():
            main_key = main_key_cb.get()
            ref_key = ref_key_cb.get()
            selected_indices = ref_val_lb.curselection()
            if not (main_key and ref_key and selected_indices):
                messagebox.showwarning("Warning", "Semua kolom harus dipilih!")
                return
            try:
                main_series = self.df[main_key].astype(str)
                ref_key_series = ref_df[ref_key].astype(str)
                for idx in selected_indices:
                    ref_val = ref_df.columns[idx]
                    ref_val_series = ref_df[ref_val]
                    ref_dict = pd.Series(ref_val_series.values, index=ref_key_series).to_dict()
                    self.df[f"XLOOKUP_{ref_val}"] = main_series.map(ref_dict)
                self.show_data(self.df)
                messagebox.showinfo("Sukses", "Kolom hasil XLOOKUP berhasil ditambahkan!")
                win.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Gagal XLOOKUP: {e}")

        tk.Button(win, text="Lakukan XLOOKUP", command=do_xlookup).grid(row=3, column=0, columnspan=2, pady=10)


    def load_excel(self):
        # Bisa pilih banyak file (maks 5). Tipe yg diperbolehkan: .txt, .xlsx, .xls, .json
        file_paths = filedialog.askopenfilenames(
            title="Pilih file (max 5)",
            filetypes=[
                ("Excel/txt/json/parquet/csv/feather files", "*.xlsx *.xls *.txt *.json *.parquet *.csv *.feather *.ndjson *.jsonl"),
                ("Excel files", "*.xlsx *.xls"),
                ("Text files", "*.txt"),
                ("CSV files", "*.csv"),
                ("NDJSON files", "*.ndjson *.jsonl"),
                ("Feather files", "*.feather"),
                ("Parquet files", "*.parquet")
            ]
        )
        if not file_paths:
            return

        # Batasi hingga 5 file
        if len(file_paths) > 5:
            messagebox.showwarning("Peringatan", "Pilih maksimal 5 file.")
            return

        import os
        # Baca header baris yang dipilih user (default 1)
        try:
            header_row = int(self.header_entry.get())
            if header_row < 1:
                header_row = 1
        except Exception:
            header_row = 1

        dfs = []
        only_one_excel = False
        only_one_excel_path = None
        try:
            for fp in file_paths:
                ext = os.path.splitext(fp)[1].lower()
                if ext == ".txt":
                    # Deteksi delimiter otomatis (atau ganti sesuai kebutuhan)
                    with open(fp, "r", encoding="utf-8") as f:
                        sample = f.read(2048)
                    if "\t" in sample:
                        delimiter = "\t"
                    elif ";" in sample:
                        delimiter = ";"
                    else:
                        delimiter = ","
                    df_tmp = pd.read_csv(fp, delimiter=delimiter, header=header_row-1)
                    # Tambahkan kolom sumber file
                    try:
                        df_tmp['Sumber File'] = os.path.basename(fp)
                    except Exception:
                        pass
                    dfs.append(df_tmp)
                elif ext in [".xlsx", ".xls"]:
                    # Jika cuma satu file yang dipilih dan itu excel, kita biarkan user memilih sheet
                    if len(file_paths) == 1:
                        only_one_excel = True
                        only_one_excel_path = fp
                        break
                    # Untuk banyak file excel, baca sheet pertama setiap file
                    if ext == ".xls":
                        try:
                            df_tmp = pd.read_excel(fp, engine="xlrd", header=header_row-1)
                        except Exception:
                            df_tmp = pd.read_excel(fp, engine="openpyxl", header=header_row-1)
                    else:
                        df_tmp = pd.read_excel(fp, engine="openpyxl", header=header_row-1)
                    # Tambahkan kolom sumber file
                    try:
                        df_tmp['Sumber File'] = os.path.basename(fp)
                    except Exception:
                        pass
                    dfs.append(df_tmp)
                elif ext == ".csv":
                    df_tmp = pd.read_csv(fp, header=header_row-1)
                    try:
                        df_tmp['Sumber File'] = os.path.basename(fp)
                    except Exception:
                        pass
                    dfs.append(df_tmp)
                elif ext == ".feather":
                    try:
                        df_tmp = pd.read_feather(fp)
                    except Exception as e:
                        import tkinter.messagebox as messagebox
                        messagebox.showerror("Error", f"Gagal membaca feather {fp}: {e}\nPastikan modul 'pyarrow' terinstall (pip install pyarrow).")
                        return
                    try:
                        df_tmp['Sumber File'] = os.path.basename(fp)
                    except Exception:
                        pass
                    dfs.append(df_tmp)
                elif ext in [".ndjson", ".jsonl"]:
                    try:
                        # newline-delimited JSON
                        df_tmp = pd.read_json(fp, lines=True)
                    except Exception:
                        # fallback to regular json read
                        try:
                            df_tmp = pd.read_json(fp)
                        except Exception as e:
                            import tkinter.messagebox as messagebox
                            messagebox.showerror("Error", f"Gagal membaca ndjson/jsonl {fp}: {e}")
                            return
                    try:
                        df_tmp['Sumber File'] = os.path.basename(fp)
                    except Exception:
                        pass
                    dfs.append(df_tmp)
                elif ext == ".parquet":
                    try:
                        # Try to read parquet (requires pyarrow or fastparquet)
                        df_tmp = pd.read_parquet(fp)
                    except Exception as e:
                        # Give helpful message about pyarrow
                        import tkinter.messagebox as messagebox
                        messagebox.showerror("Error", f"Gagal membaca parquet {fp}: {e}\nPastikan modul 'pyarrow' terinstall (pip install pyarrow).")
                        return
                    try:
                        df_tmp['Sumber File'] = os.path.basename(fp)
                    except Exception:
                        pass
                    dfs.append(df_tmp)
                elif ext == ".json":
                    try:
                        df_tmp = pd.read_json(fp)
                    except Exception:
                        df_tmp = pd.read_json(fp, orient='records')
                    try:
                        df_tmp['Sumber File'] = os.path.basename(fp)
                    except Exception:
                        pass
                    dfs.append(df_tmp)
                else:
                    messagebox.showerror("Error", f"Format file tidak didukung: {fp}")
                    return

            if only_one_excel and only_one_excel_path:
                # Perlakukan seperti sebelumnya: buka ExcelFile dan sheet selection
                fp = only_one_excel_path
                ext = os.path.splitext(fp)[1].lower()
                try:
                    if ext == ".xls":
                        try:
                            self.excel_file = pd.ExcelFile(fp, engine="xlrd")
                        except Exception as e1:
                            try:
                                self.excel_file = pd.ExcelFile(fp, engine="openpyxl")
                                messagebox.showwarning("Peringatan", "File .xls gagal dibuka dengan xlrd, dicoba dengan openpyxl. Jika data tidak sesuai, pastikan file benar-benar format Excel.")
                            except Exception as e2:
                                messagebox.showerror("Error", f"File .xls gagal dibuka dengan xlrd maupun openpyxl.\nPesan error:\n{e1}\n{e2}")
                                return
                    else:
                        self.excel_file = pd.ExcelFile(fp, engine="openpyxl")
                except ImportError:
                    messagebox.showerror("Error", "File .xls membutuhkan modul xlrd. Silakan install dengan 'pip install xlrd'.")
                    return
                # Simpan path untuk ditandai sebagai sumber saat sheet dipilih
                self._last_loaded_excel_path = fp
                self.sheet_cb["values"] = self.excel_file.sheet_names
                if self.excel_file.sheet_names:
                    self.sheet_cb.set(self.excel_file.sheet_names[0])
                    self.on_sheet_selected()  # Otomatis load sheet pertama
                return

            # Gabungkan semua DataFrame yang terbaca
            if not dfs:
                messagebox.showwarning("Warning", "Tidak ada data yang berhasil dibaca dari file yang dipilih.")
                return
            self.df = pd.concat(dfs, ignore_index=True)
            self.original_df = self.df.copy()
            self.pivot_df = None
            self.show_data(self.df)
            cols = list(self.df.columns)
            self.rows_cb["values"] = cols
            self.cols_cb["values"] = cols
            self.vals_cb["values"] = cols
            self.split_cb["values"] = cols
            self.rows_cb.set("")
            self.cols_cb.set("")
            self.vals_cb.set("")
            self.split_cb.set("")
            self.select_lb.delete(0, tk.END)
            for col in cols:
                self.select_lb.insert(tk.END, col)
            # Update filter listbox & entry
            self.filter_col_lb.delete(0, tk.END)
            for col in cols:
                self.filter_col_lb.insert(tk.END, col)
            for widget in self.filter_val_frame.winfo_children():
                widget.destroy()
            self.filter_val_entries = []

        except Exception as e:
            messagebox.showerror("Error", f"Gagal membaca file: {e}")

    def show_all_columns(self):
        if hasattr(self, "original_df") and self.original_df is not None:
            self.df = self.original_df.copy()
            self.show_data(self.df)
            # Update combobox dan listbox kolom
            cols = list(self.df.columns)
            self.rows_cb["values"] = cols
            self.cols_cb["values"] = cols
            self.vals_cb["values"] = cols
            self.split_cb["values"] = cols
            self.rows_cb.set("")
            self.cols_cb.set("")
            self.vals_cb.set("")
            self.split_cb.set("")
            self.select_lb.delete(0, tk.END)
            for col in cols:
                self.select_lb.insert(tk.END, col)
            # Update filter listbox & entry
            self.filter_col_lb.delete(0, tk.END)
            for col in cols:
                self.filter_col_lb.insert(tk.END, col)
            for widget in self.filter_val_frame.winfo_children():
                widget.destroy()
            self.filter_val_entries = []

    def apply_filter(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Load Excel file dulu!")
            return
        if not self.filter_val_entries:
            messagebox.showwarning("Warning", "Pilih minimal 1 kolom untuk filter!")
            return
        try:
            filtered_df = self.df
            for col, ent in self.filter_val_entries:
                val = ent.get().strip()
                if val == "":
                    continue
                filtered_df = filtered_df[filtered_df[col].astype(str) == val]
            self.show_data(filtered_df)
        except Exception as e:
            messagebox.showerror("Error", f"Filter gagal: {e}")

    def on_sheet_selected(self, event=None):
        sheet_name = self.sheet_cb.get()
        if not sheet_name:
            return
        try:
            header_row = int(self.header_entry.get())
            if header_row < 1:
                header_row = 1
        except Exception:
            header_row = 1
        self.df = self.excel_file.parse(sheet_name, header=header_row-1)
        self.original_df = self.df.copy()
        self.pivot_df = None
        self.show_data(self.df)
        cols = list(self.df.columns)
        self.rows_cb["values"] = cols
        self.cols_cb["values"] = cols
        self.vals_cb["values"] = cols
        self.split_cb["values"] = cols
        self.rows_cb.set("")
        self.cols_cb.set("")
        self.vals_cb.set("")
        self.split_cb.set("")
        self.select_lb.delete(0, tk.END)
        for col in cols:
            self.select_lb.insert(tk.END, col)
        # Update filter listbox & entry
        self.filter_col_lb.delete(0, tk.END)
        for col in cols:
            self.filter_col_lb.insert(tk.END, col)
        for widget in self.filter_val_frame.winfo_children():
            widget.destroy()
        self.filter_val_entries = []

    def use_selected_columns(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Load Excel file dulu!")
            return
        selected_indices = self.select_lb.curselection()
        if not selected_indices:
            messagebox.showwarning("Warning", "Pilih kolom yang ingin ditampilkan!")
            return
        selected_cols = [self.select_lb.get(i) for i in selected_indices]
        self.df = self.df[selected_cols]
        self.show_data(self.df)
        # Update combobox
        cols = list(self.df.columns)
        self.rows_cb["values"] = cols
        self.cols_cb["values"] = cols
        self.vals_cb["values"] = cols
        self.split_cb["values"] = cols
        self.rows_cb.set("")
        self.cols_cb.set("")
        self.vals_cb.set("")
        self.split_cb.set("")
        self.select_lb.delete(0, tk.END)
        for col in cols:
            self.select_lb.insert(tk.END, col)
        # Update filter listbox & entry
        self.filter_col_lb.delete(0, tk.END)
        for col in cols:
            self.filter_col_lb.insert(tk.END, col)
        for widget in self.filter_val_frame.winfo_children():
            widget.destroy()
        self.filter_val_entries = []

    def show_data(self, df):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        self.tree["show"] = "headings"
        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        for _, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def do_pivot(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Load Excel file dulu!")
            return
        rows = self.rows_cb.get()
        cols = self.cols_cb.get()
        vals = self.vals_cb.get()
        agg = self.agg_cb.get()
        if not (rows and cols and vals):
            messagebox.showwarning("Warning", "Pilih semua field pivot!")
            return
        try:
            pivot = pd.pivot_table(self.df, index=rows, columns=cols, values=vals, aggfunc=agg, fill_value=0)
            pivot = pivot.reset_index()
            self.pivot_df = pivot
            self.show_data(pivot)
            # Reset combobox pivot
            self.rows_cb.set("")
            self.cols_cb.set("")
            self.vals_cb.set("")
            self.split_cb.set("")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal pivot: {e}")

    def sort_descending(self):
        if self.pivot_df is None:
            messagebox.showwarning("Warning", "Lakukan pivot dulu!")
            return
        numeric_cols = self.pivot_df.select_dtypes(include='number').columns
        if len(numeric_cols) == 0:
            messagebox.showwarning("Warning", "Tidak ada kolom numerik untuk diurutkan!")
            return
        sort_col = None
        if len(numeric_cols) == 1:
            sort_col = numeric_cols[0]
        else:
            sort_col = self.vals_cb.get()
            if sort_col not in numeric_cols:
                sort_col = numeric_cols[0]
        sorted_df = self.pivot_df.sort_values(by=sort_col, ascending=False)
        self.pivot_df = sorted_df
        self.show_data(sorted_df)

    def reset_data(self):
        if self.df is not None:
            self.show_data(self.df)
            cols = list(self.df.columns)
            self.rows_cb["values"] = cols
            self.cols_cb["values"] = cols
            self.vals_cb["values"] = cols
            self.split_cb["values"] = cols
            self.rows_cb.set("")
            self.cols_cb.set("")
            self.vals_cb.set("")
            self.split_cb.set("")
            self.select_lb.delete(0, tk.END)
            for col in cols:
                self.select_lb.insert(tk.END, col)
            # Update filter listbox & entry
            self.filter_col_lb.delete(0, tk.END)
            for col in cols:
                self.filter_col_lb.insert(tk.END, col)
            for widget in self.filter_val_frame.winfo_children():
                widget.destroy()
            self.filter_val_entries = []


    def ask_ai(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Load Excel file dulu!")
            return

        # Prompt input
        prompt_win = tk.Toplevel(self.root)
        prompt_win.title("Tanya AI")
        tk.Label(prompt_win, text="Masukkan pertanyaan atau instruksi untuk AI:").pack(padx=10, pady=5)
        prompt_entry = scrolledtext.ScrolledText(prompt_win, height=5, width=60)
        prompt_entry.pack(padx=10, pady=5)
        prompt_entry.focus_set()

        def send_to_ai():
            prompt = prompt_entry.get("1.0", tk.END).strip()
            if not prompt:
                messagebox.showwarning("Warning", "Prompt tidak boleh kosong!")
                return
            prompt_win.destroy()
            # Kirim ke AI
            try:
                sample_data = self.df.head(10).to_csv(index=False)
                full_prompt = (
                    f"Saya punya data seperti berikut (hanya 10 baris pertama):\n\n{sample_data}\n\n"
                    f"Tolong bantu dengan instruksi berikut:\n{prompt}\n"
                )
                headers = {
                    "Authorization": f"Bearer {OPENROUTER_API_KEY}",
                    "Content-Type": "application/json"
                }
                data = {
                    "model": "openai/gpt-3.5-turbo",
                    "messages": [{"role": "user", "content": full_prompt}],
                    "temperature": 0.3
                }
                import json
                response = requests.post(
                    "https://openrouter.ai/api/v1/chat/completions",
                    headers=headers,
                    data=json.dumps(data),
                    timeout=60
                )
                if response.status_code == 200:
                    result = response.json()
                    ai_reply = result["choices"][0]["message"]["content"]
                    self.ai_result_box.delete("1.0", tk.END)
                    self.ai_result_box.insert(tk.END, ai_reply)
                else:
                    self.ai_result_box.delete("1.0", tk.END)
                    self.ai_result_box.insert(tk.END, f"Error: {response.text}")
            except Exception as e:
                self.ai_result_box.delete("1.0", tk.END)
                self.ai_result_box.insert(tk.END, f"Error: {e}")

        tk.Button(prompt_win, text="Kirim ke AI", command=send_to_ai).pack(pady=5)


def read_file(filepath):
    import os
    ext = os.path.splitext(filepath)[1].lower()
    if ext == '.csv':
        return pd.read_csv(filepath)
    elif ext == '.feather':
        try:
            return pd.read_feather(filepath)
        except Exception as e:
            import tkinter.messagebox as messagebox
            messagebox.showerror("Error", f"Gagal membaca feather {filepath}: {e}\nPastikan modul 'pyarrow' terinstall (pip install pyarrow).")
            raise
    elif ext in ['.ndjson', '.jsonl']:
        try:
            return pd.read_json(filepath, lines=True)
        except Exception:
            try:
                return pd.read_json(filepath)
            except Exception as e:
                import tkinter.messagebox as messagebox
                messagebox.showerror("Error", f"Gagal membaca ndjson/jsonl {filepath}: {e}")
                raise
    elif ext == '.json':
        # Try to read JSON; try default, then fallback to records orientation
        try:
            return pd.read_json(filepath)
        except Exception:
            return pd.read_json(filepath, orient='records')
    elif ext in ['.xlsx', '.xls']:
        if ext == '.xls':
            try:
                return pd.read_excel(filepath, engine='xlrd')
            except Exception as e1:
                try:
                    result = pd.read_excel(filepath, engine='openpyxl')
                    import tkinter.messagebox as messagebox
                    messagebox.showwarning("Peringatan", "File .xls gagal dibuka dengan xlrd, dicoba dengan openpyxl. Jika data tidak sesuai, pastikan file benar-benar format Excel.")
                    return result
                except Exception as e2:
                    raise ValueError(f"File .xls gagal dibuka dengan xlrd maupun openpyxl.\nPesan error:\n{e1}\n{e2}")
        else:
            return pd.read_excel(filepath, engine='openpyxl')
    elif ext == '.parquet':
        try:
            return pd.read_parquet(filepath)
        except Exception as e:
            import tkinter.messagebox as messagebox
            messagebox.showerror("Error", f"Gagal membaca parquet {filepath}: {e}\nPastikan modul 'pyarrow' terinstall (pip install pyarrow).")
            raise
    else:
        raise ValueError('Format file tidak didukung: ' + ext)

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("1000x700")  # Ukuran awal
    root.minsize(800, 500)     # Ukuran minimal

    # Tampilkan dialog pilihan
    show_ai = messagebox.askyesno("Fitur AI", "Tampilkan fitur AI (Tanya AI tentang Data)?")

    app = PowerQueryPivot(root, show_ai=show_ai)
    root.mainloop()