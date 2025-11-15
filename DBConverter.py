"""
merge_to_json.py
A small utility with GUI and CLI to merge multiple .txt/.csv/.xls/.xlsx files into one JSON file.

Usage (GUI):
    python merge_to_json.py

Usage (CLI):
    python merge_to_json.py --cli file1.xlsx file2.txt --header 1 --include-source --include-sheet -o out.json

Output: JSON array of records (utf-8, pretty printed)

Dependencies: pandas, openpyxl, xlrd
"""

import os
import sys
import argparse
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd


def read_file(filepath, header_row=1):
    """Read a file into a pandas DataFrame. For Excel files all sheets are returned as dict of DataFrames.
    For txt/csv returns a single DataFrame.
    header_row is 1-based index of header line.
    """
    ext = os.path.splitext(filepath)[1].lower()
    if ext in [".xlsx", ".xls"]:
        # read all sheets
        try:
            sheets = pd.read_excel(filepath, sheet_name=None, engine="openpyxl" if ext == ".xlsx" else None, header=header_row-1)
        except Exception:
            # try xlrd for xls
            sheets = pd.read_excel(filepath, sheet_name=None, header=header_row-1)
        return sheets  # dict: sheet_name -> df
    else:
        # try csv/txt delimiter detection
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            sample = f.read(4096)
        if "\t" in sample:
            delimiter = "\t"
        elif ";" in sample:
            delimiter = ";"
        else:
            delimiter = ","
        df = pd.read_csv(filepath, delimiter=delimiter, header=header_row-1, encoding="utf-8", engine="python")
        return df


def merge_files(filepaths, header_row=1, include_source=True, include_sheet=True, sheet_selection=None, progress_callback=None):
    """Merge multiple files and return a single DataFrame.

    sheet_selection: optional dict {filepath: [sheet1, sheet2, ...]} to limit sheets per Excel file.
    progress_callback: optional callable(current, total) to report progress.
    """
    dfs = []

    # Compute total steps for progress: for non-excel -> 1, for excel -> number of sheets to read
    total_steps = 0
    for fp in filepaths:
        ext = os.path.splitext(fp)[1].lower()
        if ext in [".xlsx", ".xls"]:
            try:
                xl = pd.ExcelFile(fp, engine="openpyxl" if fp.lower().endswith('.xlsx') else None)
                sheets = xl.sheet_names
            except Exception:
                try:
                    xl = pd.ExcelFile(fp)
                    sheets = xl.sheet_names
                except Exception:
                    sheets = []
            if sheet_selection and fp in sheet_selection and sheet_selection[fp]:
                total_steps += len(sheet_selection[fp])
            else:
                total_steps += max(1, len(sheets))
        else:
            total_steps += 1

    step = 0
    for fp in filepaths:
        ext = os.path.splitext(fp)[1].lower()
        try:
            if ext in [".xlsx", ".xls"]:
                # Determine which sheets to read
                sel = None
                if sheet_selection and fp in sheet_selection and sheet_selection[fp]:
                    sel = sheet_selection[fp]
                if sel is None:
                    # read all sheets
                    res = pd.read_excel(fp, sheet_name=None, header=header_row-1)
                else:
                    res = pd.read_excel(fp, sheet_name=sel, header=header_row-1)
                # res can be dict or DataFrame (if single sheet requested)
                if isinstance(res, dict):
                    items = res.items()
                else:
                    # single sheet -> try to get name from selection or use generic
                    name = sel[0] if isinstance(sel, (list, tuple)) and sel else os.path.splitext(os.path.basename(fp))[0]
                    items = [(name, res)]
                for sheet_name, df in items:
                    df = df.copy()
                    if include_source:
                        df["Sumber File"] = os.path.basename(fp)
                    if include_sheet:
                        df["Sumber Sheet"] = sheet_name
                    dfs.append(df)
                    step += 1
                    if progress_callback:
                        progress_callback(step, total_steps)
            else:
                res = read_file(fp, header_row=header_row)
                df = res.copy()
                if include_source:
                    df["Sumber File"] = os.path.basename(fp)
                if include_sheet:
                    df["Sumber Sheet"] = ""
                dfs.append(df)
                step += 1
                if progress_callback:
                    progress_callback(step, total_steps)
        except Exception as e:
            raise RuntimeError(f"Gagal membaca {fp}: {e}")

    if not dfs:
        return pd.DataFrame()
    # Align columns: union of all columns, fill NaN where missing
    all_cols = []
    for d in dfs:
        for c in d.columns:
            if c not in all_cols:
                all_cols.append(c)
    normalized = []
    for d in dfs:
        missing = [c for c in all_cols if c not in d.columns]
        if missing:
            for m in missing:
                d[m] = pd.NA
        normalized.append(d[all_cols])
    result = pd.concat(normalized, ignore_index=True)
    return result


def save_to_json(df, outpath):
    """Save DataFrame to JSON file (list of records)."""
    try:
        import json
        if outpath.lower().endswith('.ndjson'):
            # write newline-delimited JSON
            with open(outpath, "w", encoding="utf-8") as f:
                for rec in df.where(pd.notnull(df), None).to_dict(orient="records"):
                    f.write(json.dumps(rec, ensure_ascii=False))
                    f.write('\n')
        else:
            records = df.where(pd.notnull(df), None).to_dict(orient="records")
            with open(outpath, "w", encoding="utf-8") as f:
                json.dump(records, f, ensure_ascii=False, indent=2)
    except Exception as e:
        raise RuntimeError(f"Gagal menyimpan JSON: {e}")


def save_dataframe(df, outpath, fmt='json'):
    """Save DataFrame to various formats.

    fmt values: 'json', 'ndjson', 'parquet-snappy', 'parquet-gzip', 'feather', 'csv-gzip'
    """
    try:
        f = fmt.lower()
        if f == 'json' or f == 'json-pretty':
            save_to_json(df, outpath)
        elif f == 'ndjson' or outpath.lower().endswith('.ndjson'):
            # NDJSON
            save_to_json(df, outpath)
        elif f in ('parquet-snappy', 'parquet-gzip', 'parquet'):
            # write parquet with pyarrow
            try:
                import pyarrow  # noqa: F401
            except Exception:
                raise RuntimeError('pyarrow required to write parquet. Install with pip install pyarrow')
            comp = 'snappy' if 'snappy' in f else ('gzip' if 'gzip' in f else 'snappy')
            df.to_parquet(outpath, index=False, compression=comp)
        elif f == 'feather':
            try:
                import pyarrow  # feather uses pyarrow
            except Exception:
                raise RuntimeError('pyarrow required to write feather. Install with pip install pyarrow')
            df.reset_index(drop=True).to_feather(outpath)
        elif f == 'csv-gzip' or (outpath.lower().endswith('.csv.gz') or outpath.lower().endswith('.csv')):
            # write gzipped csv
            if not outpath.lower().endswith('.gz'):
                # if user provided .csv, add .gz for csv-gzip format
                if f == 'csv-gzip':
                    outpath = outpath + '.gz'
            df.to_csv(outpath, index=False, compression='gzip', encoding='utf-8')
        else:
            # fallback to JSON
            save_to_json(df, outpath)
    except Exception as e:
        raise RuntimeError(f"Gagal menyimpan file: {e}")


# --------- GUI ----------

def configure_sheets_dialog(parent, filepaths_getter, sheet_selection):
    """Opens a dialog to configure which sheets to read per Excel file.
    filepaths_getter: callable returning current list of filepaths
    sheet_selection: dict to be mutated with selections
    """
    fps = [p for p in filepaths_getter()]
    if not fps:
        messagebox.showwarning("Peringatan", "Belum ada file yang dipilih.")
        return

    dlg = tk.Toplevel(parent)
    dlg.title("Konfigurasi Sheet per-file")
    dlg.geometry("600x400")
    frame = ttk.Frame(dlg, padding=8)
    frame.pack(fill=tk.BOTH, expand=True)

    canv = tk.Canvas(frame)
    vsb = ttk.Scrollbar(frame, orient='vertical', command=canv.yview)
    inner = ttk.Frame(canv)
    inner_id = canv.create_window((0,0), window=inner, anchor='nw')
    canv.configure(yscrollcommand=vsb.set)
    canv.pack(side='left', fill='both', expand=True)
    vsb.pack(side='right', fill='y')

    def on_configure(e=None):
        canv.configure(scrollregion=canv.bbox('all'))
    inner.bind('<Configure>', on_configure)

    listboxes = {}

    for fp in fps:
        ext = os.path.splitext(fp)[1].lower()
        lbl = ttk.Label(inner, text=os.path.basename(fp))
        lbl.pack(anchor='w', pady=(8,0))
        if ext in ['.xlsx', '.xls']:
            try:
                xl = pd.ExcelFile(fp)
                sheets = xl.sheet_names
            except Exception:
                sheets = []
            lb = tk.Listbox(inner, selectmode='multiple', height=min(8, max(3, len(sheets))))
            for s in sheets:
                lb.insert(tk.END, s)
            lb.pack(fill='x')
            # preselect previous
            prev = sheet_selection.get(fp)
            if prev:
                for i, s in enumerate(sheets):
                    if s in prev:
                        lb.selection_set(i)
            listboxes[fp] = (lb, sheets)
            btn_all = ttk.Button(inner, text='Pilih Semua', command=lambda l=lb: (l.selection_set(0, tk.END)))
            btn_all.pack(anchor='e', pady=(2,4))
        else:
            ttk.Label(inner, text='(bukan Excel)').pack(anchor='w')

    def do_save():
        for fp, val in listboxes.items():
            lb, sheets = val
            sel = [sheets[i] for i in lb.curselection()]
            if sel:
                sheet_selection[fp] = sel
            else:
                # empty means all sheets
                sheet_selection.pop(fp, None)
        dlg.destroy()

    btn_frame = ttk.Frame(dlg)
    btn_frame.pack(fill='x', pady=6)
    ttk.Button(btn_frame, text='Simpan', command=do_save).pack(side='left', padx=6)
    ttk.Button(btn_frame, text='Batal', command=dlg.destroy).pack(side='left')


def run_gui():
    root = tk.Tk()
    root.title("Merge Files -> JSON")
    root.geometry("700x420")

    files_var = tk.StringVar(value="")
    # sheet_selection: dict filepath -> list of sheet names (or empty list => all)
    sheet_selection = {}
    format_var = tk.StringVar(value="json")

    def pick_files():
        paths = filedialog.askopenfilenames(title="Pilih file (txt/csv/xls/xlsx)", filetypes=[("All supported", "*.txt *.csv *.xlsx *.xls"), ("Excel files", "*.xlsx *.xls"), ("Text files", "*.txt *.csv")])
        if paths:
            # limit to 20 for safety
            files_var.set('\n'.join(paths))

    def filepaths_getter():
        raw = files_var.get().strip()
        return raw.split('\n') if raw else []

    def do_merge():
        raw = files_var.get().strip()
        if not raw:
            messagebox.showwarning("Peringatan", "Pilih minimal 1 file terlebih dahulu.")
            return
        filepaths = raw.split('\n')
        if len(filepaths) > 20:
            messagebox.showwarning("Peringatan", "Pilih maksimal 20 file pada GUI (gunakan CLI untuk lebih).")
            return
        try:
            header_row = int(header_entry.get())
            if header_row < 1:
                header_row = 1
        except Exception:
            header_row = 1
        include_source = source_var.get()
        include_sheet = sheet_var.get()
        try:
            # Disable controls while merging
            btn_merge.config(state='disabled')
            btn_pick.config(state='disabled')
            btn_config.config(state='disabled')
            root.update_idletasks()

            # progress bar reset
            progress_bar['maximum'] = 100
            progress_bar['value'] = 0

            def progress_cb(current, total):
                try:
                    val = int((current / total) * 100)
                except Exception:
                    val = 0
                progress_bar['value'] = val
                root.update_idletasks()

            df = merge_files(filepaths, header_row=header_row, include_source=include_source, include_sheet=include_sheet, sheet_selection=sheet_selection, progress_callback=progress_cb)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            # re-enable
            btn_merge.config(state='normal')
            btn_pick.config(state='normal')
            btn_config.config(state='normal')
            return
        if df.empty:
            messagebox.showinfo("Info", "Hasil penggabungan kosong.")
            btn_merge.config(state='normal')
            btn_pick.config(state='normal')
            btn_config.config(state='normal')
            return
        # Ask for output file based on selected format
        chosen_fmt = format_var.get().lower()
        if chosen_fmt == 'ndjson':
            filetypes = [("NDJSON", "*.ndjson"), ("JSON", "*.json"), ("All files", "*")]
            default_ext = ".ndjson"
            init_name = "merged.ndjson"
        elif chosen_fmt.startswith('parquet'):
            filetypes = [("Parquet files", "*.parquet"), ("All files", "*")]
            default_ext = ".parquet"
            init_name = "merged.parquet"
        elif chosen_fmt == 'feather':
            filetypes = [("Feather files", "*.feather"), ("All files", "*")]
            default_ext = ".feather"
            init_name = "merged.feather"
        elif chosen_fmt == 'csv-gzip':
            filetypes = [("CSV GZIP", "*.csv.gz"), ("CSV", "*.csv"), ("All files", "*")]
            default_ext = ".csv.gz"
            init_name = "merged.csv.gz"
        else:
            filetypes = [("JSON files", "*.json *.ndjson"), ("All files", "*")]
            default_ext = ".json"
            init_name = "merged.json"

        out = filedialog.asksaveasfilename(defaultextension=default_ext, filetypes=filetypes, initialfile=init_name)
        if not out:
            # re-enable
            btn_merge.config(state='normal')
            btn_pick.config(state='normal')
            btn_config.config(state='normal')
            return
        try:
            save_dataframe(df, out, fmt=chosen_fmt)
            messagebox.showinfo("Sukses", f"File tersimpan: {out}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            btn_merge.config(state='normal')
            btn_pick.config(state='normal')
            btn_config.config(state='normal')
            progress_bar['value'] = 0

    frm = ttk.Frame(root, padding=10)
    frm.pack(fill=tk.BOTH, expand=True)

    lbl = ttk.Label(frm, text="Files:")
    lbl.grid(row=0, column=0, sticky=tk.W)
    txt = tk.Text(frm, height=8, width=80)
    txt.grid(row=1, column=0, columnspan=4, pady=4)

    def refresh_text():
        txt.delete("1.0", tk.END)
        txt.insert(tk.END, files_var.get())

    btn_pick = ttk.Button(frm, text="Pilih Files...", command=lambda: [pick_files(), refresh_text()])
    btn_pick.grid(row=2, column=0, sticky=tk.W)

    btn_config = ttk.Button(frm, text="Konfigurasi Sheet per-file...", command=lambda: configure_sheets_dialog(root, filepaths_getter, sheet_selection))
    btn_config.grid(row=2, column=1, sticky=tk.W, padx=6)

    def filepaths_getter():
        raw = files_var.get().strip()
        return raw.split('\n') if raw else []

    ttk.Label(frm, text="Header row (1-based):").grid(row=3, column=0, sticky=tk.W, pady=6)
    header_entry = ttk.Entry(frm, width=6)
    header_entry.insert(0, "1")
    header_entry.grid(row=3, column=1, sticky=tk.W)

    source_var = tk.BooleanVar(value=True)
    sheet_var = tk.BooleanVar(value=True)
    ttk.Checkbutton(frm, text="Tambah kolom Sumber File", variable=source_var).grid(row=4, column=0, sticky=tk.W)
    ttk.Checkbutton(frm, text="Tambah kolom Sumber Sheet (Excel)", variable=sheet_var).grid(row=4, column=1, sticky=tk.W)
    ttk.Label(frm, text="Format output:").grid(row=4, column=2, sticky=tk.W)
    fmt_combo = ttk.Combobox(frm, textvariable=format_var, values=["json", "ndjson", "parquet-snappy", "parquet-gzip", "feather", "csv-gzip"], width=18)
    fmt_combo.grid(row=4, column=3, sticky=tk.W)

    btn_merge = ttk.Button(frm, text="Merge -> Save JSON", command=do_merge)
    btn_merge.grid(row=5, column=0, pady=12, sticky=tk.W)

    # Progress bar
    progress_bar = ttk.Progressbar(frm, orient='horizontal', length=420, mode='determinate')
    progress_bar.grid(row=5, column=1, columnspan=3, padx=6)

    ttk.Label(frm, text="(GUI) Pilih files lalu klik Merge -> Save").grid(row=6, column=0, columnspan=3, sticky=tk.W)

    root.mainloop()


# --------- CLI ----------

def run_cli(argv):
    p = argparse.ArgumentParser(description="Merge files (.txt/.csv/.xls/.xlsx) to a single JSON file")
    p.add_argument('files', nargs='+', help='Input files')
    p.add_argument('-H', '--header', type=int, default=1, help='Header row (1-based)')
    p.add_argument('--no-source', dest='source', action='store_false', help='Do not include Sumber File column')
    p.add_argument('--no-sheet', dest='sheet', action='store_false', help='Do not include Sumber Sheet column')
    p.add_argument('-o', '--output', default='merged.json')
    p.add_argument('--ndjson', action='store_true', help='Save output as NDJSON (newline-delimited)')
    p.add_argument('--format', choices=['json', 'ndjson', 'parquet-snappy', 'parquet-gzip', 'feather', 'csv-gzip'], default='json', help='Output format')
    p.add_argument('--sheet-selection', action='append', help='Specify sheets per file: "path/to/file.xlsx:Sheet1,Sheet2". Repeatable.')
    args = p.parse_args(argv)
    try:
        # parse sheet selections
        sheet_sel = {}
        if getattr(args, 'sheet_selection', None):
            for item in args.sheet_selection:
                if ':' in item:
                    fp, sheets = item.split(':', 1)
                    fp = fp.strip()
                    sheet_list = [s.strip() for s in sheets.split(',') if s.strip()]
                    if sheet_list:
                        sheet_sel[fp] = sheet_list
        outpath = args.output
        # determine chosen format: CLI --ndjson has precedence for backwards compatibility
        chosen_fmt = args.format
        if args.ndjson:
            chosen_fmt = 'ndjson'
        # normalize output extension for some formats
        if chosen_fmt.startswith('parquet') and not outpath.lower().endswith('.parquet'):
            outpath = os.path.splitext(outpath)[0] + '.parquet'
        if chosen_fmt == 'feather' and not outpath.lower().endswith('.feather'):
            outpath = os.path.splitext(outpath)[0] + '.feather'
        if chosen_fmt == 'ndjson' and not outpath.lower().endswith('.ndjson'):
            outpath = os.path.splitext(outpath)[0] + '.ndjson'
        if chosen_fmt == 'csv-gzip' and not (outpath.lower().endswith('.csv.gz') or outpath.lower().endswith('.csv')):
            outpath = os.path.splitext(outpath)[0] + '.csv.gz'

        df = merge_files(args.files, header_row=args.header, include_source=args.source, include_sheet=args.sheet, sheet_selection=sheet_sel)
        if df.empty:
            print('Hasil penggabungan kosong')
            return 1
        save_dataframe(df, outpath, fmt=chosen_fmt)
        print(f'Saved to {outpath} ({chosen_fmt})')
        return 0
    except Exception as e:
        print('Error:', e)
        return 2


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == '--cli':
        sys.exit(run_cli(sys.argv[2:]))
    else:
        run_gui()
