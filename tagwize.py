# final_smart_matcher_v5.py
# FINAL VERSION: Handles all industrial tag formats including P2106B, P-5401BX, FAN-5401X, BA2801, etc.
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


def extract_tag(text):
    """
    Extracts the main device tag from any text, even with pipes and noise.
    Handles:
        'EL-P-5401BX | AST-EL-P-5401BX | Hot Oil' → 'P-5401BX'
        'EL-P-5303C | AST-EL-P-5303C | BOILER FEED WATER' → 'P-5303C'
        'EL-P2106B | AST-EL-P2106B | Monomer Storage...' → 'P-2106B'
        'EL-P2902B | AST-EL-P2902B | Utility' → 'P-2902B'

    Steps:
    1. Split by | and take first meaningful part
    2. Remove system prefixes: EL-, AST-, INSTR-
    3. Find pattern: Letters + Digits + Optional Letters (e.g., P2106B, P-5401BX)
    4. Insert dash between prefix and number if missing
    5. Return standardized format
    """
    if not isinstance(text, str) or not text.strip():
        return ""

    # Step 1: Split by pipe and get first non-empty, likely tagged part
    parts = re.split(r'[|]+', text)
    for part in parts:
        candidate = part.strip()
        # Skip obvious non-tag lines
        if len(candidate) < 4:
            continue
        if re.search(r'\b(?:Section|SYSTEM|UNIT|Utility|dispatcher)\b', candidate, re.I):
            continue
        if re.fullmatch(r'[A-Za-z\s]+', candidate):  # Only letters and spaces → probably label
            continue
        text = candidate
        break
    else:
        return ""  # No valid part found

    # Step 2: Uppercase and clean
    text = text.upper()

    # Step 3: Remove system-level prefixes (not functional)
    text = re.sub(r'\bEL[-_]', '', text)
    text = re.sub(r'\bAST[-_]', '', text)
    text = re.sub(r'\bINSTR[-_]', '', text)
    text = re.sub(r'\bFIC[-_]', '', text)
    text = re.sub(r'\bTIC[-_]', '', text)

    # Step 4: Look for core pattern: [Letters][optional dash][Digits][Optional Suffix Letters]
    # Examples: P-5401BX, P5303C, P2106B, FAN-5401X, BA2801
    match = re.search(r'\b([A-Z]{1,4})[-_]?\s*(\d{3,6})\s*([A-Z]*)\b', text)
    if match:
        prefix = match.group(1).strip()
        number = match.group(2).strip()
        suffix = match.group(3).strip()
        return f"{prefix}-{number}{suffix}" if suffix else f"{prefix}-{number}"

    # Fallback 1: Match without word boundary (if inside longer text)
    loose = re.search(r'([A-Z]{1,4})[-_]?\s*(\d{3,6})\s*([A-Z]*)', text)
    if loose:
        prefix = loose.group(1)
        number = loose.group(2)
        suffix = loose.group(3)
        return f"{prefix}-{number}{suffix}" if suffix else f"{prefix}-{number}"

    # Fallback 2: Any alphanumeric sequence starting with letters and ending with digits/letters
    generic = re.search(r'[A-Z]{1,4}\d+[A-Z]*', text.replace(' ', ''))
    if generic:
        raw = generic.group(0)
        # Insert dash before first digit
        fixed = re.sub(r'^([A-Z]+)(\d+)', r'\1-\2', raw)
        return fixed

    return ""


def classify_from_text(text):
    """
    Classify device type based on keywords in description.
    Returns: 'Motor', 'Valve', 'Transmitter', or None
    """
    if not isinstance(text, str):
        return None
    text = text.lower()

    motor_keywords = ['motor', 'pump', 'fan', 'compressor', 'el-', 'electro', 'rotary', 'air compressor', 'blower', 'boiler feed water', 'hot oil']
    valve_keywords = ['valve', 'xv-', 'pcv-', 'fcv-', 'lcv-', 'hv-', 'bv-', 'pv-', 'lv-', 'fv-', 'on-off', 'regulating', 'butterfly', 'ball', 'control valve', 'shutoff']
    transmitter_keywords = ['pt-', 'tt-', 'lt-', 'ft-', 'pi-', 'ti-', 'li-', 'fi-', 'psv-', 'prv-', 'transmitter', 'gauge', 'indicator', 'element', 'switch', 'temperature', 'pressure', 'level', 'flow']

    score = {"Motor": 0, "Valve": 0, "Transmitter": 0}

    for word in motor_keywords:
        if word in text:
            score["Motor"] += 1
    for word in valve_keywords:
        if word in text:
            score["Valve"] += 1
    for word in transmitter_keywords:
        if word in text:
            score["Transmitter"] += 1

    best = max(score, key=score.get)
    return best if score[best] > 0 else None


def normalize_tag(tag):
    """
    Normalize tag for matching by removing non-alphanumeric.
    Example: 'P-5401BX' → 'P5401BX'
    """
    if not tag:
        return ""
    cleaned = re.sub(r'[^a-zA-Z0-9]', '', tag.upper())
    return cleaned


class FinalSmartMatcherV5:
    def __init__(self, root):
        self.root = root
        self.root.title("🔧 Final Smart Matcher v5 - Production Ready")
        self.root.geometry("850x700")
        self.root.resizable(False, False)

        self.main_file = None
        self.data_files = {
            "Motor": {"enabled": False, "path": "", "sheet": "", "tag_col": "", "date_col": ""},
            "Valve": {"enabled": False, "path": "", "sheet": "", "tag_col": "", "date_col": ""},
            "Transmitter": {"enabled": False, "path": "", "sheet": "", "tag_col": "", "date_col": ""}
        }

        self.setup_ui()

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky="NSEW")

        # Grid config
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # ——————————————————— Main File ——————————————————— #
        ttk.Label(main_frame, text="📁 Main Tags File", font=("Helvetica", 12, "bold")).grid(
            row=0, column=0, columnspan=6, sticky="w", pady=(0, 10))

        self.main_path = ttk.Entry(main_frame, width=45)
        self.main_path.grid(row=1, column=0, columnspan=3, sticky="ew", padx=(0, 10))
        ttk.Button(main_frame, text="Browse", command=self.browse_main).grid(row=1, column=3)

        self.main_sheet = ttk.Combobox(main_frame, state="readonly", width=25)
        self.main_sheet.set("Select Sheet")
        self.main_sheet.grid(row=2, column=0, padx=(0, 10), pady=5)
        self.main_sheet.bind("<<ComboboxSelected>>", lambda e: self.load_main_columns())

        ttk.Label(main_frame, text="Tag Column:").grid(row=3, column=0, sticky="w", pady=5)
        self.main_tag_col = ttk.Combobox(main_frame, state="readonly", width=25)
        self.main_tag_col.set("Select Tag Column")
        self.main_tag_col.grid(row=3, column=1, padx=(0, 10), pady=5)

        ttk.Label(main_frame, text="Desc Column:").grid(row=3, column=2, sticky="w", pady=5)
        self.main_desc_col = ttk.Combobox(main_frame, state="readonly", width=25)
        self.main_desc_col.set("Select Desc Column")
        self.main_desc_col.grid(row=3, column=3, padx=(0, 10), pady=5)

        # ——————————————————— Data Files ——————————————————— #
        ttk.Label(main_frame, text="🗃️ Data Source Files", font=("Helvetica", 12, "bold")).grid(
            row=4, column=0, columnspan=6, sticky="w", pady=(25, 10))

        # Headers
        ttk.Label(main_frame, text="Use?").grid(row=5, column=0, padx=10)
        ttk.Label(main_frame, text="File Path").grid(row=5, column=1, sticky="w")
        ttk.Label(main_frame, text="Sheet").grid(row=5, column=2, padx=10)
        ttk.Label(main_frame, text="Tag Col").grid(row=5, column=3, padx=10)
        ttk.Label(main_frame, text="Date Col").grid(row=5, column=4, padx=10)

        self.file_rows = {}
        sources = ["Motor", "Valve", "Transmitter"]
        for idx, name in enumerate(sources):
            row_idx = 6 + idx
            enabled_var = tk.BooleanVar()
            chk = ttk.Checkbutton(main_frame, variable=enabled_var)
            chk.grid(row=row_idx, column=0, padx=10)

            path_entry = ttk.Entry(main_frame, width=25)
            path_entry.grid(row=row_idx, column=1, padx=(0, 10))

            sheet_combo = ttk.Combobox(main_frame, state="readonly", width=12)
            sheet_combo.set("Select Sheet")
            sheet_combo.grid(row=row_idx, column=2, padx=5)

            tag_combo = ttk.Combobox(main_frame, state="readonly", width=10)
            tag_combo.set("Tag")
            tag_combo.grid(row=row_idx, column=3, padx=5)

            date_combo = ttk.Combobox(main_frame, state="readonly", width=10)
            date_combo.set("Date")
            date_combo.grid(row=row_idx, column=4, padx=5)

            browse_btn = ttk.Button(main_frame, text="Browse",
                                    command=lambda n=name: self.browse_data(n))
            browse_btn.grid(row=row_idx, column=5, padx=5)

            self.file_rows[name] = {
                "chk_var": enabled_var,
                "path_entry": path_entry,
                "sheet_combo": sheet_combo,
                "tag_combo": tag_combo,
                "date_combo": date_combo
            }

        # ——————————————————— Run Button ——————————————————— #
        self.run_btn = ttk.Button(main_frame, text="🚀 Analyze & Match", command=self.run_matching)
        self.run_btn.grid(row=10, column=0, columnspan=6, pady=30)

        self.progress = ttk.Progressbar(main_frame, mode="determinate", length=550)
        self.progress.grid(row=11, column=0, columnspan=6, pady=10)
        self.progress.grid_remove()

    def browse_main(self):
        path = filedialog.askopenfilename(title="Select Main Tags File", filetypes=[("Excel Files", "*.xlsx")])
        if not path:
            return
        self.main_path.delete(0, tk.END)
        self.main_path.insert(0, path)
        self.main_file = path
        try:
            sheets = pd.ExcelFile(path).sheet_names
            self.main_sheet['values'] = sheets
            self.main_sheet.set(sheets[0])
            self.load_main_columns()
        except Exception as e:
            messagebox.showerror("Error", f"Cannot read file:\n{e}")

    def browse_data(self, name):
        path = filedialog.askopenfilename(title=f"Select {name} Data File", filetypes=[("Excel Files", "*.xlsx")])
        if not path:
            return
        row = self.file_rows[name]
        row["path_entry"].delete(0, tk.END)
        row["path_entry"].insert(0, path)
        try:
            sheets = pd.ExcelFile(path).sheet_names
            row["sheet_combo"]['values'] = sheets
            row["sheet_combo"].set(sheets[0])
            self.load_data_columns(name)
        except Exception as e:
            messagebox.showerror("Error", f"Cannot read file:\n{e}")

    def load_main_columns(self):
        if not self.main_file or self.main_sheet.get() == "Select Sheet":
            return
        try:
            df = pd.read_excel(self.main_file, sheet_name=self.main_sheet.get())
            cols = df.columns.tolist()
            self.main_tag_col['values'] = cols
            self.main_desc_col['values'] = cols
            if len(cols) > 0:
                self.main_tag_col.set(cols[0])
            if len(cols) > 1:
                self.main_desc_col.set(cols[1])
            else:
                self.main_desc_col.set(cols[0])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read main file:\n{e}")

    def load_data_columns(self, name):
        path = self.file_rows[name]["path_entry"].get()
        sheet = self.file_rows[name]["sheet_combo"].get()
        if not path or sheet == "Select Sheet":
            return
        try:
            df = pd.read_excel(path, sheet_name=sheet)
            cols = df.columns.tolist()
            tag_combo = self.file_rows[name]["tag_combo"]
            date_combo = self.file_rows[name]["date_combo"]
            tag_combo['values'] = cols
            date_combo['values'] = cols
            if cols:
                tag_combo.set(cols[0])
                date_combo.set(cols[1] if len(cols) > 1 else cols[0])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read data file:\n{e}")

    def run_matching(self):
        if not self.main_file or self.main_sheet.get() == "Select Sheet" or \
           self.main_tag_col.get() == "Select Tag Column" or self.main_desc_col.get() == "Select Desc Column":
            messagebox.showerror("Error", "Please configure Main File completely.")
            return

        enabled_sources = [name for name in self.data_files if self.file_rows[name]["chk_var"].get()]
        if not enabled_sources:
            messagebox.showwarning("Warning", "Please enable at least one data source.")
            return

        try:
            # Read main file
            df_main = pd.read_excel(self.main_file, sheet_name=self.main_sheet.get())
            tag_col_main = self.main_tag_col.get()
            desc_col_main = self.main_desc_col.get()

            if tag_col_main not in df_main.columns or desc_col_main not in df_main.columns:
                messagebox.showerror("Error", "Selected columns not found in main file.")
                return

            results = []

            total = len(df_main)
            self.progress.grid()
            self.progress['value'] = 0

            for idx, row in df_main.iterrows():
                raw_tag_text = str(row[tag_col_main])
                desc_text = str(row[desc_col_main])

                detected_tag = extract_tag(raw_tag_text)
                device_type = classify_from_text(desc_text)

                if not detected_tag:
                    results.append({
                        "Raw_Tag": raw_tag_text,
                        "Cleaned_Tag": "",
                        "Type": device_type or "Unknown",
                        "Last_Routine_Date": "No Valid Tag"
                    })
                    continue

                if not device_type or not self.file_rows[device_type]["chk_var"].get():
                    results.append({
                        "Raw_Tag": raw_tag_text,
                        "Cleaned_Tag": detected_tag,
                        "Type": device_type or "Unknown",
                        "Last_Routine_Date": "Not Processed"
                    })
                    continue

                # Load corresponding data file
                conf = self.file_rows[device_type]
                path = conf["path_entry"].get()
                sheet = conf["sheet_combo"].get()
                tag_col_data = conf["tag_combo"].get()
                date_col_data = conf["date_combo"].get()

                if not all([path, sheet, tag_col_data, date_col_data]):
                    result_date = "Config Incomplete"
                else:
                    try:
                        df_src = pd.read_excel(path, sheet_name=sheet)
                        df_src[tag_col_data] = df_src[tag_col_data].astype(str)
                        df_src["norm_tag"] = df_src[tag_col_data].apply(normalize_tag)
                        target_norm = normalize_tag(detected_tag)

                        matched = df_src[df_src["norm_tag"] == target_norm]
                        if not matched.empty:
                            date_val = matched.iloc[0][date_col_data]
                            result_date = str(date_val) if pd.notna(date_val) else "No Date"
                        else:
                            result_date = "Not Found in Data File"
                    except Exception:
                        result_date = "Error Reading File"

                results.append({
                    "Raw_Tag": raw_tag_text,
                    "Cleaned_Tag": detected_tag,
                    "Type": device_type,
                    "Last_Routine_Date": result_date
                })

                self.progress['value'] = (idx + 1) / total * 100
                self.root.update_idletasks()

            # Save output
            result_df = pd.DataFrame(results)
            output_path = filedialog.asksaveasfilename(
                title="Save Output",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
            if output_path:
                result_df.to_excel(output_path, index=False)
                messagebox.showinfo("Success", f"✅ Done!\nSaved to:\n{output_path}")
            else:
                messagebox.showwarning("Cancelled", "Save operation was cancelled.")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        finally:
            self.progress.grid_remove()
            self.run_btn.config(state="normal")


if __name__ == "__main__":
    root = tk.Tk()
    app = FinalSmartMatcherV5(root)
    root.mainloop()
