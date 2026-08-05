import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import os
import re
from datetime import datetime

# آدرس فایل اکسل
file_path = r'z:\electrical\8-Maintenance History\Calibration history.xlsx'

# تبدیل نام ماه فارسی به عدد
persian_months = {
    'فروردین': 1, 'اردیبهشت': 2, 'خرداد': 3,
    'تیر': 4, 'مرداد': 5, 'شهریور': 6,
    'مهر': 7, 'آبان': 8, 'آذر': 9,
    'دی': 10, 'بهمن': 11, 'اسفند': 12
}

# تبدیل تقریبی سال شمسی به میلادی
def jalali_to_gregorian(jy, jm, jd):
    jy += 1595
    days = 365 * jy + int(jy / 33) * 8 + (jy % 33 + 3) // 4
    days += jd + (jm < 7) * (jm - 1) * 31 + (jm > 6) * ((jm - 7) * 30 + 186)
    gy, days = divmod(days - 226894, 1461)
    gy = gy * 4 + 1026 + (days > 365)
    days -= 1 if (days < 366) else 0
    gys = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    if (gy % 4 == 0 and gy % 100 != 0) or (gy % 400 == 0):
        gys[1] = 29
    for gm, g in enumerate(gys):
        if days < g:
            gd = days + 1
            break
        days -= g
    else:
        return None
    return gy, gm + 1, gd

# تحلیل تاریخ شمسی
def parse_jalali_date(date_str):
    try:
        date_str = str(date_str).strip()
        parts = re.split(r'[\/\-]', date_str)
        if len(parts) != 3:
            return None
        y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
        if y < 1300 or m < 1 or m > 12 or d < 1 or d > 31:
            return None
        return jalali_to_gregorian(y, m, d)
    except:
        return None

# خواندن داده‌ها
def load_data():
    if not os.path.exists(file_path):
        messagebox.showerror("خطا", f"فایل یافت نشد:\n{file_path}")
        return None

    try:
        excel_file = pd.ExcelFile(file_path)
        sheet_names = [sheet for sheet in excel_file.sheet_names if re.match(r'^1[34]\d{2}$', sheet.strip())]
        if not sheet_names:
            messagebox.showwarning("هشدار", "هیچ شیت سالی یافت نشد.")
            return None

        all_data = pd.DataFrame()
        sheets = {}

        for sheet_name in sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            df = df.dropna(how='all').reset_index(drop=True)
            if 'Tag No.' not in df.columns:
                continue
            df['Year'] = sheet_name.strip()
            df['CALIBRATION DATE'] = df['CALIBRATION DATE'].astype(str).str.strip()
            df['Sheet'] = sheet_name

            # اضافه کردن تاریخ میلادی
            df['Gregorian Date'] = df['CALIBRATION DATE'].apply(parse_jalali_date)
            df['Gregorian Date'] = df['Gregorian Date'].apply(
                lambda x: datetime(*x) if x else None
            )

            sheets[sheet_name] = df.copy()
            all_data = pd.concat([all_data, df], ignore_index=True)

        return all_data, sheets

    except Exception as e:
        messagebox.showerror("خطا", f"خطا در خواندن فایل:\n{str(e)}")
        return None

# ایجاد برنامه اصلی
class CalibrationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🔧 سیستم پیش‌بینی کالیبراسیون - نسخه نهایی")
        self.root.geometry("1400x800")
        self.root.state('zoomed')

        load_result = load_data()
        if load_result is None:
            root.destroy()
            return

        self.data, self.sheets = load_result
        self.current_display_df = pd.DataFrame()
        self.create_widgets()

    def create_widgets(self):
        control_frame = ttk.Frame(self.root)
        control_frame.pack(pady=15, padx=10, fill='x')

        # جستجوی تگ
        ttk.Label(control_frame, text="🔍 جستجو تگ:").grid(row=0, column=0, padx=5, sticky='e')
        self.tag_var = tk.StringVar()
        tag_entry = ttk.Entry(control_frame, textvariable=self.tag_var, width=20, font=('Tahoma', 10))
        tag_entry.grid(row=0, column=1, padx=5)
        ttk.Button(control_frame, text="نمایش تاریخچه", command=self.show_tag_history).grid(row=0, column=2, padx=10)

        # ماه و سال
        ttk.Label(control_frame, text="📅 ماه:").grid(row=0, column=3, padx=5, sticky='e')
        self.month_var = tk.StringVar(value="فروردین")
        month_cb = ttk.Combobox(control_frame, textvariable=self.month_var,
                                values=list(persian_months.keys()), width=10, state="readonly")
        month_cb.grid(row=0, column=4, padx=5)

        ttk.Label(control_frame, text="پيش بيني کاليبره چه سالي نمايش داده شود →").grid(row=0, column=5, padx=5, sticky='e')
        self.target_year_var = tk.StringVar(value="1404")
        target_year_entry = ttk.Entry(control_frame, textvariable=self.target_year_var, width=10)
        target_year_entry.grid(row=0, column=6, padx=5)

        # تعداد سابقه
        ttk.Label(control_frame, text="چند سال سابقه نمايش داده شود →").grid(row=0, column=7, padx=5, sticky='e')
        self.history_count_var = tk.StringVar(value="3")
        history_entry = ttk.Entry(control_frame, textvariable=self.history_count_var, width=5)
        history_entry.grid(row=0, column=8, padx=5)

        # سال مبنا
        ttk.Label(control_frame, text="بر اساس چه سالي کاليبره ج جديد انجام شود →").grid(row=0, column=9, padx=5, sticky='e')
        self.base_year_var = tk.StringVar(value="1403")
        self.base_year_cb = ttk.Combobox(control_frame, textvariable=self.base_year_var, width=10, state="readonly")
        self.base_year_cb.grid(row=0, column=10, padx=5)
        self.update_base_year_options()

        # دکمه‌ها
        ttk.Button(control_frame, text="← کالیبره‌های این ماه از سال", command=self.show_month_calibrations).grid(row=0, column=11, padx=5)
        ttk.Button(control_frame, text="🔮 پیش‌بینی آینده", command=self.predict_next_calibration).grid(row=0, column=12, padx=5)
        ttk.Button(control_frame, text="💾 ذخیره در اکسل", command=self.save_to_excel).grid(row=0, column=13, padx=5)

        # --- جدول نتایج ---
        result_frame = ttk.Frame(self.root)
        result_frame.pack(fill='both', expand=True, padx=10, pady=5)

        self.tree = ttk.Treeview(result_frame, columns=[], show='headings', height=28)
        vsb = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(result_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')

        # --- اضافه کردن eventها ---
        self.tree.bind("<Control-c>", self.copy_selection)
        self.tree.bind("<Button-1>", self.on_click)

    def update_base_year_options(self):
        if hasattr(self, 'base_year_cb'):
            years = sorted([int(x) for x in self.sheets.keys() if x.isdigit()], reverse=True)
            self.base_year_cb['values'] = years
            if years:
                self.base_year_var.set(str(years[0]))

    def show_tag_history(self):
        tag = self.tag_var.get().strip()
        if not tag:
            messagebox.showwarning("هشدار", "لطفاً یک Tag No. وارد کنید.")
            return

        # 🔍 جستجوی دقیق و بدون حساسیت به بزرگی/کوچکی
        tag_lower = tag.lower()
        filtered = self.data[
            self.data['Tag No.'].astype(str).str.strip().str.lower() == tag_lower
        ]

        if filtered.empty:
            messagebox.showinfo("نتیجه", f"تگ '{tag}' یافت نشد.")
            return

        result = filtered.sort_values(by='Gregorian Date', ascending=False).reset_index(drop=True)
        result.index += 1
        result['Row'] = result.index

        display_cols = ['Row', 'Tag No.', 'CALIBRATION DATE', 'Description', 'Year', 'AREA', 'Service', 'P & ID']
        self.current_display_df = result[display_cols].copy()
        self.display_results(result, display_columns=display_cols)

    def show_month_calibrations(self):
        month_name = self.month_var.get()
        target_year_str = self.target_year_var.get().strip()
        n_str = self.history_count_var.get().strip()
        base_year_str = self.base_year_var.get().strip()

        if not target_year_str.isdigit():
            messagebox.showwarning("هشدار", "سال هدف نامعتبر است.")
            return
        if not n_str.isdigit() or int(n_str) < 1:
            messagebox.showwarning("هشدار", "تعداد سابقه باید عدد مثبت باشد.")
            return
        if not base_year_str.isdigit():
            messagebox.showwarning("هشدار", "سال مبنا نامعتبر است.")
            return

        base_year = int(base_year_str)
        n = int(n_str)
        month_num = persian_months[month_name]

        # --- 1. گرفتن تگ‌های کالیبره‌شده در "ماه X سال مبنا"
        current_sheet = self.sheets.get(str(base_year))
        if current_sheet is None:
            messagebox.showinfo("نتیجه", f"شیت سال {base_year} یافت نشد.")
            return

        pattern = rf'^{base_year}/{month_num}/|^{base_year}/{month_num:02d}/'
        current_month_tags = current_sheet[current_sheet['CALIBRATION DATE'].str.match(pattern)]
        current_month_tags = current_month_tags.drop_duplicates(subset=['Tag No.'], keep='first')

        if current_month_tags.empty:
            messagebox.showinfo("نتیجه", f"هیچ کالیبراسیونی در {month_name} {base_year} یافت نشد.")
            return

        # مرتب‌سازی بر اساس تاریخ
        current_month_tags = current_month_tags.sort_values(by='Gregorian Date', ascending=False).reset_index(drop=True)

        # --- 2. جمع‌آوری تمام کالیبراسیون‌های قبل از سال مبنا
        all_past_data = []
        for year_str in self.sheets.keys():
            try:
                year = int(year_str)
                if year >= base_year:
                    continue
                sheet = self.sheets[year_str]
                all_past_data.append(sheet.copy())
            except:
                continue

        if not all_past_data:
            messagebox.showinfo("نتیجه", "هیچ داده‌ای از سال‌های قبل یافت نشد.")
            return

        combined_past = pd.concat(all_past_data, ignore_index=True)
        combined_past = combined_past.sort_values(by='Gregorian Date', ascending=False)

        # --- 3. پیدا کردن N بار آخرین کالیبراسیون برای هر تگ
        rows = []
        for _, row in current_month_tags.iterrows():
            tag = row['Tag No.']
            base_cal = row['CALIBRATION DATE']
            base_desc = row.get('Description', '')

            new_row = {
                'Tag No.': tag,
                'Calibration (Current)': base_cal,
                'Description (Current)': base_desc
            }

            # فیلتر سابقه این تگ
            tag_history = combined_past[combined_past['Tag No.'] == tag].head(n)

            for i in range(n):
                if i < len(tag_history):
                    cal = tag_history.iloc[i]['CALIBRATION DATE']
                    desc = tag_history.iloc[i].get('Description', '')
                    year_val = tag_history.iloc[i]['Year']
                else:
                    cal = desc = year_val = ""

                new_row[f'Calibration {year_val}'] = cal
                new_row[f'Description {year_val}'] = desc

            rows.append(new_row)

        result_df = pd.DataFrame(rows)
        result_df = result_df.reset_index(drop=True)
        result_df.index += 1
        result_df['Row'] = result_df.index

        # --- 4. ستون‌های نمایش: اول Row، Tag No.، فعلی، بعد سال‌های قبل (مرتب‌شده)
        display_cols = ['Row', 'Tag No.', 'Calibration (Current)', 'Description (Current)']

        # پیدا کردن سال‌هایی که در سابقه ظاهر شدن
        past_years = sorted(set(
            str(int(year)) for col in result_df.columns
            if col.startswith('Calibration ') and col != 'Calibration (Current)'
            for year in [col.replace('Calibration ', '')] if year.isdigit()
        ), reverse=True)

        for yr in past_years:
            if f'Calibration {yr}' in result_df.columns:
                display_cols += [f'Calibration {yr}', f'Description {yr}']

        display_cols = [col for col in display_cols if col in result_df.columns]
        self.current_display_df = result_df[display_cols].copy()
        self.display_results(result_df, display_columns=display_cols)

    def predict_next_calibration(self):
        month_name = self.month_var.get()
        target_year_str = self.target_year_var.get().strip()
        base_year_str = self.base_year_var.get().strip()

        if not target_year_str.isdigit() or not base_year_str.isdigit():
            messagebox.showwarning("هشدار", "سال هدف یا سال مبنا نامعتبر است.")
            return

        target_year = int(target_year_str)
        base_year = int(base_year_str)
        month_num = persian_months[month_name]

        base_sheet = self.sheets.get(str(base_year))
        if base_sheet is None:
            messagebox.showinfo("نتیجه", f"شیت سال {base_year} یافت نشد.")
            return

        base_filtered = base_sheet[
            base_sheet['CALIBRATION DATE'].str.match(rf'^{base_year}/{month_num}/|^{base_year}/{month_num:02d}/')
        ]

        if base_filtered.empty:
            messagebox.showinfo("نتیجه", f"هیچ دستگاهی در {month_name} {base_year} کالیبره نشده است.")
            return

        latest_data = self.data.sort_values(by='Gregorian Date', ascending=False)
        latest_per_tag = latest_data.groupby('Tag No.').first().reset_index()

        base_tags = base_filtered[['Tag No.']].drop_duplicates()
        result = pd.merge(base_tags, latest_per_tag, on='Tag No.', how='left')

        target_sheet = self.sheets.get(str(target_year))
        if target_sheet is None:
            result['Status'] = "⚠️ نیاز به کالیبراسیون"
        else:
            target_in_year = target_sheet[
                target_sheet['CALIBRATION DATE'].str.startswith(f"{target_year}/")
            ][['Tag No.', 'CALIBRATION DATE', 'Gregorian Date']].copy()

            if not target_in_year.empty:
                target_in_year = target_in_year.sort_values(by='Gregorian Date', ascending=False)
                latest_in_target = target_in_year.groupby('Tag No.').first()['CALIBRATION DATE']
            else:
                latest_in_target = pd.Series(dtype=str)

            result['Status'] = result['Tag No.'].map(latest_in_target).fillna("⚠️ نیاز به کالیبراسیون")

        result['Suggested Calibration'] = f"{target_year}/{month_num:02d}/01 (پیشنهادی)"
        result = result.reset_index(drop=True)
        result.index += 1
        result['Row'] = result.index

        rename_dict = {
            'CALIBRATION DATE': 'Last Calibration Date',
            'Year': 'Last Calibration Year'
        }
        result = result.rename(columns=rename_dict)
        final_display = [
            'Row', 'Tag No.',
            'Last Calibration Date', 'Last Calibration Year', 'Description',
            'AREA', 'Service', 'P & ID',
            'Suggested Calibration',
            'Status'
        ]
        self.current_display_df = result[final_display].copy()
        self.display_results(result, display_columns=final_display)

    def display_results(self, df, display_columns=None):
        if not hasattr(self, 'tree'):
            return

        for item in self.tree.get_children():
            self.tree.delete(item)

        cols = display_columns if display_columns else [col for col in df.columns if col != 'Gregorian Date']
        self.tree["columns"] = cols

        for col in cols:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_column(c))
            self.tree.column(col, width=130, anchor='center', minwidth=80)

        self.result_df = df[cols].copy()
        self.displayed_columns = cols

        for _, row in self.result_df.iterrows():
            values = [str(row[col]) if pd.notna(row[col]) else "" for col in cols]
            self.tree.insert("", "end", values=values)

    def sort_column(self, col):
        reverse = False
        for c in self.tree["columns"]:
            if self.tree.heading(c, "text").endswith(" 🔽") or self.tree.heading(c, "text").endswith(" 🔼"):
                if c == col:
                    reverse = not self.tree.heading(c, "text").endswith(" 🔼")
                self.tree.heading(c, text=c.replace(" 🔽", "").replace(" 🔼", ""))

        # تشخیص نوع مرتب‌سازی
        if col == "Row":
            # مرتب‌سازی عددی
            self.result_df['Row'] = pd.to_numeric(self.result_df['Row'], errors='coerce')
            self.result_df = self.result_df.sort_values(by=col, ascending=not reverse, kind='stable')
        else:
            # مرتب‌سازی رشته‌ای (حروف و اعداد مخلوط)
            self.result_df = self.result_df.sort_values(by=col, ascending=not reverse, key=lambda x: x.astype(str), kind='stable')

        # نمایش مجدد
        for item in self.tree.get_children():
            self.tree.delete(item)

        for _, row in self.result_df.iterrows():
            values = [str(row[col]) if pd.notna(row[col]) else "" for col in self.displayed_columns]
            self.tree.insert("", "end", values=values)

        # نمایش فلش
        arrow = " 🔼" if reverse else " 🔽"
        self.tree.heading(col, text=col + arrow)

        
    def on_click(self, event):
        if self.tree.identify_region(event.x, event.y) != "heading":
            return

    def copy_selection(self, event):
        selected_item = self.tree.selection()
        if not selected_item:
            return
        item = selected_item[0]
        col_id = self.tree.identify_column(event.x)
        if not col_id:
            return
        col_index = int(col_id[1:]) - 1
        values = self.tree.item(item, "values")
        if col_index < len(values):
            value = values[col_index]
            self.root.clipboard_clear()
            self.root.clipboard_append(str(value))
            self.root.update()

    def save_to_excel(self):
        if self.current_display_df.empty:
            messagebox.showwarning("هشدار", "هیچ داده‌ای برای ذخیره وجود ندارد.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="ذخیره گزارش در اکسل"
        )
        if not file_path:
            return

        try:
            self.current_display_df.to_excel(file_path, index=False, sheet_name="گزارش")
            messagebox.showinfo("موفقیت", f"گزارش با موفقیت ذخیره شد:\n{file_path}")
        except Exception as e:
            messagebox.showerror("خطا", f"ذخیره فایل ناموفق بود:\n{str(e)}")


# --- اجرای برنامه ---
if __name__ == "__main__":
    root = tk.Tk()
    app = CalibrationApp(root)
    root.mainloop()
