from openpyxl import load_workbook
from openpyxl.styles import Font
import os
import shutil
import jdatetime


# ⚙️ تابع ایجاد پوشه با تاریخ و زمان شمسی
def create_backup_folders(base_path):
    now = jdatetime.datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H-%M")

    backup_root = os.path.join(base_path, "Backup")
    date_folder = os.path.join(backup_root, date_str)
    time_folder = os.path.join(date_folder, time_str)

    os.makedirs(time_folder, exist_ok=True)

    return {
        'date_folder': date_folder,
        'time_folder': time_str
    }


# 🧼 بکاپ عمومی (همه .xlsx/.xlsb) → Backup/1403-12-21/
def backup_general_files(folder_path):
    folders = create_backup_folders(folder_path)
    general_folder = folders['date_folder']

    for filename in os.listdir(folder_path):
        if filename.lower().endswith(('.xlsx', '.xlsb')):
            source_file = os.path.join(folder_path, filename)
            dest_file = os.path.join(general_folder, filename)

            try:
                shutil.copy2(source_file, dest_file)
                print(f"[OK] بکاپ عمومی گرفته شد: {filename}")
            except Exception as e:
                print(f"[ERROR] مشکل در کپی {filename}: {str(e)}")


# 📦 بکاپ اختصاصی (فقط سه فایل مهم) → Backup/1403-12-21/15-30/
def backup_specific_files(folder_path):
    folders = create_backup_folders(folder_path)
    specific_folder = os.path.join(folders['date_folder'], folders['time_folder'])

    files_to_backup = [
        "All motors.xlsx",
        "ALL Valve.xlsx",
        "ALL TRANSMITER.xlsx"
    ]

    for filename in files_to_backup:
        src = os.path.join(folder_path, filename)
        dst = os.path.join(specific_folder, filename)

        if not os.path.exists(src):
            print(f"[SKIP] فایل وجود ندارد: {filename}")
            continue

        try:
            shutil.copy2(src, dst)
            print(f"[OK] بکاپ اختصاصی گرفته شد: {filename}")
        except Exception as e:
            print(f"[ERROR] مشکل در کپی {filename}: {str(e)}")


# 📥 تابع پردازش فایل data_motors.xlsx برای خواندن مدل بیرینگ
def process_bearing_data(data_file_path, target_tag):
    bearing_drive_model = ""
    bearing_non_drive_model = ""

    try:
        wb_data = load_workbook(data_file_path)

        # جستجو در تمام شیت‌ها
        for sheet_name in wb_data.sheetnames:
            sheet_data = wb_data[sheet_name]
            for row in sheet_data.iter_rows(values_only=True):
                tag_value = str(row[1]) if len(row) > 1 else ""  # ستون B = TAG NO.
                k_value = row[10] if len(row) > 10 else None     # ستون K = BEARING-DRIVE
                l_value = row[11] if len(row) > 11 else None     # ستون L = BEARING-NON DRIVE

                # فقط ردیف‌هایی که TAG NO. با اسم شیت فعلی مطابقت داشته باشه
                if tag_value == target_tag:
                    if k_value is not None:
                        bearing_drive_model = k_value
                    if l_value is not None:
                        bearing_non_drive_model = l_value

    except Exception as e:
        print(f"[ERROR] خطا در پردازش data_motors.xlsx: {e}")

    return bearing_drive_model, bearing_non_drive_model


# 📊 تابع پردازش اصلی فایل Excel
def process_excel(file_path):
    folder_path = os.path.dirname(file_path)

    # 🔁 مرحله ۱: بکاپ عمومی و اختصاصی
    backup_general_files(folder_path)
    backup_specific_files(folder_path)

    # مرحله ۲: بارگذاری فایل اصلی
    try:
        wb = load_workbook(file_path)
        list_sheet = wb['List']

        # پاک کردن فقط ستون‌های 1 تا 12 ردیف‌های ≤ 220
        for row in list_sheet.iter_rows(max_row=220, max_col=12):
            for cell in row:
                cell.value = None

        # محاسبه سال جاری و قبلی
        current_year = jdatetime.datetime.now().year
        last_year = current_year - 1
        two_years_ago = current_year - 2

        emsal = current_year
        emsal_str = str(emsal)
        last_year_str = str(last_year)
        two_years_ago_str = str(two_years_ago)

        # تعریف سربرگ‌ها با استفاده از سال دینامیک
        headers = [
            "List", 
            "BEARING-DRIVE", 
            "BEARING-NON DRIVE", 
            "تاریخ آخرین روتین B",
            f"تعداد تعویض بیرینگ{emsal_str}",           # column=5
            f"تعداد تعویض بیرینگ{last_year_str}",       # column=6
            f"تعداد تعویض بیرینگ{two_years_ago_str}",   # column=7 ✅ جدید
            f"تعداد تزریق گریس {last_year_str}",        # column=8
            f"تعداد تزریق گریس {emsal_str}",            # column=9
            f"روتین در {emsal} انجام شده",               # column=10
            "روتین 6 ماه قبل انجام شده",                # column=11
            "روتین در ماه جاری انجام شده",              # column=12
            f"تعداد کار بر روی مقره و روتور و تخت کلم {emsal_str}", # column=13
            f"تعداد تعمیر و تعویض قطعات LCS {emsal_str}", # column=14
            f"تعداد کار بر روی سیم پیچی {last_year_str}", # column=15
            "نیاز به خرید قطعه دارد",                   # column=16
            f"روتین گریس کاری {emsal_str}",             # column=17
            "موتور باید انتقال یابد",                   # column=18

            *[f"تعویض بیرینگ تاریخ {i-1}" for i in range(2, 22)],     # 19 تا 39
            *[f"شارژ گریس تاریخ {i-1}" for i in range(2, 22)],         # 40 تا 60
            *[f"تخته کلم تاریخ {i-1}" for i in range(2, 11)],          # 61 تا 70
            *[f"سیم پیچی تاریخ {i-1}" for i in range(2, 11)],          # 71 تا 80
            *[f"LCS تاریخ {i-1}" for i in range(2, 11)]                # 81 تا 90
        ]

        # نوشتن سربرگ‌ها
        for col_idx, header in enumerate(headers, start=1):
            list_sheet.cell(row=1, column=col_idx).value = header

        row_idx = 2  # شروع از ردیف 2

        # مسیر فایل data_motors.xlsx
        data_file_path = r"z:\electrical\8-Maintenance History\data_motors.xlsx"

        # پردازش هر شیت (غیر از List)
        for sheet_name in wb.sheetnames:
            if sheet_name == 'List':
                continue

            sheet = wb[sheet_name]

            bearing_dates = []
            grease_dates = []
            terminal_box_dates = []
            lcs_dates = []
            wiring_dates = []
            routine_dates = []
            buy_flag = False
            pic_flag = False
            shop_flag = False
            grease_current_year_count = 0

            # تحلیل داده‌ها
            for row in sheet.iter_rows(values_only=True):
                b_value = str(row[1]) if len(row) > 1 else ""
                c_value = row[2] if len(row) > 2 else None
                d_value = row[3] if len(row) > 3 else None
                e_value = row[4] if len(row) > 4 else None
                f_value = row[5] if len(row) > 5 else None
                g_value = row[6] if len(row) > 6 else None
                h_value = row[7] if len(row) > 7 else None
                i_value = row[8] if len(row) > 8 else None
                j_value = row[9] if len(row) > 9 else None
                k_value = row[10] if len(row) > 10 else None
                l_value = row[11] if len(row) > 11 else None

                # ✅ اضافه کردن تاریخ اگر در ستون E یا L علامت داشته باشه
                if e_value == chr(252) or l_value == chr(252):
                    grease_dates.append(b_value)

                if c_value == chr(252):
                    bearing_dates.append(b_value)
                if d_value == chr(252):
                    routine_dates.append(b_value)
                if f_value == chr(252):
                    terminal_box_dates.append(b_value)
                if g_value == chr(252):
                    lcs_dates.append(b_value)
                if h_value == chr(252):
                    wiring_dates.append(b_value)
                if i_value == chr(252):
                    buy_flag = True
                if j_value == chr(252):
                    pic_flag = True
                if k_value == chr(252) and b_value.startswith(emsal_str):
                    grease_current_year_count += 1

            # 🔁 خواندن مدل بیرینگ از data_motors.xlsx
            drive_model, non_drive_model = process_bearing_data(data_file_path, sheet_name)

            # ✅ نوشتن مدل بیرینگ در ستون 2 و 3
            list_sheet.cell(row=row_idx, column=2, value=drive_model)
            list_sheet.cell(row=row_idx, column=3, value=non_drive_model)

            # ✅ ایجاد هایپرلینک
            cell = list_sheet.cell(row=row_idx, column=1)
            cell.value = sheet_name
            cell.hyperlink = f"#'{sheet_name}'!A1"
            cell.font = Font(color="0000FF", underline="single")

            # پردازش روتین
            if routine_dates:
                last_routine_date = max(routine_dates)
                routine_current_year = sum(1 for x in routine_dates if x.startswith(emsal_str))

                # ✅ محاسبه روتین 6 ماهه بر اساس بازه زمانی
                current_month = jdatetime.datetime.now().month
                if 1 <= current_month <= 6:
                    routine_last_6_months = sum(
                        1 for x in routine_dates 
                        if x.startswith(emsal_str) and 1 <= int(x[5:7]) <= 6
                    )
                elif 7 <= current_month <= 12:
                    routine_last_6_months = sum(
                        1 for x in routine_dates 
                        if x.startswith(emsal_str) and 7 <= int(x[5:7]) <= 12
                    )
                else:
                    routine_last_6_months = 0

                routine_current_month = sum(
                    1 for x in routine_dates 
                    if x.startswith(emsal_str) and x[:7] == jdatetime.datetime.now().strftime("%Y/%m")
                )

                list_sheet.cell(row=row_idx, column=4, value=last_routine_date)
                list_sheet.cell(row=row_idx, column=10, value=routine_current_year)
                list_sheet.cell(row=row_idx, column=11, value=routine_last_6_months)
                list_sheet.cell(row=row_idx, column=12, value=routine_current_month)
            else:
                list_sheet.cell(row=row_idx, column=4, value="-")
                list_sheet.cell(row=row_idx, column=10, value=0)
                list_sheet.cell(row=row_idx, column=11, value=0)
                list_sheet.cell(row=row_idx, column=12, value=0)

            # پردازش بیرینگ
            bearing_current_year = sum(1 for x in bearing_dates if x.startswith(emsal_str))
            bearing_last_year = sum(1 for x in bearing_dates if x.startswith(last_year_str))
            bearing_two_years_ago = sum(1 for x in bearing_dates if x.startswith(two_years_ago_str))

            list_sheet.cell(row=row_idx, column=5, value=bearing_current_year)
            list_sheet.cell(row=row_idx, column=6, value=bearing_last_year)
            list_sheet.cell(row=row_idx, column=7, value=bearing_two_years_ago)

            # نوشتن تاریخ‌های تعویض بیرینگ (جدید به قدیم)
            for idx, date in enumerate(bearing_dates[::-1], start=19):  # 19 to 39
                list_sheet.cell(row=row_idx, column=idx, value=date)

            # پردازش گریس
            grease_last_year = sum(1 for x in grease_dates if x.startswith(last_year_str))
            grease_current_year = sum(1 for x in grease_dates if x.startswith(emsal_str))

            list_sheet.cell(row=row_idx, column=8, value=grease_last_year)
            list_sheet.cell(row=row_idx, column=9, value=grease_current_year)

            # نوشتن تاریخ‌های شارژ گریس
            for idx, date in enumerate(grease_dates[::-1], start=39):  # 40 to 60
                list_sheet.cell(row=row_idx, column=idx, value=date)

            # پردازش جعبه ترمینال
            terminal_box_current_year = sum(1 for x in terminal_box_dates if x.startswith(emsal_str))
            list_sheet.cell(row=row_idx, column=13, value=terminal_box_current_year)

            # نوشتن تاریخ‌های جعبه ترمینال
            for idx, date in enumerate(terminal_box_dates[::-1], start=59):  # 61 to 70
                list_sheet.cell(row=row_idx, column=idx, value=date)

            # پردازش LCS
            lcs_current_year = sum(1 for x in lcs_dates if x.startswith(emsal_str))
            list_sheet.cell(row=row_idx, column=14, value=lcs_current_year)

            # نوشتن تاریخ‌های LCS
            for idx, date in enumerate(lcs_dates[::-1], start=77):  # 81 to 90
                list_sheet.cell(row=row_idx, column=idx, value=date)

            # پردازش سیم‌پیچی
            wiring_last_year = sum(1 for x in wiring_dates if x.startswith(last_year_str))
            list_sheet.cell(row=row_idx, column=15, value=wiring_last_year)

            # نوشتن تاریخ‌های سیم‌پیچی
            for idx, date in enumerate(wiring_dates[::-1], start=68):  # 71 to 80
                list_sheet.cell(row=row_idx, column=idx, value=date)

            # پرچس، روتین گریس کاری سال جاری، انتقال
            list_sheet.cell(row=row_idx, column=16, value="Yes" if buy_flag else "No")
            list_sheet.cell(row=row_idx, column=17, value=grease_current_year_count)
            list_sheet.cell(row=row_idx, column=18, value="Yes" if shop_flag else "-")

            # افزایش ردیف
            row_idx += 1

        # 🔚 تمیزکاری نهایی: حذف سطرهای 405 به بعد و ستون‌های 110 به بعد
        max_row = list_sheet.max_row
        max_col = list_sheet.max_column

        if max_row >= 405:
            rows_to_delete = max_row - 404
            list_sheet.delete_rows(405, rows_to_delete)
            print(f"[INFO] {rows_to_delete} ردیف از سطر 405 به بعد حذف شد.")

        if max_col >= 110:
            cols_to_delete = max_col - 109
            list_sheet.delete_cols(110, cols_to_delete)
            print(f"[INFO] {cols_to_delete} ستون از ستون 110 به بعد حذف شد.")

        # ذخیره فایل اصلی
        wb.save(file_path)
        print("[OK] فایل ذخیره شد.")

    except Exception as e:
        print(f"[ERROR] خطا در باز کردن یا ذخیره فایل: {str(e)}")


# 📍 مسیر فایل اصلی
file_path = r"z:\electrical\8-Maintenance History\All motors.xlsx"


# 🚀 اجرای اصلی
if __name__ == "__main__":
    process_excel(file_path)
