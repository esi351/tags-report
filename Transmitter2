from openpyxl import load_workbook
from openpyxl.styles import Font
import os
import shutil
import jdatetime


# ️ تابع ایجاد پوشه با تاریخ و زمان شمسی
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
        'time_folder': time_folder
    }


# 🧼 بکاپ عمومی (همه .xlsx/.xlsb)
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


#  بکاپ اختصاصی (سه فایل مهم)
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

        for sheet_name in wb_data.sheetnames:
            sheet_data = wb_data[sheet_name]
            for row in sheet_data.iter_rows(values_only=True):
                tag_value = str(row[1]) if len(row) > 1 else ""
                k_value = row[10] if len(row) > 10 else None
                l_value = row[11] if len(row) > 11 else None

                if tag_value == target_tag:
                    if k_value is not None:
                        bearing_drive_model = k_value
                    if l_value is not None:
                        bearing_non_drive_model = l_value

    except Exception as e:
        print(f"[ERROR] خطا در پردازش data_motors.xlsx: {e}")

    return bearing_drive_model, bearing_non_drive_model


#  تابع پردازش اصلی فایل Excel
def process_excel(file_path):
    folder_path = os.path.dirname(file_path)

    # 🔁 بکاپ
    backup_general_files(folder_path)
    backup_specific_files(folder_path)

    try:
        wb = load_workbook(file_path)
        list_sheet = wb['List']

        # پاک کردن ستون‌های 1 تا 12، ردیف‌های ≤ 220
        for row in list_sheet.iter_rows(max_row=220, max_col=12):
            for cell in row:
                cell.value = None

        # محاسبه سال جاری
        current_year = jdatetime.datetime.now().year
        last_year = current_year - 1
        two_years_ago = current_year - 2

        emsal_str = str(current_year)
        last_year_str = str(last_year)
        two_years_ago_str = str(two_years_ago)

        # ✅ سربرگ‌ها (با اضافه شدن 2 ستون جدید: اسامی و ساعت)
        headers = [
            "List",
            "BEARING-DRIVE",
            "BEARING-NON DRIVE",
            "تاریخ آخرین روتین B",
            "اسامی انجام‌دهنده آخرین روتین",  # ✅ جدید - ستون 5
            "ساعت کاری آخرین روتین",           # ✅ جدید - ستون 6
            f"تعداد تعویض بیرینگ{emsal_str}",           # ستون 7
            f"تعداد تعویض بیرینگ{last_year_str}",       # ستون 8
            f"تعداد تعویض بیرینگ{two_years_ago_str}",   # ستون 9
            f"تعداد تزریق گریس {last_year_str}",        # ستون 10
            f"تعداد تزریق گریس {emsal_str}",            # ستون 11
            f"روتین در {emsal_str} انجام شده",           # ستون 12
            "روتین 6 ماه قبل انجام شده",                # ستون 13
            "روتین در ماه جاری انجام شده",              # ستون 14
            f"تعداد کار بر روی مقره و روتور و تخت کلم {emsal_str}",  # ستون 15
            f"تعداد تعمیر و تعویض قطعات LCS {emsal_str}",           # ستون 16
            f"تعداد کار بر روی سیم پیچی {last_year_str}",            # ستون 17
            "نیاز به خرید قطعه دارد",                   # ستون 18
            f"روتین گریس کاری {emsal_str}",             # ستون 19
            "موتور باید انتقال یابد",                   # ستون 20

            *[f"تعویض بیرینگ تاریخ {i}" for i in range(2, 23)],  # 21 تا 42
            *[f"شارژ گریس تاریخ {i}" for i in range(2, 23)],      # 43 تا 64
            *[f"تخته کلم تاریخ {i}" for i in range(2, 12)],       # 65 تا 75
            *[f"سیم پیچی تاریخ {i}" for i in range(2, 12)],       # 76 تا 86
            *[f"LCS تاریخ {i}" for i in range(2, 12)]             # 87 تا 97
        ]

        for col_idx, header in enumerate(headers, start=1):
            list_sheet.cell(row=1, column=col_idx).value = header

        row_idx = 2
        data_file_path = r"Z:\8-Maintenance History\data_motors.xlsx"

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
            routine_persons = []  # ✅ اسامی نفرات روتین (ستون I)
            routine_hours = []    # ✅ ساعت کاری روتین (ستون J)
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
                i_value = row[8] if len(row) > 8 else None   # ✅ ستون I = انجام دهنده
                j_value = row[9] if len(row) > 9 else None   # ✅ ستون J = ساعت کاری
                k_value = row[10] if len(row) > 10 else None
                l_value = row[11] if len(row) > 11 else None

                if e_value == chr(252) or l_value == chr(252):
                    grease_dates.append(b_value)

                if c_value == chr(252):
                    bearing_dates.append(b_value)
                if d_value == chr(252):
                    routine_dates.append(b_value)
                    # ✅ ذخیره اسامی و ساعت برای روتین
                    routine_persons.append(str(i_value) if i_value else "")
                    # ✅ تبدیل ساعت به عدد
                    try:
                        hour_val = float(j_value) if j_value is not None else 0
                    except (ValueError, TypeError):
                        hour_val = 0
                    routine_hours.append(hour_val)
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

            # خواندن مدل بیرینگ
            drive_model, non_drive_model = process_bearing_data(data_file_path, sheet_name)

            list_sheet.cell(row=row_idx, column=2, value=drive_model)
            list_sheet.cell(row=row_idx, column=3, value=non_drive_model)

            cell = list_sheet.cell(row=row_idx, column=1)
            cell.value = sheet_name
            cell.hyperlink = f"#'{sheet_name}'!A1"
            cell.font = Font(color="0000FF", underline="single")

            # ✅ پردازش روتین با استخراج اسامی و ساعت
            if routine_dates:
                last_routine_date = max(routine_dates)

                # پیدا کردن تمام ردیف‌هایی که تاریخ آخر رو دارن
                matching_indices = [i for i, d in enumerate(routine_dates) if d == last_routine_date]

                # اسامی نفرات با :::: جدا شده
                persons_list = [routine_persons[i] for i in matching_indices if routine_persons[i]]
                persons_str = "::::".join(persons_list) if persons_list else "-"

                # مجموع ساعت کاری
                total_hours = sum(routine_hours[i] for i in matching_indices)

                routine_current_year = sum(1 for x in routine_dates if x.startswith(emsal_str))

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
                list_sheet.cell(row=row_idx, column=5, value=persons_str)    # ✅ اسامی
                list_sheet.cell(row=row_idx, column=6, value=total_hours)    # ✅ ساعت
                list_sheet.cell(row=row_idx, column=12, value=routine_current_year)
                list_sheet.cell(row=row_idx, column=13, value=routine_last_6_months)
                list_sheet.cell(row=row_idx, column=14, value=routine_current_month)
            else:
                list_sheet.cell(row=row_idx, column=4, value="-")
                list_sheet.cell(row=row_idx, column=5, value="-")
                list_sheet.cell(row=row_idx, column=6, value=0)
                list_sheet.cell(row=row_idx, column=12, value=0)
                list_sheet.cell(row=row_idx, column=13, value=0)
                list_sheet.cell(row=row_idx, column=14, value=0)

            # پردازش بیرینگ
            bearing_current_year = sum(1 for x in bearing_dates if x.startswith(emsal_str))
            bearing_last_year = sum(1 for x in bearing_dates if x.startswith(last_year_str))
            bearing_two_years_ago = sum(1 for x in bearing_dates if x.startswith(two_years_ago_str))

            list_sheet.cell(row=row_idx, column=7, value=bearing_current_year)
            list_sheet.cell(row=row_idx, column=8, value=bearing_last_year)
            list_sheet.cell(row=row_idx, column=9, value=bearing_two_years_ago)

            for idx, date in enumerate(bearing_dates[::-1], start=21):
                list_sheet.cell(row=row_idx, column=idx, value=date)

            # پردازش گریس
            grease_last_year = sum(1 for x in grease_dates if x.startswith(last_year_str))
            grease_current_year = sum(1 for x in grease_dates if x.startswith(emsal_str))

            list_sheet.cell(row=row_idx, column=10, value=grease_last_year)
            list_sheet.cell(row=row_idx, column=11, value=grease_current_year)

            for idx, date in enumerate(grease_dates[::-1], start=43):
                list_sheet.cell(row=row_idx, column=idx, value=date)

            # پردازش جعبه ترمینال
            terminal_box_current_year = sum(1 for x in terminal_box_dates if x.startswith(emsal_str))
            list_sheet.cell(row=row_idx, column=15, value=terminal_box_current_year)

            for idx, date in enumerate(terminal_box_dates[::-1], start=65):
                list_sheet.cell(row=row_idx, column=idx, value=date)

            # پردازش LCS
            lcs_current_year = sum(1 for x in lcs_dates if x.startswith(emsal_str))
            list_sheet.cell(row=row_idx, column=16, value=lcs_current_year)

            for idx, date in enumerate(lcs_dates[::-1], start=87):
                list_sheet.cell(row=row_idx, column=idx, value=date)

            # پردازش سیم‌پیچی
            wiring_last_year = sum(1 for x in wiring_dates if x.startswith(last_year_str))
            list_sheet.cell(row=row_idx, column=17, value=wiring_last_year)

            for idx, date in enumerate(wiring_dates[::-1], start=76):
                list_sheet.cell(row=row_idx, column=idx, value=date)

            # پرچس، روتین گریس، انتقال
            list_sheet.cell(row=row_idx, column=18, value="Yes" if buy_flag else "No")
            list_sheet.cell(row=row_idx, column=19, value=grease_current_year_count)
            list_sheet.cell(row=row_idx, column=20, value="Yes" if shop_flag else "-")

            row_idx += 1

        # 🔚 تمیزکاری نهایی
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

        wb.save(file_path)
        print("[OK] فایل ذخیره شد.")

    except Exception as e:
        print(f"[ERROR] خطا در باز کردن یا ذخیره فایل: {str(e)}")


# 📍 مسیر فایل اصلی
file_path = r"Z:\8-Maintenance History\All motors.xlsx"

if __name__ == "__main__":
    process_excel(file_path)
