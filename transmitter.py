from openpyxl import load_workbook
from openpyxl.styles import Font


# 🧹 تابع پاک کردن سلول‌ها (فقط ستون‌های 1 تا 12، ردیف‌های ≤ 220)
def clear_columns(sheet):
    for row in sheet.iter_rows(max_row=800, max_col=9):
        for cell in row:
            cell.value = None


# 📊 تابع پردازش فایل Excel
def process_transmitter_excel(file_path):
    # 🔁 بارگذاری فایل
    wb = load_workbook(file_path)

    # بررسی وجود شیت "List"
    if 'List' not in wb.sheetnames:
        print("Sheet 'List' not found!")
        return

    # دسترسی به شیت "List"
    list_sheet = wb['List']

    # پاک کردن فقط ستون‌های مشخص شده
    clear_columns(list_sheet)

    # تعریف سربرگ‌ها
    headers = [
        "List", "تاریخ آخرین کالیبره", "تاریخ آخرین روتین B", "تاریخ اعلام نیاز به خرید قطعه",
        "تاریخ آخرین کار تعمیراتی غیر روتین و کالیبره", "تعداد کالیبره سال 1404",
        "تعداد روتین سال 1404", "تعداد کارهای تعمیراتی غیر روتین و کالیبره 1404",
        "تعداد روتین 6 ماهه دوم سال 1404"
    ]

    # نوشتن سربرگ‌ها در شیت "List"
    for col_idx, header in enumerate(headers, start=1):
        list_sheet.cell(row=1, column=col_idx).value = header

    # متغیر سال جاری
    emsal = "1404"

    # متغیر ردیف فعلی
    row_idx = 2  # شروع از ردیف 2

    # پردازش هر شیت (غیر از شیت List)
    for sheet_name in wb.sheetnames:
        if sheet_name == 'List':
            continue

        sheet = wb[sheet_name]

        # آرایه‌ها و شمارنده‌ها
        calibration_dates = []
        routine_dates = []
        repair_dates = []
        buy_flag = False

        # تحلیل داده‌ها
        for row in sheet.iter_rows(values_only=True):
            b_value = str(row[1]) if len(row) > 1 else ""
            e_value = row[4] if len(row) > 4 else None
            f_value = row[5] if len(row) > 5 else None
            g_value = row[6] if len(row) > 6 else None
            h_value = row[7] if len(row) > 7 else None

            # کالیبره (ستون F)
            if f_value == chr(252) and b_value.startswith("14"):
                calibration_dates.append(b_value)

            # روتین B (ستون E)
            if e_value == chr(252) and b_value.startswith("14"):
                routine_dates.append(b_value)

            # خرید (ستون G)
            if g_value == chr(252):
                buy_flag = True

            # تعمیرات غیر روتین (ستون H)
            if h_value == chr(252) and b_value.startswith("14"):
                repair_dates.append(b_value)

        # ایجاد هایپرلینک
        cell = list_sheet.cell(row=row_idx, column=1)
        cell.value = sheet_name
        cell.hyperlink = f"#'{sheet_name}'!A1"
        cell.font = Font(color="0000FF", underline="single")

        # آخرین کالیبره
        if calibration_dates:
            last_calibration_date = max(calibration_dates)
            list_sheet.cell(row=row_idx, column=2).value = last_calibration_date
        else:
            list_sheet.cell(row=row_idx, column=2).value = "-"

        # آخرین روتین B
        if routine_dates:
            last_routine_date = max(routine_dates)
            list_sheet.cell(row=row_idx, column=3).value = last_routine_date
        else:
            list_sheet.cell(row=row_idx, column=3).value = "-"

        # آخرین خرید
        if buy_flag:
            last_purchase_date = calibration_dates[-1] if calibration_dates else "-"
            list_sheet.cell(row=row_idx, column=4).value = last_purchase_date
        else:
            list_sheet.cell(row=row_idx, column=4).value = "-"

        # آخرین تعمیرات غیر روتین
        if repair_dates:
            last_repair_date = max(repair_dates)
            list_sheet.cell(row=row_idx, column=5).value = last_repair_date
        else:
            list_sheet.cell(row=row_idx, column=5).value = "-"

        # شمارش تعداد
        calibration_1404 = sum(1 for x in calibration_dates if x.startswith(emsal))
        routine_1404 = sum(1 for x in routine_dates if x.startswith(emsal))
        routine_last_6_months = sum(1 for x in routine_dates if x.startswith(emsal) and int(x[5:7]) <7)
        repair_1404 = sum(1 for x in repair_dates if x.startswith(emsal))

        list_sheet.cell(row=row_idx, column=6).value = calibration_1404
        list_sheet.cell(row=row_idx, column=7).value = routine_1404
        list_sheet.cell(row=row_idx, column=8).value = repair_1404
        list_sheet.cell(row=row_idx, column=9).value = routine_last_6_months

        # افزایش ردیف
        row_idx += 1

    # ذخیره فایل
    wb.save(file_path)
    print(f"✅ پردازش تکمیل و فایل ذخیره شد: {file_path}")


# 📍 مسیر فایل اصلی
file_path = r"Z:\electrical\8-Maintenance History\ALL TRANSMITER.xlsx"

# 🚀 اجرای تابع
process_transmitter_excel(file_path)
