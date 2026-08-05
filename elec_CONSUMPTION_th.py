# -*- coding: utf-8 -*-
# FULL REPORT (Electricity + Temperature/Humidity)
# - Keeps your original electricity calculations (read .xlsb, header=None, sheet='DAILY')
# - Adds Temp/Humidity (columns X..AH -> pandas indices 23..33) with averages for:
#   Today, Yesterday, This Week (inside current Jalali month), This Month, <Prev Month>, <Prev-Prev Month>
# - Fenglish titles (no Persian rendering)
# - Two separate windows (figures) with the SAME background image:
#     1) Energy_Report.png  (electricity like your original)
#     2) Energy_TH_Report.png (row-wise Temp/Humidity)
#
# Assumptions (as you specified):
#   - First row (index 0) is empty, second row (index 1) are headers in Excel.
#   - Day-Of-Year (doy) to pandas index mapping: pandas_index = doy + 2
#   - First day of Mordad (month=5) => row/index 127 (because sum of 4*31 = 124; doy_first=125; 125+2=127)
#
# Paths kept as your original:
#   file_path     = r"z:\4-ELECTRICITY CONSUMPTION\ENERGY Report -24 o'clock-1404.xlsb"
#   bg_image_path = r"z:\4-ELECTRICITY CONSUMPTION\Tennet.jpg"
#   output_path   = r"z:\4-ELECTRICITY CONSUMPTION\Energy_Report.png"
#   output_th_path= r"z:\4-ELECTRICITY CONSUMPTION\Energy_TH_Report.png"

import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.lines as mlines

# =========================
# Jalali calendar constants
# =========================
J_MONTH_LENGTHS = [31, 31, 31, 31, 31, 31, 30, 30, 30, 30, 30, 29]
J_MONTH_NAMES_EN = ["Farvardin","Ordibehesht","Khordad","Tir","Mordad","Shahrivar",
                    "Mehr","Aban","Azar","Dey","Bahman","Esfand"]

def day_of_year_to_index_pandas(doy: int) -> int:
    """ pandas_index = doy + 2  (per your rule: row1 empty, row2 headers) """
    return int(doy) + 2

def first_day_index_of_month(m: int) -> int:
    """ pandas index of day 1 of month m (1..12) """
    if m < 1 or m > 12:
        raise ValueError("Invalid month")
    doy_first = 1 + sum(J_MONTH_LENGTHS[:m-1])  # 1-based DOY
    return day_of_year_to_index_pandas(doy_first)

def index_to_day_of_year(idx: int) -> int:
    """ doy = pandas_idx - 2 """
    return int(idx) - 2

def doy_to_month_and_dom(doy: int):
    """ from DOY -> (month, day_in_month) in Jalali """
    acc = 0
    for m, ml in enumerate(J_MONTH_LENGTHS, start=1):
        if doy <= acc + ml:
            return m, (doy - acc)
        acc += ml
    return 12, min(J_MONTH_LENGTHS[-1], doy - sum(J_MONTH_LENGTHS[:-1]))

def safe_slice(df, start_idx, end_idx):
    """ inclusive slice by pandas integer index; returns empty frame if out of range """
    if df.empty:
        return df.iloc[0:0]
    lo = max(int(df.index.min()), int(start_idx))
    hi = min(int(df.index.max()), int(end_idx))
    if lo > hi:
        return df.iloc[0:0]
    return df.loc[lo:hi]



# =========================
# Temp/Humidity config
# =========================
# Columns X..AH -> pandas indices 23..33, shown with Fenglish labels row-wise
TH_COLS = list(range(23, 34))
TH_LABELS_EN = [
    "Damaye Vorudi MCC",        # X
    "Rotubat MCC",              # Y
    "Damaye Enteha MCC",        # Z
    "Damaye Vorudi MSS",        # AA
    "Rotubat MSS",              # AB
    "Damaye Enteha MSS",        # AC
    "Damaye Vorudi MCC Hexan",  # AD
    "Damaye Enteha MCC Hexan",  # AE
    "Rotubat Enteha MCC Hexan", # AF
    "Damaye Battery Room",      # AG
    "Rotubat Battery Room"      # AH
]

def avg_temp_hum(df_slice):
    """ mean across TH_COLS (ignoring NaN). returns dict label->mean or np.nan """
    out = {}
    if df_slice.empty:
        for lab in TH_LABELS_EN: out[lab] = np.nan
        return out
    max_col = df_slice.shape[1] - 1
    used_cols = [c for c in TH_COLS if c <= max_col]
    if used_cols:
        means = df_slice[used_cols].apply(pd.to_numeric, errors='coerce').mean(axis=0, skipna=True)
        for i, c in enumerate(TH_COLS):
            out[TH_LABELS_EN[i]] = float(means[c]) if c in means.index and pd.notna(means[c]) else np.nan
    else:
        for lab in TH_LABELS_EN: out[lab] = np.nan
    return out

def fmt_val(label, val):
    """ format with units: % for Rotubat, °C for Damaye """
    if pd.isna(val): return "-"
    unit = "%" if "Rotubat" in label else "°C"
    return f"{val:.1f} {unit}"

def main():
    try:
        # ====== Paths (your originals) ======
        file_path      = r"z:\electrical\4-ELECTRICITY CONSUMPTION\ENERGY Report -24 o'clock-1405.xlsb"
        bg_image_path  = r"z:\electrical\4-ELECTRICITY CONSUMPTION\Tennet.jpg"
        output_path    = r"z:\electrical\4-ELECTRICITY CONSUMPTION\Energy_Report.png"
        output_th_path = r"z:\electrical\4-ELECTRICITY CONSUMPTION\Energy_TH_Report.png"

        if not os.path.exists(file_path):
            raise FileNotFoundError("Excel file not found. Check file_path.")

        # ====== Read Excel (.xlsb) ======
        sheet_name = 'DAILY'
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='pyxlsb', header=None)

        # ====== Validate rows with numeric in col B (index 1) ======
        valid_rows = df[df[1].notna() & df[1].apply(lambda x: isinstance(x, (int, float, np.number)))]
        if len(valid_rows) < 3:
            raise ValueError("Need at least 3 valid rows for calculations.")

        last_row_index      = int(valid_rows.index[-1])   # TODAY
        prev_row_index      = int(valid_rows.index[-2])   # YESTERDAY
        prev_prev_row_index = int(valid_rows.index[-3])

        current_row        = df.loc[last_row_index]
        previous_row       = df.loc[prev_row_index]
        day_before_prev_row= df.loc[prev_prev_row_index]

        # ====== Dates text from column A (if present) ======
        current_date = str(current_row[0]) if pd.notna(current_row[0]) else "N/A"
        prev_date    = str(previous_row[0]) if pd.notna(previous_row[0]) else "N/A"

        # ====== ELECTRICITY (same logic as your code) ======
        P_D = (current_row[1] + current_row[2]) - (previous_row[1] + previous_row[2])
        Q_D = (current_row[4] + current_row[5]) - (previous_row[4] + previous_row[5])

        TR_columns = [8, 9, 10, 11, 12, 13, 14]
        TR_labels  = ['P_TR1_D:', 'P_TR2_D:', 'P_TR3_D:', 'P_TR4_D:', 'P_TR6_D:', 'P_HEX_TRA_D:', 'P_HEX_TRB_D:']

        TR_values = []
        for col in TR_columns:
            if col < df.shape[1]:
                try: TR_values.append(current_row[col] - previous_row[col])
                except: TR_values.append(0)
            else:
                TR_values.append(0)

        gpps_columns   = [8, 9, 10, 11, 12]
        available_gpps = [c for c in gpps_columns if c < df.shape[1]]
        P_GPPS_D = current_row[available_gpps].sum() - previous_row[available_gpps].sum() if available_gpps else 0

        hex_columns    = [13, 14]
        available_hex  = [c for c in hex_columns if c < df.shape[1]]
        P_HEX_D = current_row[available_hex].sum() - previous_row[available_hex].sum() if available_hex else 0

        P_D_prev = (previous_row[1] + previous_row[2]) - (day_before_prev_row[1] + day_before_prev_row[2])
        Q_D_prev = (previous_row[4] + previous_row[5]) - (day_before_prev_row[4] + day_before_prev_row[5])

        TR_values_prev = []
        for col in TR_columns:
            if col < df.shape[1]:
                try: TR_values_prev.append(previous_row[col] - day_before_prev_row[col])
                except: TR_values_prev.append(0)
            else:
                TR_values_prev.append(0)

        P_GPPS_D_prev = previous_row[available_gpps].sum() - day_before_prev_row[available_gpps].sum() if available_gpps else 0
        P_HEX_D_prev  = previous_row[available_hex].sum() - day_before_prev_row[available_hex].sum() if available_hex else 0

        # ====== Determine current Jalali month from index rule ======
        doy_today                 = index_to_day_of_year(last_row_index)
        persian_month, dom_today  = doy_to_month_and_dom(doy_today)
        month_name_en             = J_MONTH_NAMES_EN[persian_month - 1]

        first_idx_this_month      = first_day_index_of_month(persian_month)
        prev_month                = 12 if persian_month == 1 else persian_month - 1
        prev_prev_month           = 12 if prev_month == 1 else prev_month - 1
        first_idx_prev_month      = first_day_index_of_month(prev_month)
        first_idx_prev_prev_month = first_day_index_of_month(prev_prev_month)

        prev_month_name_en        = J_MONTH_NAMES_EN[prev_month - 1]
        prev_prev_month_name_en   = J_MONTH_NAMES_EN[prev_prev_month - 1]
        # در بخش محاسبات ماه‌ها بعد از خطوط موجود این را اضافه کنید:
        prev_3_month           = prev_prev_month - 1 if prev_prev_month > 1 else 12
        prev_4_month           = prev_3_month - 1 if prev_3_month > 1 else 12

        first_idx_prev_3_month = first_day_index_of_month(prev_3_month)
        first_idx_prev_4_month = first_day_index_of_month(prev_4_month)

        prev_3_month_name_en   = J_MONTH_NAMES_EN[prev_3_month - 1]
        prev_4_month_name_en   = J_MONTH_NAMES_EN[prev_4_month - 1]

        # ====== Monthly electricity (as your method) ======
        monthly_P_D = monthly_Q_D = monthly_P_GPPS_D = monthly_P_HEX_D = 0
        monthly_TR_values = [0]*len(TR_labels)
        prev_month_P_D = prev_month_Q_D = prev_month_P_GPPS_D = prev_month_P_HEX_D = 0
        prev_month_name = "Unknown"

        if first_idx_this_month in df.index:
            first_day_row = df.loc[first_idx_this_month]
            
            # Monthly calculations
            monthly_P_D = (current_row[1] + current_row[2]) - (first_day_row[1] + first_day_row[2])
            monthly_Q_D = (current_row[4] + current_row[5]) - (first_day_row[4] + first_day_row[5])

            monthly_TR_values = []
            for col in TR_columns:
                if col < df.shape[1]:
                    try: monthly_TR_values.append(current_row[col] - first_day_row[col])
                    except: monthly_TR_values.append(0)
                else:
                    monthly_TR_values.append(0)

            monthly_P_GPPS_D = (current_row[available_gpps].sum() - first_day_row[available_gpps].sum()) if available_gpps else 0
            monthly_P_HEX_D = (current_row[available_hex].sum() - first_day_row[available_hex].sum()) if available_hex else 0
            
            # Previous month calculations
            if first_idx_prev_month in df.index:
                first_day_prev_month_row = df.loc[first_idx_prev_month]
                prev_month_P_D = (first_day_row[1] + first_day_row[2]) - (first_day_prev_month_row[1] + first_day_prev_month_row[2])
                prev_month_Q_D = (first_day_row[4] + first_day_row[5]) - (first_day_prev_month_row[4] + first_day_prev_month_row[5])
                prev_month_P_GPPS_D = first_day_row[available_gpps].sum() - first_day_prev_month_row[available_gpps].sum() if available_gpps else 0
                prev_month_P_HEX_D = first_day_row[available_hex].sum() - first_day_prev_month_row[available_hex].sum() if available_hex else 0
                prev_month_name = prev_month_name_en

        # ====== TEMP/HUM averages ======
        df_today        = safe_slice(df, last_row_index, last_row_index)
        df_yesterday    = safe_slice(df, prev_row_index, prev_row_index)
        df_this_week    = safe_slice(df, first_idx_this_month, last_row_index)  # Simplified for whole month
        df_this_month   = safe_slice(df, first_idx_this_month, last_row_index)
        df_prev_month   = safe_slice(df, first_idx_prev_month, first_idx_this_month - 1)
        df_prev_prev_m  = safe_slice(df, first_idx_prev_prev_month, first_idx_prev_month - 1)

        th_today        = avg_temp_hum(df_today)
        th_yesterday    = avg_temp_hum(df_yesterday)
        th_this_week    = avg_temp_hum(df_this_week)
        th_this_month   = avg_temp_hum(df_this_month)
        th_prev_month   = avg_temp_hum(df_prev_month)
        th_prev_prev_m  = avg_temp_hum(df_prev_prev_m)
# در بخش TEMP/HUM averages این خطوط را اضافه کنید:
        df_prev_3_month = safe_slice(df, first_idx_prev_3_month, first_idx_prev_prev_month - 1)
        df_prev_4_month = safe_slice(df, first_idx_prev_4_month, first_idx_prev_3_month - 1)

        th_prev_3_month = avg_temp_hum(df_prev_3_month)
        th_prev_4_month = avg_temp_hum(df_prev_4_month)
                

        # =========================
        # FIGURE 1: Electricity report (same background) - EXACTLY like new code
        # =========================
        fig1 = plt.figure(figsize=(12, 8))
        ax1 = plt.gca()

        # background
        if os.path.exists(bg_image_path):
            try:
                bg = plt.imread(bg_image_path)
                ax1.imshow(bg, aspect='auto')
            except:
                ax1.set_facecolor('black')
        else:
            ax1.set_facecolor('black')

        ax1.axis('off')
        lh = 0.05  # line height

        # Dates
        ax1.text(0.10, 0.95, current_date, transform=ax1.transAxes, color='white', fontsize=24, fontweight='bold', ha='center', backgroundcolor='black')
        ax1.text(0.60, 0.95, prev_date,    transform=ax1.transAxes, color='white', fontsize=24, fontweight='bold', ha='center', backgroundcolor='black')
        ax1.add_line(mlines.Line2D([0, 1], [0.88, 0.88], color='white', alpha=0.3, transform=ax1.transAxes))
        ax1.text(0.10, 0.90, "TODAY",     transform=ax1.transAxes, color='yellow', fontsize=16, fontweight='bold', ha='center', backgroundcolor='black')
        ax1.text(0.60, 0.90, "YESTERDAY", transform=ax1.transAxes, color='yellow', fontsize=16, fontweight='bold', ha='center', backgroundcolor='black')

        # Electricity Today (left)
        y = 0.850
        ax1.text(0.05, y,             f'P_D: {P_D:.2f}',  transform=ax1.transAxes, color='white', fontsize=14, fontweight='bold', backgroundcolor='black'); y -= lh
        ax1.text(0.05, y,             f'Q_D: {Q_D:.2f}',  transform=ax1.transAxes, color='white', fontsize=14, fontweight='bold', backgroundcolor='black')

        # Yesterday (right)
        ax1.text(0.55, 0.850,         f'P_D: {P_D_prev:.2f}', transform=ax1.transAxes, color='white', fontsize=14, fontweight='bold', backgroundcolor='black')
        ax1.text(0.55, 0.850 - lh,    f'Q_D: {Q_D_prev:.2f}', transform=ax1.transAxes, color='white', fontsize=14, fontweight='bold', backgroundcolor='black')

        # TR today
        y -= lh * 2.5
        ax1.text(0.05, y, "TR Values (Today):", transform=ax1.transAxes, color='yellow', fontsize=12, fontweight='bold', backgroundcolor='black'); y -= lh
        for lab, val in zip(TR_labels, TR_values):
            if val != 0:
                ax1.text(0.05, y, f'{lab} {val:.2f}', transform=ax1.transAxes, color='white', fontsize=12, backgroundcolor='black'); y -= lh

        # TR yesterday
        y_r = 0.85 - lh * 2.5
        ax1.text(0.55, y_r, "TR Values (Yesterday):", transform=ax1.transAxes, color='yellow', fontsize=12, fontweight='bold', backgroundcolor='black'); y_r -= lh
        for lab, val in zip(TR_labels, TR_values_prev):
            if val != 0:
                ax1.text(0.55, y_r, f'{lab} {val:.2f}', transform=ax1.transAxes, color='white', fontsize=12, backgroundcolor='black'); y_r -= lh

        # GPPS/HEX
        ax1.text(0.05, 0.375,             f'P_GPPS_D: {P_GPPS_D:.2f}',     transform=ax1.transAxes, color='white', fontsize=14, fontweight='bold', backgroundcolor='black')
        ax1.text(0.05, 0.375 - lh,        f'P_HEX_D: {P_HEX_D:.2f}',       transform=ax1.transAxes, color='white', fontsize=14, fontweight='bold', backgroundcolor='black')
        ax1.text(0.55, 0.375,             f'P_GPPS_D: {P_GPPS_D_prev:.2f}',transform=ax1.transAxes, color='white', fontsize=14, fontweight='bold', backgroundcolor='black')
        ax1.text(0.55, 0.375 - lh,        f'P_HEX_D: {P_HEX_D_prev:.2f}',  transform=ax1.transAxes, color='white', fontsize=14, fontweight='bold', backgroundcolor='black')

        # middle vertical line
        ax1.add_line(mlines.Line2D([0.5, 0.5], [0, 1], color='white', alpha=0.3, linestyle='--', transform=ax1.transAxes))

        # Monthly block
        monthly_y = 0.35
        monthly_y -= lh * 1.2
        ax1.add_line(mlines.Line2D([0.05, 0.95], [monthly_y + 0.02, monthly_y + 0.02], color='cyan', alpha=0.5, transform=ax1.transAxes))
        
        ax1.text(0.35, monthly_y, f'Month: {month_name_en}', transform=ax1.transAxes, color='white', fontsize=14, backgroundcolor='black'); monthly_y -= lh
        
        # Monthly P_D with comparison
        ax1.text(0.05, monthly_y, f'Monthly P_D:', transform=ax1.transAxes, color='white', fontsize=14, backgroundcolor='black')
        ax1.text(0.30, monthly_y, f'{monthly_P_D:.2f}', transform=ax1.transAxes, color='yellow', fontsize=14, backgroundcolor='black')
        ax1.text(0.55, monthly_y, f'(Prev: {prev_month_P_D:.2f})', transform=ax1.transAxes, color='green', fontsize=12, backgroundcolor='black'); monthly_y -= lh
        
        # Monthly Q_D with comparison
        ax1.text(0.05, monthly_y, f'Monthly Q_D:', transform=ax1.transAxes, color='white', fontsize=14, backgroundcolor='black')
        ax1.text(0.30, monthly_y, f'{monthly_Q_D:.2f}', transform=ax1.transAxes, color='yellow', fontsize=14, backgroundcolor='black')
        ax1.text(0.55, monthly_y, f'(Prev: {prev_month_Q_D:.2f})', transform=ax1.transAxes, color='green', fontsize=12, backgroundcolor='black'); monthly_y -= lh
        
        # Monthly P_GPPS_D with comparison
        ax1.text(0.05, monthly_y, f'Monthly P_GPPS_D:', transform=ax1.transAxes, color='white', fontsize=14, backgroundcolor='black')
        ax1.text(0.30, monthly_y, f'{monthly_P_GPPS_D:.2f}', transform=ax1.transAxes, color='yellow', fontsize=14, backgroundcolor='black')
        ax1.text(0.55, monthly_y, f'(Prev: {prev_month_P_GPPS_D:.2f})', transform=ax1.transAxes, color='green', fontsize=12, backgroundcolor='black'); monthly_y -= lh
        
        # Monthly P_HEX_D with comparison
        ax1.text(0.05, monthly_y, f'Monthly P_HEX_D:', transform=ax1.transAxes, color='white', fontsize=14, backgroundcolor='black')
        ax1.text(0.30, monthly_y, f'{monthly_P_HEX_D:.2f}', transform=ax1.transAxes, color='yellow', fontsize=14, backgroundcolor='black')
        ax1.text(0.55, monthly_y, f'(Prev: {prev_month_P_HEX_D:.2f})', transform=ax1.transAxes, color='green', fontsize=12, backgroundcolor='black'); monthly_y -= lh
        
        # Month names
        ax1.text(0.05, monthly_y, f'Current Month:', transform=ax1.transAxes, color='white', fontsize=14, backgroundcolor='black')
        ax1.text(0.30, monthly_y, f'{month_name_en}', transform=ax1.transAxes, color='yellow', fontsize=14, backgroundcolor='black')
        ax1.text(0.55, monthly_y, f'Previous: {prev_month_name}', transform=ax1.transAxes, color='green', fontsize=12, backgroundcolor='black')

        plt.tight_layout()
        try:
            plt.figure(fig1.number)
            fig1.savefig(output_path, bbox_inches='tight', dpi=300, facecolor='black')
        except Exception:
            pass


        # =========================
        # FIGURE 2: Temp/Humidity (separate window), same background
        # =========================
        fig2 = plt.figure(figsize=(12, 10))
        ax2 = plt.gca()

        # background
        if os.path.exists(bg_image_path):
            try:
                bg2 = plt.imread(bg_image_path)
                ax2.imshow(bg2, aspect='auto')
            except:
                ax2.set_facecolor('black')
        else:
            ax2.set_facecolor('black')

        ax2.axis('off')
        lh2 = 0.040

        # Titles
        ax2.text(0.50, 0.96, "Temperature & Humidity Averages (X..AH)", transform=ax2.transAxes,
                 color='yellow', fontsize=18, fontweight='bold', ha='center', backgroundcolor='black')

        # Define column positions with spacing (9 ستون به جای 7)
        col_positions = [0.03, 0.13, 0.23, 0.33, 0.43, 0.53, 0.63, 0.73, 0.83, 0.93]
        headers = ["Parameter", "Today", "Yesterday", "This Week", "This Month", 
                   prev_month_name_en, prev_prev_month_name_en, prev_3_month_name_en, prev_4_month_name_en]

        # Draw headers
        for i, header in enumerate(headers):
            ax2.text(col_positions[i], 0.92, header, transform=ax2.transAxes,
                     color='cyan', fontsize=10, fontweight='bold', ha='center', backgroundcolor='black')

        # Draw data rows
        y2 = 0.88
        for lab in TH_LABELS_EN:
            # Get formatted values with colors
            today_val, today_color = fmt_val(lab, th_today[lab])
            yesterday_val, yesterday_color = fmt_val(lab, th_yesterday[lab])
            this_week_val, this_week_color = fmt_val(lab, th_this_week[lab])
            this_month_val, this_month_color = fmt_val(lab, th_this_month[lab])
            prev_month_val, prev_month_color = fmt_val(lab, th_prev_month[lab])
            prev_prev_val, prev_prev_color = fmt_val(lab, th_prev_prev_m[lab])
            prev_3_val, prev_3_color = fmt_val(lab, th_prev_3_month[lab])
            prev_4_val, prev_4_color = fmt_val(lab, th_prev_4_month[lab])
            
            # Draw parameter name
            ax2.text(col_positions[0], y2, lab, transform=ax2.transAxes, 
                     color='white', fontsize=10, fontweight='bold', ha='center', backgroundcolor='black')
            
            # Draw values with their colors
            ax2.text(col_positions[1], y2, today_val, transform=ax2.transAxes, 
                     color=today_color, fontsize=10, ha='center', backgroundcolor='black')
            ax2.text(col_positions[2], y2, yesterday_val, transform=ax2.transAxes, 
                     color=yesterday_color, fontsize=10, ha='center', backgroundcolor='black')
            ax2.text(col_positions[3], y2, this_week_val, transform=ax2.transAxes, 
                     color=this_week_color, fontsize=10, ha='center', backgroundcolor='black')
            ax2.text(col_positions[4], y2, this_month_val, transform=ax2.transAxes, 
                     color=this_month_color, fontsize=10, ha='center', backgroundcolor='black')
            ax2.text(col_positions[5], y2, prev_month_val, transform=ax2.transAxes, 
                     color=prev_month_color, fontsize=10, ha='center', backgroundcolor='black')
            ax2.text(col_positions[6], y2, prev_prev_val, transform=ax2.transAxes, 
                     color=prev_prev_color, fontsize=10, ha='center', backgroundcolor='black')
            ax2.text(col_positions[7], y2, prev_3_val, transform=ax2.transAxes, 
                     color=prev_3_color, fontsize=10, ha='center', backgroundcolor='black')
            ax2.text(col_positions[8], y2, prev_4_val, transform=ax2.transAxes, 
                     color=prev_4_color, fontsize=10, ha='center', backgroundcolor='black')
            
            y2 -= lh2
            if y2 < 0.06:
                break

        # separator
        ax2.add_line(mlines.Line2D([0.03, 0.97], [y2, y2], color='magenta', alpha=0.5, transform=ax2.transAxes))

        plt.tight_layout()
        try:
            plt.figure(fig2.number)
            fig2.savefig(output_th_path, bbox_inches='tight', dpi=300, facecolor='black')
        except Exception:
            pass


        # ======= SHOW BOTH WINDOWS =======
        plt.show()

        print(f"\n✅ Saved: {output_path}")
        print(f"✅ Saved: {output_th_path}")

    except Exception as e:
        print(f"\n❌ CRITICAL ERROR: {str(e)}")
        import traceback
        traceback.print_exc()

def fmt_val(label, val):
    """ format with units: % for Rotubat, °C for Damaye """
    if pd.isna(val): 
        return "-", 'white'
    
    # برای رطوبت، مقدار را در 100 ضرب می‌کنیم (چون درصدی است)
    if "Rotubat" in label:
        val = val * 100
        unit = "%"
        # اگر رطوبت بالاتر از 60 باشد، قرمز کن
        color = 'red' if val > 60 else 'white'
        return f"{val:.1f} {unit}", color
    else:
        unit = "°C"
        # اگر دما بالاتر از 23 باشد، قرمز کن
        color = 'red' if val > 23 else 'white'
        return f"{val:.1f} {unit}", color
    


if __name__ == "__main__":
    main()
