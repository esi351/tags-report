import tkinter as tk
from tkinter import font
import pytz
from datetime import datetime, timedelta

# متغیر جهانی برای ذخیره وضعیت "همیشه بالا"
is_always_on_top = True

# تابع برای فعال/غیرفعال کردن "همیشه بالا"
def toggle_always_on_top():
    global is_always_on_top
    is_always_on_top = not is_always_on_top  # تغییر وضعیت
    root.attributes("-topmost", is_always_on_top)  # به‌روزرسانی ویژگی
    toggle_button.config(text="▲" if is_always_on_top else "▼")  # نمادهای ساده‌تر

# تابع برای به‌روزرسانی ساعت
def update_clock():
    # تنظیم منطقه زمانی لندن
    london_time = datetime.now(pytz.timezone('Europe/London')).strftime('%H:%M:%S')
    london_label.config(text=f"London\n{london_time}")
    
    # تنظیم منطقه زمانی امریکا (نیویورک)
    new_york_now = datetime.now(pytz.timezone('America/New_York'))
    new_york_label.config(text=f"New York\n{new_york_now.strftime('%H:%M:%S')}")
    
    # محاسبه زمان شروع و پایان تایمر
    timer_start = new_york_now.replace(hour=9, minute=30, second=0, microsecond=0)  # ساعت 9:30 صبح
    pre_timer_start = timer_start - timedelta(minutes=15)  # 15 دقیقه قبل از شروع
    timer_end = timer_start + timedelta(hours=8)  # 8 ساعت بعد از شروع
    
    # بررسی وضعیت تایمر
    if new_york_now < pre_timer_start:
        # قبل از شروع پیش‌گرم
        stock_label.config(text="Stock\nNot Started", fg="#999999")
    elif pre_timer_start <= new_york_now < timer_start:
        # در حال پیش‌گرم (منفی شمارش)
        remaining_time = timer_start - new_york_now
        stock_label.config(text=f"Stock\n-{str(remaining_time).split('.')[0]}", fg="#FFA500")
    elif timer_start <= new_york_now <= timer_end:
        # در حال اجرا (شمارش مثبت)
        elapsed_time = new_york_now - timer_start
        stock_label.config(text=f"Stock\n{str(elapsed_time).split('.')[0]}", fg="#008000")
    else:
        # پایان تایمر
        stock_label.config(text="Stock\nEnd", fg="#FF0000")
    
    # به‌روزرسانی هر 100 میلی‌ثانیه (0.1 ثانیه)
    root.after(100, update_clock)

# ایجاد پنجره اصلی
root = tk.Tk()
root.title("ساعت جهانی - لندن، امریکا و Stock Timer")
root.geometry("600x200")  # عرض و ارتفاع کاهش یافته
root.configure(bg="#f0f0f0")

# تنظیم ویژگی "همیشه بالا" به صورت پیش‌فرض
root.attributes("-topmost", True)

# تنظیم فونت
custom_font = font.Font(family="Helvetica", size=14, weight="bold")  # فونت کوچک‌تر
button_font = font.Font(family="Helvetica", size=10)  # فونت کوچک‌تر برای دکمه

# برچسب ساعت لندن
london_label = tk.Label(root, text="London\n00:00:00", font=custom_font, bg="#ffffff", fg="#333333", padx=10, pady=5, relief="groove", borderwidth=2)
london_label.pack(side="left", expand=True, fill="both")

# خط جداکننده اول
separator1 = tk.Frame(root, bg="#cccccc", width=1)  # ضخامت خط کاهش یافته
separator1.pack(side="left", fill="y", padx=3)

# برچسب ساعت امریکا
new_york_label = tk.Label(root, text="New York\n00:00:00", font=custom_font, bg="#ffffff", fg="#333333", padx=10, pady=5, relief="groove", borderwidth=2)
new_york_label.pack(side="left", expand=True, fill="both")

# خط جداکننده دوم
separator2 = tk.Frame(root, bg="#cccccc", width=1)  # ضخامت خط کاهش یافته
separator2.pack(side="left", fill="y", padx=3)

# برچسب تایمر Stock
stock_label = tk.Label(root, text="Stock\nInitializing...", font=custom_font, bg="#ffffff", fg="#333333", padx=10, pady=5, relief="groove", borderwidth=2)
stock_label.pack(side="right", expand=True, fill="both")

# دکمه فعال/غیرفعال کردن "همیشه بالا"
toggle_button = tk.Button(root, text="▲", font=button_font, bg="#e0e0e0", fg="#333333", command=toggle_always_on_top, width=2, height=1)
toggle_button.place(relx=0.97, rely=0.03, anchor="ne")  # قرار دادن دکمه در گوشه بالا سمت راست

# شروع به‌روزرسانی ساعت
update_clock()

# اجرای حلقه اصلی
root.mainloop()
