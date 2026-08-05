import tkinter as tk
from tkinter import ttk
import ctypes
import threading
import pystray
from PIL import Image, ImageDraw

# تابع برای دریافت وضعیت چراغ‌های کیبورد
def get_keyboard_status():
    # وضعیت Num Lock
    num_lock = bool(ctypes.windll.user32.GetKeyState(0x90) & 0x0001)
    # وضعیت Caps Lock
    caps_lock = bool(ctypes.windll.user32.GetKeyState(0x14) & 0x0001)
    # وضعیت Scroll Lock
    scroll_lock = bool(ctypes.windll.user32.GetKeyState(0x91) & 0x0001)
    return num_lock, caps_lock, scroll_lock

# تابع برای دریافت زبان فعال
def get_active_language():
    user32 = ctypes.windll.user32
    hwnd = user32.GetForegroundWindow()
    thread_id = user32.GetWindowThreadProcessId(hwnd, 0)
    klid = user32.GetKeyboardLayout(thread_id)
    lang_id = klid & (2 ** 16 - 1)  # Extract the low word (language ID)

    # تشخیص زبان بر اساس کد زبان
    if lang_id == 0x0429:  # Persian (Farsi)
        return "Persian"
    elif lang_id == 0x0409:  # English (United States)
        return "English"
    else:
        return "Unknown"

# تابع برای ایجاد تصویر آیکون
def create_icon(color, text):
    image = Image.new("RGBA", (64, 64), (255, 255, 255, 0))
    draw = ImageDraw.Draw(image)
    draw.rectangle((10, 10, 54, 54), fill=color)
    draw.text((15, 20), text, fill="white")
    return image

# تابع برای مدیریت هر آیکون به‌صورت مستقل
def manage_icon(icon_name, condition_func, on_text, off_text, on_color, off_color, title_on, title_off):
    def update_icon():
        while not stop_event.is_set():  # بررسی متغیر stop_event برای توقف حلقه
            condition = condition_func()  # بررسی شرط (مثل وضعیت Num Lock)
            if condition:
                icon.icon = create_icon(on_color, on_text)  # تنظیم تصویر و متن برای حالت روشن
                icon.title = title_on  # تنظیم عنوان برای حالت روشن
                icon.visible = True
            else:
                icon.icon = create_icon(off_color, off_text)  # تنظیم تصویر و متن برای حالت خاموش
                icon.title = title_off  # تنظیم عنوان برای حالت خاموش
                icon.visible = True
            stop_event.wait(0.1)

    # ایجاد آیکون با یک تصویر پیش‌فرض
    menu = pystray.Menu(pystray.MenuItem("Exit", lambda: stop_event.set()))
    icon = pystray.Icon(icon_name, menu=menu)
    icon.icon = create_icon(off_color, off_text)  # تصویر پیش‌فرض برای حالت غیرفعال
    icon.title = title_off  # عنوان پیش‌فرض برای حالت غیرفعال
    icon.visible = False  # ابتدا آیکون مخفی است

    # اجرای حلقه به‌روزرسانی آیکون در یک Thread جداگانه
    threading.Thread(target=update_icon, daemon=True).start()

    # اجرای آیکون
    icon.run()

# ایجاد پنجره اصلی
root = tk.Tk()
root.title("وضعیت کیبورد")
root.geometry("400x250")
root.resizable(False, False)

# عنوان
title_label = tk.Label(root, text="وضعیت چراغ‌های کیبورد و زبان", font=("Arial", 14))
title_label.pack(pady=10)

# قسمت نمایش وضعیت چراغ‌ها
frame = ttk.Frame(root)
frame.pack(pady=10)

num_lock_label = tk.Label(frame, text="Num Lock: خاموش", width=15, height=2, bg="pink", font=("Arial", 10))
num_lock_label.grid(row=0, column=0, padx=5)

caps_lock_label = tk.Label(frame, text="Caps Lock: خاموش", width=15, height=2, bg="pink", font=("Arial", 10))
caps_lock_label.grid(row=0, column=1, padx=5)

scroll_lock_label = tk.Label(frame, text="Scroll Lock: خاموش", width=15, height=2, bg="pink", font=("Arial", 10))
scroll_lock_label.grid(row=0, column=2, padx=5)

# قسمت نمایش وضعیت زبان
language_label = tk.Label(root, text="زبان: انگلیسی", font=("Arial", 12), width=20, height=2, bg="lightblue")
language_label.pack(pady=10)

# تابع به‌روزرسانی وضعیت
def update_status():
    num_lock, caps_lock, scroll_lock = get_keyboard_status()
    language = get_active_language()

    # به‌روزرسانی وضعیت چراغ‌ها
    num_lock_label.config(
        text="Num Lock: روشن" if num_lock else "Num Lock: خاموش",
        bg="lightgreen" if num_lock else "pink"
    )
    caps_lock_label.config(
        text="Caps Lock: روشن" if caps_lock else "Caps Lock: خاموش",
        bg="lightgreen" if caps_lock else "pink"
    )
    scroll_lock_label.config(
        text="Scroll Lock: روشن" if scroll_lock else "Scroll Lock: خاموش",
        bg="lightgreen" if scroll_lock else "pink"
    )

    # به‌روزرسانی وضعیت زبان
    if language == "Persian":
        language_label.config(text="زبان: فارسی", bg="lightgreen")
    elif language == "English":
        language_label.config(text="زبان: انگلیسی", bg="lightblue")
    else:
        language_label.config(text="زبان: ناشناخته", bg="gray")

    # فراخوانی مجدد تابع به‌صورت دوره‌ای
    root.after(100, update_status)

# شروع به‌روزرسانی وضعیت
update_status()

# ایجاد آیکون‌های System Tray در Threadهای جداگانه
stop_event = threading.Event()

# آیکون Num Lock
threading.Thread(
    target=manage_icon,
    args=(
        "Num Lock",
        lambda: get_keyboard_status()[0],
        "Num ON", "Num OFF",
        "orange", "red",
        "Num Lock is ON", "Num Lock is OFF"
    ),
    daemon=True
).start()

# آیکون Caps Lock
threading.Thread(
    target=manage_icon,
    args=(
        "Caps Lock",
        lambda: get_keyboard_status()[1],
        "Caps ON", "Caps OFF",
        "yellow", "blue",
        "Caps Lock is ON", "Caps Lock is OFF"
    ),
    daemon=True
).start()
 
# آیکون Scroll Lock
threading.Thread(
    target=manage_icon,
    args=(
        "Scroll Lock",
        lambda: get_keyboard_status()[2],
        "Scroll ON", "Scroll OFF",
        "green", "red",
        "Scroll Lock is ON", "Scroll Lock is OFF"
    ),
    daemon=True
).start()

# آیکون زبان
def language_condition():
    current_language = get_active_language()
    return current_language == "Persian" or current_language == "English"

threading.Thread(
    target=manage_icon,
    args=(
        "Language",
        lambda: get_active_language() == "Persian",
        "فارسی", "English",
        "white", "black",
        "Language: Persian", "Language: English"
    ),
    daemon=True
).start()

# اجرای حلقه اصلی
root.mainloop()
