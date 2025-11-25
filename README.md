# محاسبه‌گر پلکانی کارمزد بیمه (Tiered-Wage-Calculator)

این نرم‌افزار یک ابزار ساده و کاربردی برای محاسبه‌ی **کارمزد پلکانی بازاریابی و صدور** بر اساس آیین‌نامه‌ها و درصدهای تعریف‌شده است.  
برنامه دارای نسخه **قابل اجرا (Portable EXE)** و همچنین **سورس کامل پایتون** می‌باشد.

---

## ✨ ویژگی‌ها
- محاسبه‌ی دقیق کارمزد پلکانی  
- پشتیبانی از کارمزد بازاریابی و صدور  
- اعمال خودکار مالیات، ضرایب بیمه خصوصی/بازرگانی  
- امکان تعیین سهم بازاریاب  
- خروجی اکسل بر اساس فایل نمونه  
- نسخه‌ی قابل اجرا بدون نیاز به نصب  
- رابط کاربری ساده و سریع  

---

## ⚠️ فایل‌های ضروری (حتماً باید در کنار برنامه باشند)

فایل‌های زیر **باید در پوشه‌ی ریشه (root)** کنار **فایل exe** یا **فایل اصلی پایتون** قرار بگیرند:

برنامه برای دسترسی به این فایل‌ها از مسیر زیر استفاده می‌کند:

```python
resource_path("template.xlsx")


Tiered-Wage-Calculator/
├── Tiered Wage V.1.2.0.py
├── Tiered Wage V.1.2.0.exe
├── template.xlsx
├── coin.ico
├── requirements.txt
├── .gitignore
├── LICENSE
└── README.md

اجرای نسخه پایتون

۱. نصب وابستگی‌ها:
pip install -r requirements.txt

۲. قرار دادن template.xlsx و coin.ico کنار فایل پایتون
۳. اجرای برنامه:
python "Tiered Wage V.1.2.0.py"

اجرای نسخه EXE

۱. فایل EXE را در یک پوشه قرار دهید
۲. فایل‌های template.xlsx و coin.ico را نیز در همان پوشه قرار دهید
۳. فایل EXE را اجرا کنید

ساخت EXE با PyInstaller
ویندوز:
pyinstaller --onefile --noconsole ^
  --add-data "template.xlsx;." ^
  --add-data "coin.ico;." ^
  --icon=coin.ico ^
  "Tiered Wage V.1.2.0.py"

لینوکس / مک:

pyinstaller --onefile --noconsole \
  --add-data "template.xlsx:." \
  --add-data "coin.ico:." \
  --icon=coin.ico \
  "Tiered Wage V.1.2.0.py"

وابستگی‌ها

openpyxl
pyperclip

کتابخانه‌ی tkinter به‌صورت پیش‌فرض همراه پایتون است و نیازی به نصب ندارد.



