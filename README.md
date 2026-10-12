<div align="center">

# 🌳 Receivables Management & Reconciliation System
# ماشین وصول مطالبات

**A Python desktop app + Excel engine + browser dashboard for tracking installment receivables**
**اپ دسکتاپ پایتون + موتور تطبیق اکسل + داشبورد مدیریتی برای پیگیری مطالبات اقساطی**

[English](#english) · [فارسی](#فارسی)

**Developer / توسعه‌دهنده: Ebrahim Salimi — ابراهیم سلیمی** · MIT License

</div>

---

<a id="english"></a>

## English

### Background

This project was originally written as a private tool for **Rah-e Danesh-e Aref**, a nationwide educational and counselling institute in Iran, to run its receivables collection. It is now released as a **public, generic edition** so that the finance departments of other institutes and companies can use it to make collecting receivables easier. Nothing institute-specific (names, data, logo) is left in this edition; organisation name, logo and manager are all settings.

### What it does

| Part | Purpose |
|---|---|
| 🖥️ **Desktop app** (`python/app_desktop.py`) | A simple window: pick the Excel file, run the engine, open the output, open the dashboard, change settings. No command line needed. |
| 🐍 **Engine** (`python/wosool_engine.py`) | Reconciles the follow-up list with the latest raw finance export and rebuilds every report sheet. |
| 📊 **Dashboard** (`dashboard/dashboard_modiriati_FINAL.html`) | Opens in any browser, reads the engine output and shows totals, trends, portfolio, risk tiers, expert ranking, top debtors and period-over-period changes. |

### Input workbook

One `.xlsx` file with two sheets and these columns:

`counter, term, telephone, date, price_of_mounth, state, name_family, status, coach`

| Sheet | Meaning |
|---|---|
| **Main sheet** | Your follow-up list (must be an Excel *Table*). This is what the engine updates. |
| **Initial raw financial sheet** | The latest raw export from finance. **Never modified.** |

- `term` is the unique installment/cheque id.
- `status`: empty = normal · `قرمز` = red (special collection) · anything else = an agency name (e.g. `agency 1`).
- `coach`: the expert responsible for the person.

Old institute files using `main_sheet_ebi` / `main_sheet_ahmadi` and the old Persian sheet names are still accepted and are renamed in the output.

**Start from the blank template:** [`template/input_template.xlsx`](template/input_template.xlsx). It has a **Guide** sheet that explains every column, plus the two empty sheets as Excel Tables ready for your data (hover a header to see its note). Put your current follow-up list in *Main sheet* and paste the finance export into *Initial raw financial sheet* before each run.

A filled, fully fictional example is in [`sample/sample_input.xlsx`](sample/sample_input.xlsx) (regenerate with `python sample/make_sample.py`).

### Business rules

- **Unique `term`.** A term that disappeared from the raw sheet is a cleared cheque: it is removed from *Main sheet* and appended to **Cleared checks** with the date/time it was detected. That sheet is a permanent history; nothing is ever deleted from it.
- **New terms** from the raw sheet are added to *Main sheet* and listed in **Newly added installments**.
- **Terms in both sheets are never overwritten.** Any difference (phone, date, amount, name, …) is listed in **Manual review discrepancy** for a person to decide.
- **Red has absolute priority.** If a person is red, every new installment of theirs is red too, with no expert. Only a manual edit in *Main sheet* takes someone out of red.
- **Agencies.** A new installment of an agency member gets the same agency name and the **Manager** as expert.
- **Normal people never stay without an expert.** A new installment goes to the expert the person already worked with most. A person with no history gets the expert who currently has the **fewest** students (the Manager is not a candidate). You can change it by hand afterwards.

### Output workbook

Saved in an `Excel` folder next to the input, named with the Jalali date and time. Sheets:

| Sheet | Contents |
|---|---|
| Main sheet | Updated follow-up list |
| Initial raw financial sheet | Unchanged raw export |
| Cleared checks | Cumulative history of cleared terms |
| Reds | All red installments |
| Newly added installments | Installments added in this run |
| Manual review discrepancy | Shared terms whose values differ |
| `agency N` … | One sheet per agency |
| `Expert N` …, Manager | One sheet per expert |

Every sheet except the raw sheet has a title row, a header row, `counter` (1, 2, 3, …) as the first column and `term` as the second, auto-fit column widths, row height 20 (title row 30), red/agency row colours, frozen headers and an Excel Table.

### Settings (organisation name, logo, manager)

- **Desktop app:** **⚙ Settings** button. Stored in `settings.json` next to the app. Here you set the organisation name, website, logo, the **Manager name** (the fixed expert of all agencies, exactly as written in the `coach` column) and optional alternative spellings of it.
- **Dashboard:** **⚙ تنظیمات برند** button. Organisation name, website and logo are stored in that browser and appear in print and PNG exports.

### Run from source

```bash
pip install -r python/requirements.txt
python python/app_desktop.py
```

### Build the Windows app

On Windows, double-click `python/build_exe.bat`. It installs the requirements, builds the dashboard, builds `receivables_management.exe` with PyInstaller and assembles `python/release/` (the exe + the `dashboard` folder). Give that folder to the finance team; they need nothing else installed.

If you change the dashboard sources (`dashboard-modiriati/template.html`, `app.js`, `default_logo.svg`), rebuild it with `python dashboard-modiriati/build_dashboard.py`.

### Privacy

Real Excel files are ignored by `.gitignore`. Never commit customer names, phone numbers or financial records; use the sample file for demos. The dashboard runs entirely in the browser; no data leaves the computer.

---

<a id="فارسی"></a>

<div dir="rtl">

## فارسی

### پیشینه

این پروژه ابتدا به‌صورت اختصاصی برای **موسسه‌ی آموزشی و مشاوره‌ی کشوری راه دانش عارف** نوشته شده بود. در حال حاضر به‌صورت **نسخه‌ی عمومی‌شده** در اختیار عموم قرار گرفته تا وصول مطالبات بخش‌های مالی موسسات و شرکت‌ها راحت‌تر شود. در این نسخه هیچ اسم، داده یا لوگوی اختصاصی وجود ندارد. نام سازمان، لوگو و نام مدیر همه از تنظیمات قابل تغییرند.

### اجزا

| بخش | کار |
|---|---|
| 🖥️ **اپ دسکتاپ** | یک پنجره‌ی ساده: انتخاب فایل اکسل، اجرای موتور، باز کردن خروجی و داشبورد، تنظیمات. بدون خط فرمان. |
| 🐍 **موتور** (`wosool_engine.py`) | فهرست پیگیری را با آخرین خروجی خام مالی تطبیق می‌دهد و همه‌ی شیت‌های گزارش را از نو می‌سازد. |
| 📊 **داشبورد** | در هر مرورگری باز می‌شود و خروجی موتور را به‌صورت آمار، روند، پرتفوی، سطح ریسک، رتبه‌بندی کارشناس‌ها و بدهکاران برتر نشان می‌دهد. |

### فایل ورودی

یک فایل `.xlsx` با دو شیت و این ستون‌ها:
`counter, term, telephone, date, price_of_mounth, state, name_family, status, coach`

| شیت | معنی |
|---|---|
| **Main sheet** | فهرست پیگیری شما (باید Table اکسل باشد). موتور همین را به‌روز می‌کند. |
| **Initial raw financial sheet** | آخرین خروجی خام بخش مالی. **هرگز تغییر نمی‌کند.** |

- ستون `status`: خالی یعنی عادی · `قرمز` یعنی قرمز · هر مقدار دیگر یعنی نام نمایندگی (مثل `agency 1`).
- فایل‌های قدیمی با نام‌های `main_sheet_ebi` و `main_sheet_ahmadi` هم پذیرفته می‌شوند.
- **فایل قالب خالی برای شروع:** [`template/input_template.xlsx`](template/input_template.xlsx). یک شیت **Guide** دارد که همه‌ی ستون‌ها را توضیح می‌دهد، و دو شیت خالی که به‌صورت Table اکسل آماده‌ی ورود داده‌اند (روی هر سرستون بروید تا توضیحش را ببینید). فهرست فعلی پیگیری را در Main sheet وارد کنید و قبل از هر اجرا، خروجی بخش مالی را در Initial raw financial sheet بچسبانید.
- نمونه‌ی پرشده و کاملاً ساختگی: [`sample/sample_input.xlsx`](sample/sample_input.xlsx)

### قوانین

- **term یکتاست.** termی که از شیت خام حذف شده یعنی چکش پاس شده. از Main sheet برداشته می‌شود و با تاریخ و ساعت ثبت به **Cleared checks** اضافه می‌شود. این شیت تاریخچه‌ی دائمی است و چیزی از آن پاک نمی‌شود.
- **term جدید** به Main sheet اضافه می‌شود و در **Newly added installments** هم می‌آید.
- **term مشترک هرگز overwrite نمی‌شود.** هر تفاوتی در **Manual review discrepancy** گزارش می‌شود تا یک نفر تصمیم بگیرد.
- **قرمز اولویت مطلق دارد.** قسط جدید فرد قرمز هم قرمز می‌شود و کارشناس ندارد. فقط با اصلاح دستی در Main sheet از قرمز خارج می‌شود.
- **نمایندگی:** قسط جدید عضو نمایندگی همان نام نمایندگی و کارشناس **Manager** را می‌گیرد.
- **فرد عادی هرگز بی‌کارشناس نمی‌ماند.** اگر سابقه دارد، کارشناسی که بیشترین سابقه را با او دارد. اگر ندارد، کارشناسی که الان **کمترین** دانشجو را دارد (مدیر جزو کاندیدها نیست). بعداً دستی قابل تغییر است.

### خروجی

در پوشه‌ی `Excel` کنار فایل ورودی، با تاریخ شمسی و ساعت در نام فایل. همه‌ی شیت‌ها به‌جز شیت خام: ردیف عنوان، ستون `counter` (۱، ۲، ۳، …) اول و `term` دوم، عرض ستون Auto-fit، ارتفاع ردیف ۲۰ (عنوان ۳۰)، رنگ ردیف قرمز و نمایندگی، فریز هدر و Table اکسل.

### تنظیمات

- **اپ دسکتاپ:** دکمه‌ی **⚙ تنظیمات**. نام سازمان، وب‌سایت، لوگو، **نام مدیر** (کارشناس ثابت نمایندگی‌ها، دقیقاً همان‌طور که در ستون coach می‌نویسید) و املاهای دیگر آن. در `settings.json` کنار برنامه ذخیره می‌شود.
- **داشبورد:** دکمه‌ی **⚙ تنظیمات برند**. نام، وب‌سایت و لوگو در همان مرورگر ذخیره می‌شود و در چاپ و خروجی PNG هم دیده می‌شود.

### اجرا و ساخت

- اجرا از سورس: `pip install -r python/requirements.txt` و بعد `python python/app_desktop.py`
- ساخت exe ویندوز: روی `python/build_exe.bat` دوبار کلیک کنید. پوشه‌ی `python/release` را به تیم مالی بدهید.

### حریم خصوصی

فایل‌های اکسل واقعی در `.gitignore` هستند. هرگز اسم، شماره تماس یا داده‌ی مالی واقعی را Commit نکنید. داشبورد کاملاً داخل مرورگر اجرا می‌شود و داده‌ای از کامپیوتر خارج نمی‌شود.

</div>

---

<div align="center">

MIT License · © 2026 Ebrahim Salimi — ابراهیم سلیمی

</div>
