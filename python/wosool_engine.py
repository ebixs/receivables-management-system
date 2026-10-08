# ============================================================
# موتور عملیاتی وصول: احمدی → ابراهیم
# نسخه ماژول (قابل فراخوانی از برنامه دسکتاپ)
#
# قوانین (بدون تغییر نسبت به نسخه اسکریپتی):
# - term یکتا
# - تطبیق دقیق نرمال‌شده نام
# - قرمز اولویت مطلق
# - قرمز/نمایندگی بدون Coach معمولی (نمایندگی همیشه "آقای احمدی")
# - عدم overwrite برای term تکراری
# - گزارش مغایرت
# - خروجی در پوشه Excel کنار فایل ورودی، با تاریخ شمسی و ساعت در اسم فایل
# ============================================================

import re
import io
import copy
import os
import contextlib
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo


def run_engine(source_file):
    """
    اجرای کامل موتور تطبیق روی یک فایل اکسل مشخص.

    ورودی:
        source_file: مسیر فایل اکسل ورودی (شامل شیت‌های main_sheet_ebi و main_sheet_ahmadi)

    خروجی: یک دیکشنری شامل
        output_file            مسیر کامل فایل خروجی ساخته‌شده
        log                    متن کامل گزارش (همان چیزی که در کنسول چاپ می‌شد)
        new_count               تعداد رکورد جدید افزوده‌شده
        duplicate_count         تعداد term مشترک/تکراری
        mismatch_count          تعداد مغایرت نیازمند بررسی دستی
        red_count               تعداد قرمز جدید
        representation_count    تعداد نمایندگی جدید
        coach_count             تعداد Coach عادی تخصیص‌یافته
    """

    SOURCE_FILE = source_file

    if not SOURCE_FILE:
        raise FileNotFoundError("هیچ فایل Excel انتخاب نشد.")

    if not os.path.isfile(SOURCE_FILE):
        raise FileNotFoundError(
            f"فایل ورودی در مسیر زیر پیدا نشد:\n{SOURCE_FILE}"
        )

    _log_buffer = io.StringIO()
    with contextlib.redirect_stdout(_log_buffer):
        def _gregorian_to_jalali(gy, gm, gd):
            """تبدیل تاریخ میلادی به شمسی؛ الگوریتم استاندارد و بدون نیاز به کتابخانه جانبی."""
            g_d_m = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334]
            if gy > 1600:
                jy = 979
                gy -= 1600
            else:
                jy = 0
                gy -= 621
            gy2 = gy + 1 if gm > 2 else gy
            days = (
                (365 * gy)
                + ((gy2 + 3) // 4)
                - ((gy2 + 99) // 100)
                + ((gy2 + 399) // 400)
                - 80
                + gd
                + g_d_m[gm - 1]
            )
            jy += 33 * (days // 12053)
            days %= 12053
            jy += 4 * (days // 1461)
            days %= 1461
            if days > 365:
                jy += (days - 1) // 365
                days = (days - 1) % 365
            if days < 186:
                jm = 1 + days // 31
                jd = 1 + (days % 31)
            else:
                jm = 7 + (days - 186) // 30
                jd = 1 + ((days - 186) % 30)
            return jy, jm, jd


        base_name = os.path.splitext(os.path.basename(SOURCE_FILE))[0]
        base_name = re.sub(r"_\d{3,4}-\d{2}-\d{2}_\d{2}-\d{2}$", "", base_name)
        _, ext = os.path.splitext(SOURCE_FILE)

        _now = datetime.now()
        _jy, _jm, _jd = _gregorian_to_jalali(_now.year, _now.month, _now.day)
        _jalali_date_str = f"{_jy}-{_jm:02d}-{_jd:02d}"
        _time_str = _now.strftime("%H-%M")

        # خروجی همیشه داخل پوشه Excel (کنار فایل ورودی) ذخیره می‌شود
        _source_dir = os.path.dirname(os.path.abspath(SOURCE_FILE))
        if os.path.basename(_source_dir) == "Excel":
            _excel_dir = _source_dir  # فایل ورودی از قبل داخل پوشه Excel است؛ پوشه تودرتو نسازیم
        else:
            _excel_dir = os.path.join(_source_dir, "Excel")
        os.makedirs(_excel_dir, exist_ok=True)

        OUTPUT_FILE = os.path.join(
            _excel_dir,
            f"{base_name}_{_jalali_date_str}_{_time_str}{ext}",
        )


        # ============================================================
        # چاپ مسیر فایل‌ها
        # ============================================================

        print("فایل ورودی:")
        print(SOURCE_FILE)

        print("\nفایل خروجی:")
        print(OUTPUT_FILE)


        # ============================================================
        # ستون‌های اصلی
        # ============================================================

        MASTER_COLS = [
            "term",
            "telephone",
            "date",
            "price_of_mounth",
            "state",
            "name_family",
            "status",
            "coach",
            "counter",
        ]


        # ============================================================
        # توابع نرمال‌سازی
        # ============================================================

        def normalize_name(value):
            s = "" if pd.isna(value) else str(value)

            s = (
                s.strip()
                .replace("ي", "ی")
                .replace("ك", "ک")
                .replace("\u200c", "")
                .replace("\u200d", "")
            )

            s = re.sub(r"\s+", "", s).casefold()

            return s


        def clean_text(value):
            return "" if pd.isna(value) else str(value).strip()


        def normalize_terms(df):
            out = df.copy()

            out["term"] = (
                pd.to_numeric(out["term"], errors="coerce")
                .fillna(0)
                .astype(int)
            )

            return out


        def field_equal(a, b, field):
            if field == "telephone":
                x = pd.to_numeric(
                    pd.Series([a]),
                    errors="coerce"
                ).iloc[0]

                y = pd.to_numeric(
                    pd.Series([b]),
                    errors="coerce"
                ).iloc[0]

                return (
                    (pd.isna(x) and pd.isna(y))
                    or (
                        not pd.isna(x)
                        and not pd.isna(y)
                        and int(x) == int(y)
                    )
                )

            if field == "price_of_mounth":
                x = pd.to_numeric(
                    pd.Series([a]),
                    errors="coerce"
                ).iloc[0]

                y = pd.to_numeric(
                    pd.Series([b]),
                    errors="coerce"
                ).iloc[0]

                return (
                    (pd.isna(x) and pd.isna(y))
                    or (
                        not pd.isna(x)
                        and not pd.isna(y)
                        and float(x) == float(y)
                    )
                )

            if field == "name_family":
                return normalize_name(a) == normalize_name(b)

            return clean_text(a) == clean_text(b)


        # ============================================================
        # تشخیص سطر عنوان (برای سازگاری با ظاهر جدید main_sheet_ebi)
        # اگر این فایل قبلاً یک‌بار توسط همین اسکریپت استایل‌دهی شده
        # باشد، سطر ۱ یک عنوان merge‌شده است و هدر واقعی ستون‌ها در
        # سطر ۲ قرار دارد. این بخش فقط محل خواندن هدر را تشخیص می‌دهد
        # و هیچ تاثیری در منطق تطبیق ندارد.
        # ============================================================

        def _sheet_has_title_row(path, sheet_name):
            probe_wb = load_workbook(path)
            try:
                probe_ws = probe_wb[sheet_name]
                return any(
                    mr.min_row == 1 and mr.max_row == 1 and mr.min_col == 1
                    for mr in probe_ws.merged_cells.ranges
                )
            finally:
                probe_wb.close()


        _ebi_has_title_row = _sheet_has_title_row(SOURCE_FILE, "main_sheet_ebi")


        # ============================================================
        # خواندن شیت‌ها
        # ============================================================

        ebi_raw = pd.read_excel(
            SOURCE_FILE,
            sheet_name="main_sheet_ebi",
            header=1 if _ebi_has_title_row else 0,
        )

        ahmadi_raw = pd.read_excel(
            SOURCE_FILE,
            sheet_name="main_sheet_ahmadi"
        )


        # ============================================================
        # نرمال‌سازی داده‌ها
        # ============================================================

        ebi = normalize_terms(ebi_raw)
        ahmadi = normalize_terms(ahmadi_raw)

        for df in (ebi, ahmadi):
            df["_name_norm"] = df["name_family"].map(normalize_name)
            df["_status"] = df["status"].map(clean_text)

            if "coach" in df.columns:
                df["_coach"] = df["coach"].map(clean_text)


        # ============================================================
        # تشخیص termهای جدید و تکراری
        # ============================================================

        ebi_terms = set(ebi.loc[ebi["term"] > 0, "term"])
        ahmadi_terms = set(ahmadi.loc[ahmadi["term"] > 0, "term"])

        # قانون جدید: term غایب از احمدی باید از ابراهیم حذف شود
        paid_terms = ebi_terms - ahmadi_terms
        new_terms = ahmadi_terms - ebi_terms
        common_terms = ebi_terms & ahmadi_terms

        new_rows = ahmadi[
            (ahmadi["term"] > 0) & ahmadi["term"].isin(new_terms)
        ].copy()

        duplicate_rows = ahmadi[
            (ahmadi["term"] > 0) & ahmadi["term"].isin(common_terms)
        ].copy()

        paid_removed_rows = ebi[
            (ebi["term"] > 0) & ebi["term"].isin(paid_terms)
        ].copy()


        # ============================================================
        # شناسایی افراد قرمز و اعضای نمایندگی‌ها
        # - قرمز: status دقیقاً "قرمز" است.
        # - نمایندگی: هر status غیرخالی و غیر از "قرمز" (مثل سیرجان یا هر
        #   نام دیگری) — یعنی نام نمایندگی می‌تواند هر چیزی باشد، نه فقط
        #   سیرجان.
        # این‌ها فقط از تاریخچه‌ی ebi خوانده می‌شوند: تغییر status همیشه
        # دستی و در ebi انجام می‌شود، نه از فایل احمدی.
        # ============================================================

        red_names = set(
            ebi.loc[
                (ebi["_status"] == "قرمز")
                & (ebi["_name_norm"] != ""),
                "_name_norm"
            ]
        )

        representation_by_name = {}
        for _name_norm, _group in ebi[ebi["_name_norm"] != ""].groupby("_name_norm"):
            _rep_statuses = [s for s in _group["_status"] if s not in ("", "قرمز")]
            if _rep_statuses:
                representation_by_name[_name_norm] = _rep_statuses[0]

        REPRESENTATION_COACH = "آقای احمدی"


        # ============================================================
        # محاسبه Coach مناسب برای هر فرد
        # فقط از رکوردهای «عادی» (بدون قرمز/نمایندگی) استفاده می‌شود؛
        # در غیر این صورت کارشناس ثابت نمایندگی‌ها (آقای احمدی) وارد
        # محاسبه‌ی تخصیص کارشناسِ افراد عادی می‌شد.
        # ============================================================

        _normal_ebi = ebi[
            (ebi["_name_norm"] != "")
            & (ebi["_coach"] != "")
            & (ebi["_status"] == "")
        ]

        coach_counts = (
            _normal_ebi
            .groupby(["_name_norm", "_coach"])
            .size()
            .reset_index(name="record_count")
        )

        coach_totals = (
            _normal_ebi
            .groupby("_coach")["_name_norm"]
            .nunique()
            .to_dict()
        )

        coach_choice = {}

        for name, group in coach_counts.groupby("_name_norm"):
            group = group.copy()

            group["_student_total"] = (
                group["_coach"]
                .map(coach_totals)
                .fillna(10**9)
            )

            group = group.sort_values(
                [
                    "record_count",
                    "_student_total",
                    "_coach",
                ],
                ascending=[
                    False,
                    True,
                    True,
                ]
            )

            coach_choice[name] = group.iloc[0]["_coach"]


        # ============================================================
        # تعیین Status و Coach برای رکوردهای جدید
        # ============================================================

        new_rows["status_final"] = ""
        new_rows["coach_final"] = ""

        for i, row in new_rows.iterrows():

            # قرمز: Status قرمز و Coach همیشه خالی (فقط با اصلاح دستی در ebi خارج می‌شود)
            if row["_name_norm"] in red_names:
                new_rows.at[i, "status_final"] = "قرمز"
                new_rows.at[i, "coach_final"] = ""

            # نمایندگی (سیرجان یا هر نمایندگی دیگر): Status = نام نمایندگی، Coach = آقای احمدی
            elif row["_name_norm"] in representation_by_name:
                new_rows.at[i, "status_final"] = representation_by_name[row["_name_norm"]]
                new_rows.at[i, "coach_final"] = REPRESENTATION_COACH

            # سایر افراد (status خالی): تخصیص Coach بر اساس سابقه
            else:
                new_rows.at[i, "status_final"] = ""
                new_rows.at[i, "coach_final"] = coach_choice.get(
                    row["_name_norm"],
                    ""
                )


        # ============================================================
        # ساخت رکوردهای جدید برای انتقال به EBI
        # ============================================================

        new_master = new_rows.copy()

        new_master["status"] = new_master["status_final"]
        new_master["coach"] = new_master["coach_final"]

        for column in MASTER_COLS:
            if column not in new_master.columns:
                new_master[column] = None

        new_master = new_master[MASTER_COLS]


        # ============================================================
        # بررسی مغایرت termهای تکراری
        # هیچ overwrite انجام نمی‌شود.
        # ============================================================

        compare_fields = [
            "telephone",
            "date",
            "price_of_mounth",
            "state",
            "name_family",
            "status",
            "coach",
        ]

        mismatches = []

        for term, group in duplicate_rows.groupby("term"):

            ebi_match = ebi[ebi["term"] == term]

            if ebi_match.empty:
                continue

            e = ebi_match.iloc[0]
            a = group.iloc[0]

            diffs = []

            for field in compare_fields:

                # اگر Status یا Coach در احمدی خالی باشد،
                # تغییر محسوب نمی‌شود.
                if (
                    field in ("status", "coach")
                    and clean_text(a[field]) == ""
                ):
                    continue

                if field == "coach":
                    # تفاوت صرفاً املایی «اقای احمدی»/«آقای احمدی» مغایرت واقعی نیست
                    e_val = e[field]
                    a_val = a[field]
                    if normalize_name(e_val) == normalize_name("اقای احمدی"):
                        e_val = REPRESENTATION_COACH
                    if normalize_name(a_val) == normalize_name("اقای احمدی"):
                        a_val = REPRESENTATION_COACH
                    if not field_equal(e_val, a_val, field):
                        diffs.append(field)
                    continue

                if not field_equal(e[field], a[field], field):
                    diffs.append(field)

            if diffs:
                mismatches.append(
                    {
                        "term": term,
                        "ebi_name_family": e["name_family"],
                        "ahmadi_name_family": a["name_family"],
                        "different_fields": "، ".join(diffs),
                    }
                )


        mismatch_df = pd.DataFrame(mismatches)


        # ============================================================
        # باز کردن فایل اصلی برای حفظ قالب
        # ============================================================

        wb = load_workbook(SOURCE_FILE)

        # ============================================================
        # چک‌های پاس‌شده حذف‌شده — حالا یک تاریخچه‌ی دائمی است، نه یک
        # عکس فوری که هر بار پاک و بازنویسی شود. رکوردهای قبلی حفظ
        # می‌شوند و رکوردهای تازه‌ی این اجرا با تاریخ/ساعت ثبت به آن‌ها
        # اضافه می‌شوند تا معلوم باشد چه کسی چه زمانی پرداختش را انجام
        # داده است.
        # ============================================================

        def _find_header_row_generic(ws_generic, target="term"):
            for r in range(1, min(5, ws_generic.max_row) + 1):
                for c in range(1, ws_generic.max_column + 1):
                    v = ws_generic.cell(r, c).value
                    if v is not None and str(v).strip().lower() == target:
                        return r
            return 1

        paid_columns = [
            "term", "name_family", "telephone", "date", "price_of_mounth",
            "reason", "تاریخ ثبت",
        ]

        _paid_history_rows = []
        if "چک‌های پاس‌شده حذف‌شده" in wb.sheetnames:
            _existing_ws = wb["چک‌های پاس‌شده حذف‌شده"]
            _hdr_row = _find_header_row_generic(_existing_ws, "term")
            _headers = [
                (_existing_ws.cell(_hdr_row, c).value or "")
                for c in range(1, _existing_ws.max_column + 1)
            ]
            for r in range(_hdr_row + 1, _existing_ws.max_row + 1):
                _vals = [
                    _existing_ws.cell(r, c).value
                    for c in range(1, len(_headers) + 1)
                ]
                if all(v is None for v in _vals):
                    continue
                _paid_history_rows.append(dict(zip(_headers, _vals)))
            _paid_sheet_position = wb.sheetnames.index("چک‌های پاس‌شده حذف‌شده")
            del wb["چک‌های پاس‌شده حذف‌شده"]
        else:
            _paid_sheet_position = None

        _registered_at = f"{_jalali_date_str} {_time_str}"

        _new_paid_rows = []
        for _, row in paid_removed_rows.iterrows():
            _new_paid_rows.append({
                "term": row.get("term"),
                "name_family": row.get("name_family"),
                "telephone": row.get("telephone"),
                "date": row.get("date"),
                "price_of_mounth": row.get("price_of_mounth"),
                "reason": "term در main_sheet_ahmadi وجود ندارد؛ حذف از EBI",
                "تاریخ ثبت": _registered_at,
            })

        _seen_paid_terms = set()
        _all_paid_rows = []
        for _r in (_paid_history_rows + _new_paid_rows):
            _t = _r.get("term")
            if _t in _seen_paid_terms:
                continue
            _seen_paid_terms.add(_t)
            _all_paid_rows.append(_r)

        if _paid_sheet_position is not None:
            ws_paid = wb.create_sheet("چک‌های پاس‌شده حذف‌شده", _paid_sheet_position)
        else:
            ws_paid = wb.create_sheet("چک‌های پاس‌شده حذف‌شده")

        for col_idx, col_name in enumerate(paid_columns, 1):
            ws_paid.cell(1, col_idx).value = col_name

        for row_idx, row_dict in enumerate(_all_paid_rows, 2):
            for col_idx, col_name in enumerate(paid_columns, 1):
                value = row_dict.get(col_name)
                if value is not None and isinstance(value, float) and pd.isna(value):
                    value = None
                ws_paid.cell(row_idx, col_idx).value = value


        ws = wb["main_sheet_ebi"]


        # ============================================================
        # پیدا کردن جدول اصلی
        # ============================================================

        if not ws.tables:
            raise ValueError(
                "در شیت main_sheet_ebi هیچ جدول Excel پیدا نشد."
            )

        table = next(iter(ws.tables.values()))

        # سطر هدر واقعی main_sheet_ebi: اگر سطر عنوان از قبل وجود دارد سطر ۲،
        # در غیر این صورت (اولین اجرا روی فایل خام) سطر ۱.
        # این فقط محل نوشتن/خواندن را مشخص می‌کند و هیچ اثری روی منطق ندارد.
        EBI_HEADER_ROW = 2 if _ebi_has_title_row else 1


        # بازسازی کامل main_sheet_ebi طبق لیست احمدی
        existing_ebi = ebi[ebi["term"] > 0].copy()
        kept_ebi = existing_ebi[existing_ebi["term"].isin(common_terms)].copy()
        kept_ebi = kept_ebi.drop_duplicates("term", keep="first")
        new_master = new_master.drop_duplicates("term", keep="first")

        final_ebi = pd.concat(
            [kept_ebi[MASTER_COLS], new_master[MASTER_COLS]],
            ignore_index=True
        )

        # یکسان‌سازی املای «آقای احمدی» در همه‌جا (فقط املای متن؛ منطق تطبیق دست‌نخورده)
        _ahmadi_spelling_keys = {
            normalize_name("اقای احمدی"),
            normalize_name("آقای احمدی"),
        }

        def _normalize_coach_spelling(value):
            if pd.isna(value):
                return value
            if normalize_name(value) in _ahmadi_spelling_keys:
                return REPRESENTATION_COACH
            return value

        final_ebi["coach"] = final_ebi["coach"].map(_normalize_coach_spelling)

        # حذف تمام ردیف‌های قبلی و نوشتن فقط ردیف‌های مجاز
        if ws.max_row > EBI_HEADER_ROW:
            ws.delete_rows(EBI_HEADER_ROW + 1, ws.max_row - EBI_HEADER_ROW)

        for _, row in final_ebi.iterrows():
            excel_row = ws.max_row + 1
            for column_index, column_name in enumerate(MASTER_COLS, 1):
                value = None if pd.isna(row[column_name]) else row[column_name]
                cell = ws.cell(excel_row, column_index)
                cell.value = value

                if excel_row > EBI_HEADER_ROW + 1:
                    source_cell = ws.cell(excel_row - 1, column_index)
                    if source_cell.has_style:
                        cell._style = copy.copy(source_cell._style)
                    cell.number_format = source_cell.number_format
                    cell.alignment = copy.copy(source_cell.alignment)
                    cell.font = copy.copy(source_cell.font)
                    cell.fill = copy.copy(source_cell.fill)
                    cell.border = copy.copy(source_cell.border)
                    cell.protection = copy.copy(source_cell.protection)

        # ============================================================
        # مرتب‌سازی جدول اصلی بر اساس نام و سپس term
        # ============================================================

        rows = list(
            ws.iter_rows(
                min_row=EBI_HEADER_ROW + 1,
                max_row=ws.max_row,
                max_col=ws.max_column
            )
        )

        snapshot = [
            (
                [cell.value for cell in row],
                [copy.copy(cell._style) for cell in row]
            )
            for row in rows
        ]


        def sort_key(item):
            values = item[0]

            name_value = (
                values[5]
                if len(values) > 5
                else ""
            )

            term_value = (
                values[0]
                if len(values) > 0
                else 0
            )

            term_number = (
                pd.to_numeric(
                    pd.Series([term_value]),
                    errors="coerce"
                )
                .fillna(0)
                .iloc[0]
            )

            return (
                normalize_name(name_value),
                int(term_number)
            )


        snapshot.sort(key=sort_key)


        for row_index, (values, styles) in enumerate(
            snapshot,
            EBI_HEADER_ROW + 1
        ):

            for column_index, (value, style) in enumerate(
                zip(values, styles),
                1
            ):

                destination_cell = ws.cell(
                    row_index,
                    column_index
                )

                destination_cell.value = value
                destination_cell._style = copy.copy(style)


        # ============================================================
        # به‌روزرسانی محدوده جدول
        # ============================================================

        table.ref = (
            f"A{EBI_HEADER_ROW}:"
            f"{get_column_letter(ws.max_column)}"
            f"{ws.max_row}"
        )


        # ============================================================
        # بازسازی کامل شیت‌های ثانویه — هر بار از روی نتیجه‌ی نهایی شیت ابراهیم
        #   • کارشناسان (یک شیت برای هر کارشناس، از جمله «آقای احمدی»)
        #   • قرمزها و red
        #   • نمایندگی‌ها (سیرجان یا هر نام دیگر)
        #   • اقساط جدید اضافه شده / مغایرت بررسی دستی
        # این شیت‌ها فقط نمایی از نتیجه‌ی همین اجرا هستند و هر بار از نو ساخته
        # می‌شوند؛ جایگاه تب‌ها حفظ می‌شود. هیچ اثری روی منطق تطبیق بالا ندارد.
        # (شیت «چک‌های پاس‌شده حذف‌شده» جداست و تاریخچه‌ی تجمعی است.)
        # ============================================================

        _PAID_SHEET = "چک‌های پاس‌شده حذف‌شده"
        _NEW_ADDED_SHEET = "اقساط جدید اضافه شده"
        _MISMATCH_SHEET = "مغایرت بررسی دستی"
        _RED_SHEET = "قرمزها"
        _RED_LEGACY_SHEET = "red"
        _RESERVED_SHEETS = {
            "main_sheet_ebi", "main_sheet_ahmadi", _PAID_SHEET,
            _NEW_ADDED_SHEET, _MISMATCH_SHEET, _RED_SHEET, _RED_LEGACY_SHEET,
        }
        # نام‌های قدیمی هم‌ارز که باید در همان جایگاه با نام درست جایگزین شوند
        _SHEET_ALIASES = {REPRESENTATION_COACH: ["اقای احمدی"]}
        _RED_LEGACY_DEFAULT_LABELS = [
            "ردیف", "تاریخ دریافت", "تاریخ", "مبلغ", "مقطع", "نام و نام خانوادگی", "وضعیت",
        ]

        _sheet_report = []      # (نام شیت، تعداد رکورد)
        _emptied_sheets = []    # شیت‌هایی که دیگر رکوردی ندارند
        _written_names = set()

        def _safe_sheet_name(name):
            cleaned = re.sub(r"[\[\]\:\*\?\/\\]", " ", str(name)).strip().strip("'")
            return cleaned[:31] or "sheet"

        def _unique_sheet_name(base):
            name = base
            n = 2
            while name in _written_names or name in _RESERVED_SHEETS:
                suffix = f" ({n})"
                name = base[: 31 - len(suffix)] + suffix
                n += 1
            return name

        def _has_title_row(ws_any):
            return any(
                mr.min_row == 1 and mr.max_row == 1 and mr.min_col == 1
                for mr in ws_any.merged_cells.ranges
            )

        def _existing_header_labels(sheet_name):
            if sheet_name not in wb.sheetnames:
                return None
            ws_any = wb[sheet_name]
            hdr_row = 2 if _has_title_row(ws_any) else 1
            labels = []
            for c in range(1, ws_any.max_column + 1):
                v = ws_any.cell(hdr_row, c).value
                if v is None:
                    break
                labels.append(str(v))
            return labels or None

        def _looks_like_list_sheet(ws_any):
            needed = {"term", "name_family", "status", "coach"}
            for r in range(1, min(3, ws_any.max_row) + 1):
                row_vals = {
                    str(ws_any.cell(r, c).value).strip().lower()
                    for c in range(1, ws_any.max_column + 1)
                    if ws_any.cell(r, c).value is not None
                }
                if needed <= row_vals:
                    return True
            return False

        def _records_for(df_subset, columns):
            if df_subset.empty:
                return []
            ordered = df_subset.assign(
                _sort_name=df_subset["name_family"].map(normalize_name),
                _sort_term=pd.to_numeric(df_subset["term"], errors="coerce").fillna(0),
            ).sort_values(["_sort_name", "_sort_term"], kind="stable")
            return [
                [None if pd.isna(row.get(col)) else row.get(col) for col in columns]
                for _, row in ordered.iterrows()
            ]

        def _replace_sheet(sheet_name, headers, records, aliases=()):
            """حذف شیت قدیمی (و نام‌های هم‌ارز) و ساخت دوباره‌ی آن در همان جایگاه تب‌ها."""
            names_to_drop = [n for n in [sheet_name, *aliases] if n in wb.sheetnames]
            position = min((wb.sheetnames.index(n) for n in names_to_drop), default=None)
            for n in names_to_drop:
                del wb[n]
            if position is not None:
                ws_new = wb.create_sheet(sheet_name, position)
            else:
                ws_new = wb.create_sheet(sheet_name)
            for col_idx, label in enumerate(headers, 1):
                ws_new.cell(1, col_idx).value = label
            for row_idx, record in enumerate(records, 2):
                for col_idx, value in enumerate(record, 1):
                    ws_new.cell(row_idx, col_idx).value = value
            _written_names.add(sheet_name)
            return ws_new

        _final_status = final_ebi["status"].fillna("").astype(str).str.strip()
        _final_coach = final_ebi["coach"].fillna("").astype(str).str.strip()

        # ---- قرمزها ----
        _red_df = final_ebi[_final_status == "قرمز"]
        _replace_sheet(_RED_SHEET, MASTER_COLS, _records_for(_red_df, MASTER_COLS))
        _sheet_report.append((_RED_SHEET, len(_red_df)))

        # ---- red (شیت قدیمی با عنوان‌های فارسی؛ فقط اگر وجود داشته باشد) ----
        if _RED_LEGACY_SHEET in wb.sheetnames:
            _red_labels = _existing_header_labels(_RED_LEGACY_SHEET)
            if not _red_labels or len(_red_labels) != 7:
                _red_labels = _RED_LEGACY_DEFAULT_LABELS
            _replace_sheet(
                _RED_LEGACY_SHEET, _red_labels, _records_for(_red_df, MASTER_COLS[:7])
            )
            _sheet_report.append((_RED_LEGACY_SHEET, len(_red_df)))

        # ---- اقساط جدید اضافه شده (رکوردهای افزوده‌شده در همین اجرا) ----
        _new_added_df = final_ebi[final_ebi["term"].isin(set(new_master["term"]))]
        _replace_sheet(
            _NEW_ADDED_SHEET,
            ["نوع عملیات"] + MASTER_COLS,
            [["افزودن به main_sheet_ebi"] + rec for rec in _records_for(_new_added_df, MASTER_COLS)],
        )
        _sheet_report.append((_NEW_ADDED_SHEET, len(_new_added_df)))

        # ---- مغایرت بررسی دستی (همان مغایرت‌های همین اجرا) ----
        _ebi_first = ebi.drop_duplicates("term", keep="first").set_index("term")
        _ahmadi_first = duplicate_rows.drop_duplicates("term", keep="first").set_index("term")

        def _pick(frame, term_value, column):
            if term_value in frame.index:
                value = frame.at[term_value, column]
                return None if pd.isna(value) else value
            return None

        _mismatch_records = []
        for _m in mismatches:
            _t = _m["term"]
            _mismatch_records.append([
                _t,
                _m["ebi_name_family"],
                _m["ahmadi_name_family"],
                _m["different_fields"],
                _pick(_ebi_first, _t, "telephone"),
                _pick(_ahmadi_first, _t, "telephone"),
                _pick(_ebi_first, _t, "name_family"),
                _pick(_ahmadi_first, _t, "name_family"),
            ])
        _replace_sheet(
            _MISMATCH_SHEET,
            ["term", "ebi_name_family", "ahmadi_name_family", "different_fields",
             "ebi_telephone", "ahmadi_telephone", "ebi_name", "ahmadi_name"],
            _mismatch_records,
        )
        _sheet_report.append((_MISMATCH_SHEET, len(_mismatch_records)))

        # ---- نمایندگی‌ها: یک شیت برای هر status غیرخالی و غیر از قرمز ----
        _rep_names = sorted(s for s in _final_status.unique() if s not in ("", "قرمز"))
        for _rep in _rep_names:
            _rep_df = final_ebi[_final_status == _rep]
            _rep_sheet = _safe_sheet_name(_rep)
            if _rep_sheet in _written_names or _rep_sheet in _RESERVED_SHEETS:
                _rep_sheet = _unique_sheet_name(_rep_sheet)
            _replace_sheet(_rep_sheet, MASTER_COLS, _records_for(_rep_df, MASTER_COLS))
            _sheet_report.append((_rep_sheet, len(_rep_df)))

        # ---- کارشناسان: یک شیت برای هر کارشناس (بر اساس ستون coach) ----
        _coach_names = set(c for c in _final_coach.unique() if c != "")
        for _canonical, _alias_names in _SHEET_ALIASES.items():
            if any(a in wb.sheetnames for a in _alias_names):
                _coach_names.add(_canonical)   # حتی اگر رکوردی ندارد، شیت قدیمی باید جایگزین شود
        for _coach in sorted(_coach_names):
            _coach_df = final_ebi[_final_coach == _coach]
            _coach_sheet = _safe_sheet_name(_coach)
            if _coach_sheet in _written_names or _coach_sheet in _RESERVED_SHEETS:
                _coach_sheet = _unique_sheet_name(_coach_sheet)
            _replace_sheet(
                _coach_sheet, MASTER_COLS, _records_for(_coach_df, MASTER_COLS),
                aliases=_SHEET_ALIASES.get(_coach, ()),
            )
            _sheet_report.append((_coach_sheet, len(_coach_df)))
            if _coach_df.empty:
                _emptied_sheets.append(_coach_sheet)

        # ---- شیت‌های لیستی قدیمی که دیگر نه کارشناس/نمایندگی فعلی‌اند: خالی می‌شوند ----
        for _stale in list(wb.sheetnames):
            if _stale in _RESERVED_SHEETS or _stale in _written_names:
                continue
            if _looks_like_list_sheet(wb[_stale]):
                _replace_sheet(_stale, MASTER_COLS, [])
                _emptied_sheets.append(_stale)


        # ============================================================
        # گزارش نهایی
        # ============================================================

        print()
        print("=" * 60)
        print("گزارش اجرای موتور وصول")
        print("=" * 60)

        print(
            f"رکورد جدید برای افزودن: {len(new_master)}"
        )

        print(
            f"term تکراری و بدون انتقال: {len(duplicate_rows)}"
        )

        print(
            f"مغایرت‌های نیازمند بررسی دستی: {len(mismatch_df)}"
        )

        red_count = (new_master["status"] == "قرمز").sum()
        representation_count = (
            (new_master["status"] != "")
            & (new_master["status"] != "قرمز")
        ).sum()
        coach_count = (
            (new_master["status"] == "")
            & new_master["coach"].fillna("").astype(str).str.strip().ne("")
        ).sum()

        print(
            f"قرمز جدید: {red_count} | "
            f"نمایندگی جدید: {representation_count} | "
            f"Coach عادی تخصیص‌یافته: {coach_count}"
        )

        if representation_count:
            rep_breakdown = (
                new_master.loc[
                    (new_master["status"] != "") & (new_master["status"] != "قرمز"),
                    "status",
                ]
                .value_counts()
            )
            for rep_name, rep_n in rep_breakdown.items():
                print(f"  - نمایندگی «{rep_name}»: {rep_n} رکورد جدید")

        print()
        print("-" * 60)
        print("چک‌های پاس‌شده (تاریخچه‌ی تجمعی):")
        print(
            f"  {len(_new_paid_rows)} مورد جدید در این اجرا ثبت شد | "
            f"مجموع تاریخچه: {len(_all_paid_rows)} مورد"
        )

        print()
        print("شیت‌های ثانویه از روی نتیجه‌ی نهایی شیت ابراهیم به‌روز شدند:")
        for _name, _count in _sheet_report:
            print(f"  - {_name}: {_count} رکورد")
        for _name in _emptied_sheets:
            print(f"  - {_name}: دیگر رکوردی ندارد (خالی شد)")
