# ============================================================
# ماشین وصول مطالبات — اپلیکیشن دسکتاپ (نسخه‌ی عمومی)
# Receivables Management System — desktop app (public edition)
# توسعه‌دهنده: ابراهیم سلیمی
#
# یک رابط کاربری گرافیکی ساده (بدون نیاز به خط فرمان یا دانش
# برنامه‌نویسی) روی موتور تطبیق (wosool_engine.py) و داشبورد مدیریتی.
#
# نام سازمان، لوگو و نام «مدیر» از دکمه‌ی «⚙ تنظیمات» قابل تغییر است و
# در فایل settings.json کنار برنامه ذخیره می‌شود.
# ============================================================

import os
import sys
import json
import shutil
import subprocess
import traceback
from datetime import datetime

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

from wosool_engine import run_engine, DEFAULT_SETTINGS

try:  # Pillow اختیاری است؛ بدون آن فقط لوگوی PNG/GIF نمایش داده می‌شود
    from PIL import Image, ImageTk
except Exception:  # pragma: no cover
    Image = ImageTk = None


# ============================================================
# رنگ‌ها و فونت
# ============================================================

COLOR_CORAL = "#F0785A"
COLOR_CORAL_DARK = "#C85A3F"
COLOR_PAPER = "#FFF9F5"
COLOR_PANEL = "#FFFFFF"
COLOR_INK = "#2B211C"
COLOR_MUTED = "#8A7A6E"
COLOR_LINE = "#EEE0D6"

FONT_TITLE = ("Tahoma", 16, "bold")
FONT_NORMAL = ("Tahoma", 10)
FONT_BUTTON = ("Tahoma", 10, "bold")
FONT_LOG = ("Consolas", 9)

APP_TITLE = "ماشین وصول مطالبات"
LOGO_SIZE = 60


# ============================================================
# مسیرهای برنامه (سازگار با حالت اسکریپت و حالت exe بسته‌بندی‌شده)
# ============================================================

def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_PATH = os.path.join(app_dir(), "app_config.json")      # وضعیت (آخرین فایل‌ها)
SETTINGS_PATH = os.path.join(app_dir(), "settings.json")      # تنظیمات برند و مدیر
BRAND_DIR = os.path.join(app_dir(), "brand")
DASHBOARD_PATH = os.path.join(app_dir(), "dashboard", "dashboard_modiriati_FINAL.html")

DEFAULT_APP_SETTINGS = {
    "org_name": "نام سازمان شما",
    "website": "",
    "logo_path": "",
    **DEFAULT_SETTINGS,          # manager_name, manager_aliases
}


def _load_json(path):
    if os.path.isfile(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def _save_json(path, data):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def load_settings():
    merged = dict(DEFAULT_APP_SETTINGS)
    merged.update({k: v for k, v in _load_json(SETTINGS_PATH).items() if v is not None})
    return merged


def open_with_default_app(path):
    """باز کردن یک فایل با برنامه‌ی پیش‌فرض سیستم عامل (اکسل، مرورگر و ...)."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # noqa: S606 - Windows only
        elif sys.platform == "darwin":
            subprocess.run(["open", path], check=False)
        else:
            subprocess.run(["xdg-open", path], check=False)
    except Exception as e:
        messagebox.showerror("خطا در باز کردن فایل", f"{path}\n\n{e}")


def open_folder(path):
    if not os.path.isdir(path):
        messagebox.showwarning("پوشه پیدا نشد", f"پوشه‌ی زیر هنوز وجود ندارد:\n{path}")
        return
    open_with_default_app(path)


def load_logo_image(path, size=LOGO_SIZE):
    """لوگو را برای نمایش در Tk آماده می‌کند؛ اگر نشد None."""
    if not path or not os.path.isfile(path):
        return None
    try:
        if Image is not None:
            img = Image.open(path).convert("RGBA")
            img.thumbnail((size, size))
            return ImageTk.PhotoImage(img)
        img = tk.PhotoImage(file=path)          # فقط PNG/GIF
        factor = max(1, max(img.width(), img.height()) // size)
        return img.subsample(factor, factor)
    except Exception:
        return None


# ============================================================
# پنجره‌ی تنظیمات
# ============================================================

class SettingsDialog(tk.Toplevel):
    def __init__(self, master, settings, on_save):
        super().__init__(master)
        self.title("تنظیمات")
        self.configure(bg=COLOR_PAPER, padx=18, pady=14)
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()
        self.on_save = on_save
        self.logo_path = settings.get("logo_path", "")

        def row(label, r, value, hint=""):
            tk.Label(self, text=label, font=FONT_NORMAL, bg=COLOR_PAPER, fg=COLOR_INK,
                     anchor="e").grid(row=r, column=1, sticky="e", pady=5, padx=(8, 0))
            var = tk.StringVar(value=value)
            tk.Entry(self, textvariable=var, font=FONT_NORMAL, width=36, justify="right").grid(
                row=r, column=0, sticky="ew", pady=5)
            if hint:
                tk.Label(self, text=hint, font=("Tahoma", 8), bg=COLOR_PAPER, fg=COLOR_MUTED,
                         anchor="e", justify="right").grid(row=r + 1, column=0, sticky="e")
            return var

        self.var_org = row("نام سازمان / شرکت:", 0, settings.get("org_name", ""))
        self.var_site = row("وب‌سایت (اختیاری):", 2, settings.get("website", ""))
        self.var_manager = row(
            "نام مدیر:", 4, settings.get("manager_name", ""),
            "کارشناس ثابت همه‌ی اعضای نمایندگی‌ها؛ باید عیناً همان نامی باشد که در ستون coach می‌نویسید",
        )
        self.var_aliases = row(
            "املاهای دیگر مدیر:", 6, "، ".join(settings.get("manager_aliases", [])),
            "اختیاری؛ با ویرگول جدا کنید. این املاها با نام مدیر یکی حساب می‌شوند",
        )

        tk.Label(self, text="لوگو:", font=FONT_NORMAL, bg=COLOR_PAPER, fg=COLOR_INK).grid(
            row=8, column=1, sticky="e", pady=8, padx=(8, 0))
        logo_frame = tk.Frame(self, bg=COLOR_PAPER)
        logo_frame.grid(row=8, column=0, sticky="ew", pady=8)
        self.logo_preview = tk.Label(logo_frame, bg=COLOR_PAPER)
        self.logo_preview.pack(side="right")
        tk.Button(logo_frame, text="انتخاب تصویر…", font=FONT_NORMAL,
                  command=self._pick_logo).pack(side="right", padx=8)
        tk.Button(logo_frame, text="حذف", font=FONT_NORMAL,
                  command=self._clear_logo).pack(side="right")
        self._refresh_preview()

        btns = tk.Frame(self, bg=COLOR_PAPER)
        btns.grid(row=9, column=0, columnspan=2, pady=(12, 0))
        tk.Button(btns, text="ذخیره", font=FONT_BUTTON, bg=COLOR_CORAL, fg="white",
                  activebackground=COLOR_CORAL_DARK, relief="flat", padx=18, pady=6,
                  command=self._save).pack(side="right", padx=4)
        tk.Button(btns, text="انصراف", font=FONT_NORMAL, relief="flat", padx=14, pady=6,
                  command=self.destroy).pack(side="right", padx=4)

    def _refresh_preview(self):
        self._preview_img = load_logo_image(self.logo_path, 40)
        if self._preview_img is not None:
            self.logo_preview.configure(image=self._preview_img, text="")
        else:
            self.logo_preview.configure(image="", text="(بدون لوگو)", font=("Tahoma", 8),
                                        fg=COLOR_MUTED)

    def _pick_logo(self):
        types = [("Images", "*.png *.gif *.jpg *.jpeg")] if Image is not None else [("Images", "*.png *.gif")]
        path = filedialog.askopenfilename(title="تصویر لوگو را انتخاب کنید", filetypes=types)
        if not path:
            return
        try:  # یک کپی کنار برنامه نگه داشته می‌شود تا جابه‌جایی فایل اصلی لوگو را خراب نکند
            os.makedirs(BRAND_DIR, exist_ok=True)
            dest = os.path.join(BRAND_DIR, "logo" + os.path.splitext(path)[1].lower())
            shutil.copyfile(path, dest)
            self.logo_path = dest
        except Exception:
            self.logo_path = path
        self._refresh_preview()

    def _clear_logo(self):
        self.logo_path = ""
        self._refresh_preview()

    def _save(self):
        manager = self.var_manager.get().strip() or DEFAULT_SETTINGS["manager_name"]
        aliases = [
            a.strip() for a in self.var_aliases.get().replace("،", ",").split(",") if a.strip()
        ]
        self.on_save({
            "org_name": self.var_org.get().strip(),
            "website": self.var_site.get().strip(),
            "logo_path": self.logo_path,
            "manager_name": manager,
            "manager_aliases": aliases,
        })
        self.destroy()


# ============================================================
# اپلیکیشن اصلی
# ============================================================

class WosoolApp:
    def __init__(self, root):
        self.root = root
        self.cfg = _load_json(CONFIG_PATH)
        self.settings = load_settings()

        root.geometry("760x620")
        root.configure(bg=COLOR_PAPER)
        root.minsize(640, 520)

        self._build_header()
        self._build_actions()
        self._build_log_area()
        self._build_footer()
        self._apply_brand()

        self._log("آماده. یکی از دکمه‌های بالا را انتخاب کنید.")
        if not os.path.isfile(SETTINGS_PATH):
            self._log("نکته: از دکمه‌ی «⚙ تنظیمات» نام سازمان، لوگو و نام مدیر را وارد کنید.")

    # ---------------- UI construction ----------------

    def _build_header(self):
        header = tk.Frame(self.root, bg=COLOR_CORAL, height=96)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        inner = tk.Frame(header, bg=COLOR_CORAL)
        inner.pack(expand=True)
        self.logo_label = tk.Label(inner, bg=COLOR_CORAL)
        self.logo_label.pack(side="right", padx=(0, 12))
        texts = tk.Frame(inner, bg=COLOR_CORAL)
        texts.pack(side="right")
        tk.Label(texts, text=APP_TITLE, font=FONT_TITLE, bg=COLOR_CORAL, fg="white").pack(anchor="e")
        self.org_label = tk.Label(texts, text="", font=FONT_NORMAL, bg=COLOR_CORAL, fg="white")
        self.org_label.pack(anchor="e")

    def _build_actions(self):
        frame = tk.Frame(self.root, bg=COLOR_PAPER, pady=16, padx=16)
        frame.pack(fill="x")

        btn_style = dict(
            font=FONT_BUTTON, bg=COLOR_PANEL, fg=COLOR_INK,
            activebackground=COLOR_LINE, relief="flat",
            padx=14, pady=10, cursor="hand2", bd=1,
        )

        self.btn_process = tk.Button(
            frame, text="📄  پردازش فایل اکسل جدید", command=self.on_process_new_file,
            **{**btn_style, "bg": COLOR_CORAL, "fg": "white", "activebackground": COLOR_CORAL_DARK},
        )
        self.btn_process.grid(row=0, column=0, sticky="ew", padx=(0, 6), pady=4)

        tk.Button(
            frame, text="📊  باز کردن آخرین خروجی در اکسل",
            command=self.on_open_last_output, **btn_style,
        ).grid(row=0, column=1, sticky="ew", padx=6, pady=4)

        tk.Button(
            frame, text="🌳  باز کردن داشبورد مدیریتی",
            command=self.on_open_dashboard, **btn_style,
        ).grid(row=1, column=0, sticky="ew", padx=(0, 6), pady=4)

        tk.Button(
            frame, text="📁  باز کردن پوشه Excel",
            command=self.on_open_excel_folder, **btn_style,
        ).grid(row=1, column=1, sticky="ew", padx=6, pady=4)

        tk.Button(
            frame, text="⚙  تنظیمات (نام سازمان، لوگو، مدیر)",
            command=self.on_settings, **btn_style,
        ).grid(row=2, column=0, columnspan=2, sticky="ew", pady=4)

        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=1)

    def _build_log_area(self):
        outer = tk.Frame(self.root, bg=COLOR_PAPER, padx=16)
        outer.pack(fill="both", expand=True)

        tk.Label(
            outer, text="گزارش اجرا:", font=FONT_NORMAL, bg=COLOR_PAPER, fg=COLOR_MUTED,
            anchor="w",
        ).pack(fill="x")

        self.log_box = scrolledtext.ScrolledText(
            outer, font=FONT_LOG, bg=COLOR_INK, fg="#E8E0DA",
            insertbackground="white", wrap="word", state="disabled", height=16,
        )
        self.log_box.pack(fill="both", expand=True, pady=(4, 10))

    def _build_footer(self):
        footer = tk.Frame(self.root, bg=COLOR_PAPER, pady=8, padx=16)
        footer.pack(fill="x", side="bottom")
        self.site_label = tk.Label(
            footer, text="", font=("Tahoma", 8), bg=COLOR_PAPER, fg=COLOR_MUTED,
        )
        self.site_label.pack(side="left")
        tk.Label(
            footer, text="توسعه‌دهنده: ابراهیم سلیمی", font=("Tahoma", 8, "italic"),
            bg=COLOR_PAPER, fg=COLOR_MUTED,
        ).pack(side="right")

    def _apply_brand(self):
        org = self.settings.get("org_name") or ""
        self.root.title(f"{APP_TITLE} — {org}" if org else APP_TITLE)
        self.org_label.configure(text=org)
        self.site_label.configure(text=self.settings.get("website") or "")
        self._logo_img = load_logo_image(self.settings.get("logo_path"))
        if self._logo_img is not None:
            self.logo_label.configure(image=self._logo_img)
            self.logo_label.pack(side="right", padx=(0, 12))
        else:
            self.logo_label.configure(image="")
            self.logo_label.pack_forget()

    # ---------------- logging ----------------

    def _log(self, text):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    # ---------------- actions ----------------

    def on_settings(self):
        def save(new_values):
            self.settings.update(new_values)
            if _save_json(SETTINGS_PATH, self.settings):
                self._log("✅ تنظیمات ذخیره شد.")
            else:
                messagebox.showerror("خطا", f"ذخیره‌ی تنظیمات ناموفق بود:\n{SETTINGS_PATH}")
            self._apply_brand()

        SettingsDialog(self.root, self.settings, save)

    def on_process_new_file(self):
        start_dir = self.cfg.get("last_source_dir") or app_dir()
        source_file = filedialog.askopenfilename(
            title="فایل Excel ورودی را انتخاب کنید",
            initialdir=start_dir,
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xltx *.xltm"), ("All files", "*.*")],
        )
        if not source_file:
            return

        self._clear_log()
        self._log(f"در حال پردازش:\n{source_file}\n")
        self.btn_process.configure(state="disabled", text="⏳  در حال پردازش...")
        self.root.update_idletasks()

        try:
            result = run_engine(source_file, settings={
                "manager_name": self.settings.get("manager_name"),
                "manager_aliases": self.settings.get("manager_aliases"),
            })
        except Exception as e:
            self._log("❌ خطا در پردازش فایل:\n" + "".join(traceback.format_exc()))
            messagebox.showerror("خطا", f"پردازش با خطا متوقف شد:\n\n{e}")
            self.btn_process.configure(state="normal", text="📄  پردازش فایل اکسل جدید")
            return

        self._log(result["log"])
        self._log("\n✅ پردازش با موفقیت انجام شد.")

        self.cfg["last_source_dir"] = os.path.dirname(source_file)
        self.cfg["last_output_file"] = result["output_file"]
        self.cfg["last_run_at"] = datetime.now().isoformat()
        _save_json(CONFIG_PATH, self.cfg)

        self.btn_process.configure(state="normal", text="📄  پردازش فایل اکسل جدید")

        if messagebox.askyesno(
            "پایان پردازش",
            "پردازش با موفقیت انجام شد.\nآیا می‌خواهید فایل خروجی را همین الان در اکسل باز کنید؟",
        ):
            open_with_default_app(result["output_file"])

    def on_open_last_output(self):
        last = self.cfg.get("last_output_file")
        if not last or not os.path.isfile(last):
            messagebox.showinfo(
                "خروجی‌ای یافت نشد",
                "هنوز هیچ فایلی پردازش نشده، یا آخرین خروجی جابه‌جا/حذف شده است.\n"
                "ابتدا از دکمه‌ی «پردازش فایل اکسل جدید» استفاده کنید.",
            )
            return
        open_with_default_app(last)

    def on_open_dashboard(self):
        if not os.path.isfile(DASHBOARD_PATH):
            messagebox.showerror(
                "داشبورد پیدا نشد",
                f"فایل داشبورد در مسیر زیر پیدا نشد:\n{DASHBOARD_PATH}\n\n"
                "مطمئن شوید پوشه‌ی «dashboard» کنار این برنامه قرار دارد.",
            )
            return
        open_with_default_app(DASHBOARD_PATH)

    def on_open_excel_folder(self):
        last_output = self.cfg.get("last_output_file")
        if last_output:
            folder = os.path.dirname(last_output)
        else:
            source_dir = self.cfg.get("last_source_dir") or app_dir()
            folder = os.path.join(source_dir, "Excel")
        open_folder(folder)


def main():
    root = tk.Tk()
    WosoolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
