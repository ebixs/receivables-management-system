# ============================================================
# ماشین وصول مطالبات — اپلیکیشن دسکتاپ
# موسسه راه دانش عارف
# توسعه‌دهنده: ابراهیم سلیمی
#
# این فایل یک رابط کاربری گرافیکی ساده (بدون نیاز به خط فرمان یا
# دانش برنامه‌نویسی) روی موتور تطبیق (wosool_engine.py) و داشبورد
# مدیریتی (فایل HTML) قرار می‌دهد.
# ============================================================

import os
import sys
import json
import subprocess
import traceback
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

from wosool_engine import run_engine


COLOR_CORAL = "#F0785A"
COLOR_CORAL_DARK = "#C85A3F"
COLOR_PAPER = "#FFF9F5"
COLOR_PANEL = "#FFFFFF"
COLOR_INK = "#2B211C"
COLOR_MUTED = "#8A7A6E"
COLOR_LINE = "#EEE0D6"
COLOR_GOLD = "#F0B400"
COLOR_RED = "#E51A1A"
COLOR_GREEN = "#2F9E6E"

FONT_TITLE = ("Tahoma", 16, "bold")
FONT_NORMAL = ("Tahoma", 10)
FONT_BUTTON = ("Tahoma", 10, "bold")
FONT_LOG = ("Consolas", 9)


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_PATH = os.path.join(app_dir(), "app_config.json")
DASHBOARD_PATH = os.path.join(app_dir(), "dashboard", "dashboard_modiriati_FINAL.html")


def load_config():
    if os.path.isfile(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def open_with_default_app(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)
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


class WosoolApp:
    def __init__(self, root):
        self.root = root
        self.cfg = load_config()

        root.title("ماشین وصول مطالبات — راه دانش عارف")
        root.geometry("760x600")
        root.configure(bg=COLOR_PAPER)
        root.minsize(640, 500)

        self._build_header()
        self._build_actions()
        self._build_log_area()
        self._build_footer()

        self._log("آماده. یکی از دکمه‌های بالا را انتخاب کنید.")

    def _build_header(self):
        header = tk.Frame(self.root, bg=COLOR_CORAL, height=90)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        tk.Label(
            header, text="ماشین وصول مطالبات", font=FONT_TITLE,
            bg=COLOR_CORAL, fg="white",
        ).pack(pady=(16, 2))
        tk.Label(
            header, text="موسسه راه دانش عارف", font=FONT_NORMAL,
            bg=COLOR_CORAL, fg="white",
        ).pack()

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

        self.btn_open_output = tk.Button(
            frame, text="📊  باز کردن آخرین خروجی در اکسل",
            command=self.on_open_last_output, **btn_style,
        )
        self.btn_open_output.grid(row=0, column=1, sticky="ew", padx=6, pady=4)

        self.btn_dashboard = tk.Button(
            frame, text="🌳  باز کردن داشبورد مدیریتی",
            command=self.on_open_dashboard, **btn_style,
        )
        self.btn_dashboard.grid(row=1, column=0, sticky="ew", padx=(0, 6), pady=4)

        self.btn_open_folder = tk.Button(
            frame, text="📁  باز کردن پوشه Excel",
            command=self.on_open_excel_folder, **btn_style,
        )
        self.btn_open_folder.grid(row=1, column=1, sticky="ew", padx=6, pady=4)

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
            insertbackground="white", wrap="word", state="disabled", height=18,
        )
        self.log_box.pack(fill="both", expand=True, pady=(4, 10))

    def _build_footer(self):
        footer = tk.Frame(self.root, bg=COLOR_PAPER, pady=8, padx=16)
        footer.pack(fill="x", side="bottom")
        tk.Label(
            footer, text="www.mmrd.ir", font=("Tahoma", 8), bg=COLOR_PAPER, fg=COLOR_MUTED,
        ).pack(side="left")
        tk.Label(
            footer, text="توسعه‌دهنده: ابراهیم سلیمی", font=("Tahoma", 8, "italic"),
            bg=COLOR_PAPER, fg=COLOR_MUTED,
        ).pack(side="right")

    def _log(self, text, tag=None):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

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
            result = run_engine(source_file)
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
        save_config(self.cfg)

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
