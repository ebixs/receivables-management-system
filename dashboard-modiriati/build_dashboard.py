"""
ساخت فایل نهایی داشبورد از روی منابع:
    template.html + app.js + default_logo.svg  →  dashboard_modiriati_FINAL.html

خروجی هم کنار همین فایل و هم در ../dashboard/ نوشته می‌شود
(پوشه‌ی dashboard همانی است که build_exe.bat کنار exe کپی می‌کند).

اجرا:  python build_dashboard.py
"""
import base64
import os

HERE = os.path.dirname(os.path.abspath(__file__))


def read(name, mode="r"):
    with open(os.path.join(HERE, name), mode, **({} if "b" in mode else {"encoding": "utf-8"})) as f:
        return f.read()


def build():
    template = read("template.html")
    app_js = read("app.js")
    logo = base64.b64encode(read("default_logo.svg", "rb")).decode("ascii")
    html = (
        template
        .replace("__DEFAULT_LOGO__", "data:image/svg+xml;base64," + logo)
        .replace("__APP_JS__", app_js)
    )
    assert "__APP_JS__" not in html and "__DEFAULT_LOGO__" not in html
    targets = [
        os.path.join(HERE, "dashboard_modiriati_FINAL.html"),
        os.path.join(HERE, "..", "dashboard", "dashboard_modiriati_FINAL.html"),
    ]
    for path in targets:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            f.write(html)
        print("written:", os.path.normpath(path))


if __name__ == "__main__":
    build()
