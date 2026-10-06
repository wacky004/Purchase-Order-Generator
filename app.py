import base64
import calendar
import datetime
import html
import json
import os
import pathlib
import re
import shutil
import subprocess
import sys
import tempfile
import time
import tkinter as tk
from tkinter import messagebox, ttk

FROZEN = getattr(sys, "frozen", False)
RESOURCE_DIR = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
APP_DIR = os.path.dirname(sys.executable) if FROZEN else RESOURCE_DIR
TEMPLATE_PATH = os.path.join(RESOURCE_DIR, "template.html")
LOGO_PATH = os.path.join(RESOURCE_DIR, "assets", "logo.png")
ICON_PATH = os.path.join(RESOURCE_DIR, "assets", "icon.png")


def _known_folder_documents():
    try:
        import ctypes
        from ctypes import wintypes

        class GUID(ctypes.Structure):
            _fields_ = [
                ("Data1", wintypes.DWORD),
                ("Data2", wintypes.WORD),
                ("Data3", wintypes.WORD),
                ("Data4", ctypes.c_byte * 8),
            ]

        folder_id = GUID(
            0xFDD39AD0,
            0x238F,
            0x46AF,
            (ctypes.c_byte * 8)(0xAD, 0xB4, 0x6C, 0x85, 0x48, 0x03, 0x69, 0xC7),
        )
        path_pointer = ctypes.c_wchar_p()
        result = ctypes.windll.shell32.SHGetKnownFolderPath(
            ctypes.byref(folder_id), 0, None, ctypes.byref(path_pointer)
        )
        if result == 0 and path_pointer.value:
            path = path_pointer.value
            ctypes.windll.ole32.CoTaskMemFree(path_pointer)
            return path
    except Exception:
        return None
    return None


def resolve_library_dir():
    base = _known_folder_documents() or os.path.join(os.path.expanduser("~"), "Documents")
    try:
        os.makedirs(base, exist_ok=True)
    except OSError:
        base = APP_DIR
    return os.path.join(base, "Takara Purchase Orders")


LIBRARY_DIR = resolve_library_dir()
ORDERS_DIR = os.path.join(LIBRARY_DIR, "orders")
DOCUMENTS_DIR = os.path.join(LIBRARY_DIR, "documents")
BROWSER_PROFILE_DIR = os.path.join(LIBRARY_DIR, ".browser")
RENDER_LOG_PATH = os.path.join(LIBRARY_DIR, "po_render.log")
SINGLE_PAGE_MAX_ITEMS = 14
SINGLE_TOTAL_ROWS = 19
CREATE_NO_WINDOW = 0x08000000
DEFAULT_APPROVER = "Mrs. Teresita Alcera - VP Finance"


def find_browsers():
    candidates = [
        (
            "Edge",
            [
                r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            ],
            ("msedge",),
        ),
        (
            "Chrome",
            [
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                os.path.join(
                    os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "Application", "chrome.exe"
                ),
            ],
            ("chrome",),
        ),
    ]
    browsers = []
    for display_name, paths, which_names in candidates:
        found = None
        for path in paths:
            if path and os.path.isfile(path):
                found = path
                break
        if not found:
            for which_name in which_names:
                found = shutil.which(which_name)
                if found:
                    break
        if found:
            browsers.append((display_name, found))
    return browsers


def browser_profile_dir(browser_name):
    return os.path.join(BROWSER_PROFILE_DIR, browser_name.lower())


def log_render(message):
    try:
        os.makedirs(LIBRARY_DIR, exist_ok=True)
        stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(RENDER_LOG_PATH, "a", encoding="utf-8") as handle:
            handle.write(f"[{stamp}] {message}\n")
    except OSError:
        pass


def cleanup_stale_render_dirs():
    temp_root = tempfile.gettempdir()
    try:
        names = os.listdir(temp_root)
    except OSError:
        return
    for name in names:
        if not (name.startswith("po_edge_") or name.startswith("po_render_")):
            continue
        path = os.path.join(temp_root, name)
        for _attempt in range(3):
            shutil.rmtree(path, ignore_errors=True)
            if not os.path.exists(path):
                break
            time.sleep(0.3)


def _prepare_target(path):
    if not os.path.exists(path):
        return True
    try:
        os.remove(path)
        return True
    except OSError:
        return False


def valid_pdf(path):
    try:
        if os.path.getsize(path) < 100:
            return False
        with open(path, "rb") as handle:
            return handle.read(5) == b"%PDF-"
    except OSError:
        return False


def valid_png(path):
    try:
        if os.path.getsize(path) < 100:
            return False
        with open(path, "rb") as handle:
            return handle.read(8) == b"\x89PNG\r\n\x1a\n"
    except OSError:
        return False


def esc(value):
    return html.escape("" if value is None else str(value), quote=False)


def fmt_money(value):
    return f"{value:,.2f}"


def fmt_qty(value):
    if value is None:
        return ""
    if abs(value - round(value)) < 1e-9:
        return f"{int(round(value)):,}"
    return f"{value:,.2f}"


def parse_number(text):
    text = str(text).replace(",", "").strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def safe_filename(text):
    text = re.sub(r"[^A-Za-z0-9._-]+", "_", str(text).strip())
    return text.strip("._") or "document"


def order_path(file_name):
    return os.path.join(ORDERS_DIR, file_name)


def order_file_name(po_no, name):
    return f"PO_{safe_filename(po_no)}_{safe_filename(name)}.json"


def next_copy_file_name(po_no, name):
    base = f"PO_{safe_filename(po_no)}_{safe_filename(name)}"
    number = 2
    while os.path.exists(order_path(f"{base} ({number}).json")):
        number += 1
    return f"{base} ({number}).json"


def order_copy_suffix(file_name):
    match = re.search(r" \((\d+)\)\.json$", str(file_name))
    return f" ({match.group(1)})" if match else ""


def read_order(file_name):
    try:
        with open(order_path(file_name), encoding="utf-8") as handle:
            data = json.load(handle)
    except (OSError, ValueError):
        return None
    return data if isinstance(data, dict) else None


def write_order(order, file_name):
    os.makedirs(ORDERS_DIR, exist_ok=True)
    with open(order_path(file_name), "w", encoding="utf-8") as handle:
        json.dump(order, handle, indent=2, ensure_ascii=False)


def delete_order(file_name):
    try:
        os.remove(order_path(file_name))
        return True
    except OSError:
        return False


def document_paths(file_name):
    base = file_name[:-5] if file_name.lower().endswith(".json") else file_name
    return (
        os.path.join(DOCUMENTS_DIR, base + ".pdf"),
        os.path.join(DOCUMENTS_DIR, base + ".png"),
    )


def has_document(file_name):
    pdf_path, _png_path = document_paths(file_name)
    return os.path.isfile(pdf_path)


def delete_order_documents(file_name):
    for path in document_paths(file_name):
        try:
            os.remove(path)
        except OSError:
            pass


def migrate_legacy_orders():
    legacy_dir = os.path.join(APP_DIR, "orders")
    if not os.path.isdir(legacy_dir) or os.path.abspath(legacy_dir) == os.path.abspath(ORDERS_DIR):
        return
    legacy_output = os.path.join(APP_DIR, "output")
    try:
        os.makedirs(ORDERS_DIR, exist_ok=True)
        for file_name in os.listdir(legacy_dir):
            if not file_name.lower().endswith(".json"):
                continue
            target = order_path(file_name)
            if os.path.exists(target):
                continue
            shutil.copy2(os.path.join(legacy_dir, file_name), target)
            for source in document_paths(file_name):
                legacy_source = os.path.join(legacy_output, os.path.basename(source))
                if os.path.isfile(legacy_source):
                    os.makedirs(DOCUMENTS_DIR, exist_ok=True)
                    shutil.copy2(legacy_source, source)
    except OSError:
        pass


def normalize_order(data, existing=None):
    now = datetime.datetime.now().isoformat(timespec="seconds")
    order = {
        "schema": 1,
        "po_no": data.get("po_no", ""),
        "name": data.get("name", ""),
        "order_date": data.get("order_date", ""),
        "delivery_date": data.get("delivery_date", ""),
        "items": data.get("items", []),
        "approved_by": data.get("approved_by", ""),
        "approved_date": data.get("approved_date", ""),
        "verified_by": data.get("verified_by", ""),
        "verified_date": data.get("verified_date", ""),
        "received_by": data.get("received_by", ""),
        "received_date": data.get("received_date", ""),
        "created": (existing or {}).get("created") or now,
        "modified": now,
    }
    document = data.get("document") or (existing or {}).get("document")
    if document:
        order["document"] = document
    return order


def order_total(order):
    return sum(
        amount for amount in (item_amount(item) for item in order.get("items", [])) if amount is not None
    )


def list_orders():
    entries = []
    if not os.path.isdir(ORDERS_DIR):
        return entries
    for file_name in os.listdir(ORDERS_DIR):
        if not file_name.lower().endswith(".json"):
            continue
        order = read_order(file_name)
        if order is None:
            continue
        entries.append(
            {
                "file": file_name,
                "order": order,
                "total": order_total(order),
                "modified": order.get("modified", ""),
            }
        )
    entries.sort(key=lambda entry: entry["modified"], reverse=True)
    return entries


def parse_date_text(text):
    parts = re.split(r"[/\-.]", str(text).strip())
    if len(parts) == 3:
        try:
            month, day, year = int(parts[0]), int(parts[1]), int(parts[2])
            if year < 100:
                year += 2000
            return datetime.date(year, month, day)
        except ValueError:
            return None
    return None


def format_timestamp(value):
    try:
        stamp = datetime.datetime.fromisoformat(str(value))
    except ValueError:
        return str(value)
    return stamp.strftime("%b %d, %Y %I:%M %p")


def item_amount(item):
    qty = item.get("qty")
    rate = item.get("rate")
    if qty is None or rate is None:
        return None
    return qty * rate


def items_table_rows(items, filler_rows=0):
    rows = []
    for item in items:
        qty = item.get("qty")
        rate = item.get("rate")
        amount = item_amount(item)
        rows.append(
            "<tr>"
            f"<td>{esc(item.get('item_no', ''))}</td>"
            f"<td>{esc(item.get('description', ''))}</td>"
            f"<td class='c'>{esc(fmt_qty(qty))}</td>"
            f"<td class='r'>{esc(fmt_money(rate)) if rate is not None else ''}</td>"
            f"<td>{esc(item.get('uom', ''))}</td>"
            f"<td class='r'>{esc(fmt_money(amount)) if amount is not None else ''}</td>"
            "</tr>"
        )
    for _ in range(filler_rows):
        rows.append("<tr><td>&nbsp;</td><td></td><td></td><td></td><td></td><td></td></tr>")
    return "\n".join(rows)


def render_document(data):
    with open(TEMPLATE_PATH, encoding="utf-8") as handle:
        template = handle.read()
    with open(LOGO_PATH, "rb") as handle:
        logo_data = "data:image/png;base64," + base64.b64encode(handle.read()).decode("ascii")

    items = list(data.get("items", []))
    single_mode = len(items) <= SINGLE_PAGE_MAX_ITEMS
    filler_rows = max(0, SINGLE_TOTAL_ROWS - len(items)) if single_mode else 0
    total = sum(
        amount for amount in (item_amount(item) for item in items) if amount is not None
    )
    tokens = {
        "logo_data": logo_data,
        "page_mode": "single" if single_mode else "multi",
        "page_margin": "margin: 0;" if single_mode else "margin: 0.34in 0.42in 0.28in;",
        "po_no": esc(data.get("po_no")),
        "name": esc(data.get("name")),
        "order_date": esc(data.get("order_date")),
        "delivery_date": esc(data.get("delivery_date")),
        "items_rows": items_table_rows(items, filler_rows),
        "total": fmt_money(total),
        "approved_by": esc(data.get("approved_by")) or "&nbsp;",
        "approved_date": esc(data.get("approved_date")),
        "verified_by": esc(data.get("verified_by")) or "&nbsp;",
        "verified_date": esc(data.get("verified_date")),
        "received_by": esc(data.get("received_by")) or "&nbsp;",
        "received_date": esc(data.get("received_date")),
    }
    for key, value in tokens.items():
        template = template.replace("{{" + key + "}}", value)
    return template


def _run_browser(browser_path, args, timeout=60):
    try:
        process = subprocess.Popen(
            [browser_path, *args],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            creationflags=CREATE_NO_WINDOW,
        )
    except OSError as error:
        return -1, str(error)
    try:
        stdout, _stderr = process.communicate(timeout=timeout)
        output = stdout.decode("utf-8", errors="replace").strip()
        return process.returncode, output
    except subprocess.TimeoutExpired:
        try:
            subprocess.run(
                ["taskkill", "/T", "/F", "/PID", str(process.pid)],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=CREATE_NO_WINDOW,
                timeout=30,
            )
        except (OSError, subprocess.SubprocessError):
            process.kill()
        output = ""
        try:
            stdout, _stderr = process.communicate(timeout=15)
            output = stdout.decode("utf-8", errors="replace").strip()
        except (subprocess.TimeoutExpired, OSError):
            pass
        return -9, output + "\n[timed out and was terminated]"


BROWSER_FLAG_VARIANTS = (
    ["--headless=new", "--disable-gpu", "--no-sandbox"],
    ["--headless", "--disable-gpu", "--no-sandbox", "--use-angle=swiftshader"],
    ["--headless=new", "--no-sandbox", "--use-angle=swiftshader"],
)

BROWSER_COMMON_FLAGS = (
    "--disable-extensions",
    "--disable-component-update",
    "--disable-sync",
    "--disable-background-networking",
    "--no-first-run",
    "--no-default-browser-check",
)


def _browser_base_args(browser_name, flags):
    return [
        *flags,
        *BROWSER_COMMON_FLAGS,
        f"--user-data-dir={browser_profile_dir(browser_name)}",
    ]


def browser_print_pdf(browser_name, browser_path, html_path, pdf_path):
    url = pathlib.Path(html_path).as_uri()
    for flags in BROWSER_FLAG_VARIANTS:
        if not _prepare_target(pdf_path):
            log_render(f"{browser_name}: PDF target is open/locked, cannot replace {os.path.basename(pdf_path)}")
            return False
        code, output = _run_browser(
            browser_path,
            [
                *_browser_base_args(browser_name, flags),
                "--virtual-time-budget=10000",
                "--no-pdf-header-footer",
                f"--print-to-pdf={pdf_path}",
                url,
            ],
        )
        if valid_pdf(pdf_path):
            log_render(f"{browser_name}: PDF created (exit {code}) with {' '.join(flags)}")
            return True
        tail = output[-600:] if output else "-"
        log_render(f"{browser_name}: PDF failed (exit {code}) flags={' '.join(flags)} | {tail}")
    return False


def browser_screenshot(browser_name, browser_path, html_path, png_path):
    url = pathlib.Path(html_path).as_uri()
    for flags in BROWSER_FLAG_VARIANTS:
        if not _prepare_target(png_path):
            log_render(f"{browser_name}: PNG target is open/locked, cannot replace {os.path.basename(png_path)}")
            return False
        code, output = _run_browser(
            browser_path,
            [
                *_browser_base_args(browser_name, flags),
                "--hide-scrollbars",
                "--default-background-color=FFFFFFFF",
                "--run-all-compositor-stages-before-draw",
                "--window-size=816,1056",
                "--force-device-scale-factor=1.5625",
                "--virtual-time-budget=10000",
                f"--screenshot={png_path}",
                url,
            ],
        )
        if valid_png(png_path):
            log_render(f"{browser_name}: PNG created (exit {code}) with {' '.join(flags)}")
            return True
        tail = output[-600:] if output else "-"
        log_render(f"{browser_name}: PNG failed (exit {code}) flags={' '.join(flags)} | {tail}")
    return False


def _native_font(size, bold=False, script=False):
    from PIL import ImageFont

    if script:
        names = ("brushsci.ttf", "segoesc.ttf", "segoeuii.ttf", "ariali.ttf")
    elif bold:
        names = ("arialbd.ttf", "segoeuib.ttf")
    else:
        names = ("arial.ttf", "segoeui.ttf")
    for name in names:
        path = os.path.join(r"C:\Windows\Fonts", name)
        if os.path.isfile(path):
            try:
                return ImageFont.truetype(path, size)
            except OSError:
                continue
    return ImageFont.load_default()


def render_native_document(data, pdf_path, png_path):
    from PIL import Image, ImageDraw
    import PIL.PdfImagePlugin  # noqa: F401
    import PIL.PngImagePlugin  # noqa: F401

    dpi = 200
    scale = dpi / 96.0
    page_w, page_h = int(round(816 * scale)), int(round(1056 * scale))

    red = "#9b1c1c"
    pad_left = 40.32
    pad_right = 40.32
    pad_top = 32.64
    pad_bottom = 26.88
    x0 = pad_left
    x1 = 816 - pad_right
    content_width = x1 - x0
    content_bottom = 1056 - pad_bottom

    table_header_height = 26
    row_height = 24
    columns = [0.13, 0.38, 0.08, 0.13, 0.09, 0.19]
    headers = ["ITEM NO.", "DESCRIPTION", "QTY.", "RATE", "U/M", "AMOUNT"]
    column_x = [x0]
    for fraction in columns:
        column_x.append(column_x[-1] + content_width * fraction)

    info_font = _native_font(int(round(9.5 * scale)))
    title_font = _native_font(int(round(27 * scale)), bold=True)
    po_font = _native_font(int(round(16 * scale)), bold=True)
    name_font = _native_font(int(round(12 * scale)), bold=True)
    value_font = _native_font(int(round(12 * scale)))
    date_font = _native_font(int(round(11.5 * scale)), bold=True)
    header_font = _native_font(int(round(11 * scale)), bold=True)
    body_font = _native_font(int(round(11 * scale)))
    notes_font = _native_font(int(round(8.6 * scale)), bold=True)
    total_font = _native_font(int(round(18 * scale)), bold=True)
    sig_label_font = _native_font(int(round(11.5 * scale)))
    sig_name_font = _native_font(int(round(12 * scale)), bold=True)
    sig_caption_font = _native_font(int(round(10 * scale)))

    def u(value):
        return value * scale

    def text(draw, x, y, content, font, fill="black"):
        draw.text((u(x), u(y)), content, font=font, fill=fill)

    def text_right(draw, x_right, y, content, font, fill="black"):
        w = draw.textlength(content, font=font)
        draw.text((u(x_right) - w, u(y)), content, font=font, fill=fill)

    def text_center(draw, x_center, y, content, font, fill="black"):
        w = draw.textlength(content, font=font)
        draw.text((u(x_center) - w / 2, u(y)), content, font=font, fill=fill)

    def fit_text(draw, value, font, max_width):
        content = str(value)
        if not content:
            return content
        if draw.textlength(content, font=font) <= max_width:
            return content
        while content and draw.textlength(content + "...", font=font) > max_width:
            content = content[:-1]
        return content + "..."

    def make_page():
        image = Image.new("RGB", (page_w, page_h), "white")
        return image, ImageDraw.Draw(image)

    def draw_company_header(draw, image):
        logo = Image.open(LOGO_PATH).convert("RGBA")
        logo_width = int(round(u(250)))
        logo_height = int(round(logo.height * logo_width / logo.width))
        logo = logo.resize((logo_width, logo_height), Image.LANCZOS)
        image.paste(logo, (int(u(x0)), int(u(pad_top))), logo)

        info_lines = [
            "Email Add.: takaraexport@gmail.com",
            "Telefax No.: +632 86609256",
            "#18 Gómez St. Kapalaran Subd. Brgy. San Juan, Taytay Rizal",
        ]
        info_top = pad_top + logo_height / scale + 10
        for index, line in enumerate(info_lines):
            text(draw, x0, info_top + index * 14.5, line, info_font)
        info_bottom = info_top + 2 * 14.5 + 10

        text_right(draw, x1, pad_top, "PURCHASE", title_font)
        text_right(draw, x1, pad_top + 29, "ORDER", title_font)

        po_label = "P.O. No.: "
        po_value = str(data.get("po_no", ""))
        box_top = pad_top + 67
        box_pad_x, box_pad_y = 12, 5
        label_width = draw.textlength(po_label, font=po_font)
        value_width = draw.textlength(po_value, font=po_font)
        box_width = label_width + value_width + u(box_pad_x * 2)
        box_height = u(16 + box_pad_y * 2)
        box_x0 = u(x1) - box_width
        box_y0 = u(box_top)
        draw.rectangle(
            [box_x0, box_y0, box_x0 + box_width, box_y0 + box_height],
            outline=red,
            width=max(2, int(u(2))),
        )
        text(
            draw,
            x1 - box_pad_x - (label_width + value_width) / scale,
            box_top + box_pad_y,
            po_label,
            po_font,
        )
        text(
            draw,
            x1 - box_pad_x - value_width / scale,
            box_top + box_pad_y,
            po_value,
            po_font,
            fill=red,
        )

        return max(info_bottom, box_top + box_height / scale)

    def draw_name_row(draw, header_bottom):
        name_top = header_bottom + 24
        name_box_width = content_width * 0.72
        name_box_height = 58
        draw.rectangle(
            [u(x0), u(name_top), u(x0 + name_box_width), u(name_top + name_box_height)],
            outline="black",
            width=max(2, int(u(1.5))),
        )
        text(draw, x0 + 10, name_top + 8, "NAME:", name_font)
        name_label_width = draw.textlength("NAME:", font=name_font) / scale
        text(
            draw,
            x0 + 10 + name_label_width + 6,
            name_top + 8,
            str(data.get("name", "")).upper(),
            value_font,
        )

        date_x0 = x0 + name_box_width + 20
        text(draw, date_x0, name_top + 9, "Order Date:", date_font)
        text_right(draw, x1, name_top + 9, str(data.get("order_date", "")), date_font, fill=red)
        text(draw, date_x0, name_top + 31, "Delivery Date:", date_font)
        text_right(draw, x1, name_top + 31, str(data.get("delivery_date", "")), date_font, fill=red)
        return name_top + name_box_height + 16

    def draw_table(draw, top, row_items):
        bottom = top + table_header_height + len(row_items) * row_height
        draw.rectangle([u(x0), u(top), u(x1), u(bottom)], outline="black", width=max(2, int(u(1.5))))
        draw.rectangle(
            [u(x0), u(top), u(x1), u(top + table_header_height)],
            fill="#ececec",
            outline="black",
            width=max(1, int(u(1))),
        )
        for index, header in enumerate(headers):
            if index == 2:
                text_center(draw, (column_x[index] + column_x[index + 1]) / 2, top + 8, header, header_font)
            elif index in (3, 5):
                text_right(draw, column_x[index + 1] - 7, top + 8, header, header_font)
            else:
                text(draw, column_x[index] + 7, top + 8, header, header_font)

        for column in column_x[1:-1]:
            draw.line([u(column), u(top), u(column), u(bottom)], fill="black", width=max(1, int(u(1))))
        for row_index in range(len(row_items) + 1):
            row_y = top + table_header_height + row_index * row_height
            draw.line([u(x0), u(row_y), u(x1), u(row_y)], fill="black", width=max(1, int(u(1))))

        for row_index, item in enumerate(row_items):
            qty = item.get("qty")
            rate = item.get("rate")
            amount = item_amount(item) if item else None
            values = [
                item.get("item_no", ""),
                item.get("description", ""),
                fmt_qty(qty) if qty is not None else "",
                fmt_money(rate) if rate is not None else "",
                item.get("uom", ""),
                fmt_money(amount) if amount is not None else "",
            ]
            row_top = top + table_header_height + row_index * row_height
            baseline = row_top + (row_height - 12) / 2 + 1
            for index, value in enumerate(values):
                if not value:
                    continue
                if index == 2:
                    text_center(draw, (column_x[index] + column_x[index + 1]) / 2, baseline, str(value), body_font)
                elif index in (3, 5):
                    text_right(draw, column_x[index + 1] - 7, baseline, str(value), body_font)
                else:
                    max_width = u(content_width * columns[index]) - u(14)
                    text(
                        draw,
                        column_x[index] + 7,
                        baseline,
                        fit_text(draw, value, body_font, max_width),
                        body_font,
                    )
        return bottom

    notes_lines = [
        "NOTE:",
        "- ALL ACCESSORIES & MATERIALS ARE CHARGEABLE.",
        "- PLEASE MAKE COUNTER SAMPLE BEFORE PRODUCTION.",
        "- PARTIAL DELIVERY ALLOWED.",
        "- PLEASE FOLLOW DUE DATE TO AVOID PENALTY CHARGES & CANCELLATION.",
        '- PLEASE MAKE DELIVERY RECEIPT DURING DELIVERY "NO D.R. NO PAYMENT".',
        "- 4% PENALTY FOR EVERY LATE DELIVERIES.",
    ]

    items = list(data.get("items", []))
    total = sum(amount for amount in (item_amount(item) for item in items) if amount is not None)

    def draw_notes_total(draw, top):
        for index, line in enumerate(notes_lines):
            text(draw, x0, top + index * 13.4, line, notes_font)

        total_value = f"PHP {fmt_money(total)}"
        total_label_width = draw.textlength("Total Amount", font=total_font)
        total_value_width = draw.textlength(total_value, font=total_font)
        total_box_width = max(u(320), total_label_width + total_value_width + u(60))
        total_box_height = u(46)
        total_box_x1 = u(x1)
        total_box_y0 = u(top + 6)
        draw.rectangle(
            [total_box_x1 - total_box_width, total_box_y0, total_box_x1, total_box_y0 + total_box_height],
            outline="black",
            width=max(2, int(u(2))),
        )
        text(
            draw,
            (total_box_x1 - total_box_width) / scale + 20,
            (total_box_y0 + u(14)) / scale,
            "Total Amount",
            total_font,
        )
        text_right(draw, x1 - 20, (total_box_y0 + u(14)) / scale, total_value, total_font)
        return top + len(notes_lines) * 13.4

    signatures = [
        ("Approved by:", str(data.get("approved_by", "")), str(data.get("approved_date", ""))),
        ("Verified by:", str(data.get("verified_by", "")), str(data.get("verified_date", ""))),
        ("Received by:", str(data.get("received_by", "")), str(data.get("received_date", ""))),
    ]

    def draw_signatures(draw, label_y):
        for index, (label, signer, date_value) in enumerate(signatures):
            column_left = x0 + index * (content_width / 3)
            text(draw, column_left, label_y, label, sig_label_font)
            if signer:
                text(draw, column_left, label_y + 24, signer, sig_name_font)
            text(draw, column_left, label_y + 66, "Signature over Printed Name", sig_caption_font)
            text(draw, column_left, label_y + 82, f"Date: {date_value}".rstrip(), sig_caption_font)

    signature_block_height = 95
    notes_block_height = len(notes_lines) * 13.4
    single_mode = len(items) <= SINGLE_PAGE_MAX_ITEMS

    if single_mode:
        image, draw = make_page()
        header_bottom = draw_company_header(draw, image)
        table_top = draw_name_row(draw, header_bottom)
        rows = items + [{}] * (SINGLE_TOTAL_ROWS - len(items))
        table_bottom = draw_table(draw, table_top, rows)
        draw_notes_total(draw, table_bottom + 16)
        draw_signatures(draw, content_bottom - signature_block_height)
        pages = [image]
    else:
        image, draw = make_page()
        header_bottom = draw_company_header(draw, image)
        table_top = draw_name_row(draw, header_bottom)
        first_capacity = int((content_bottom - table_top - table_header_height) // row_height)
        next_capacity = int((content_bottom - pad_top - table_header_height) // row_height)
        page_rows = [items[:first_capacity]]
        remaining = items[first_capacity:]
        while remaining:
            page_rows.append(remaining[:next_capacity])
            remaining = remaining[next_capacity:]

        last_rows = page_rows[-1]
        last_top = table_top if len(page_rows) == 1 else pad_top
        last_bottom = last_top + table_header_height + len(last_rows) * row_height
        block_height = 16 + notes_block_height + 10 + signature_block_height
        tail_on_new_page = last_bottom + block_height > content_bottom

        pages = [image]
        for page_index, rows in enumerate(page_rows):
            if page_index > 0:
                image, draw = make_page()
                pages.append(image)
            top = table_top if page_index == 0 else pad_top
            bottom = draw_table(draw, top, rows)
            if page_index == len(page_rows) - 1:
                if tail_on_new_page:
                    image, draw = make_page()
                    pages.append(image)
                    notes_top = pad_top
                else:
                    notes_top = bottom + 16
                notes_bottom = draw_notes_total(draw, notes_top)
                draw_signatures(draw, notes_bottom + 10)

    pages[0].save(png_path, "PNG")
    pages[0].save(pdf_path, "PDF", resolution=dpi, save_all=True, append_images=pages[1:])
    page_word = "page" if len(pages) == 1 else "pages"
    log_render(
        f"built-in renderer: created {os.path.basename(pdf_path)} and {os.path.basename(png_path)} "
        f"({len(pages)} {page_word})"
    )

def generate_documents(data, out_dir=None, base_name=None):
    if out_dir is None:
        out_dir = DOCUMENTS_DIR
    os.makedirs(out_dir, exist_ok=True)
    base = base_name or f"PO_{safe_filename(data.get('po_no'))}_{safe_filename(data.get('name'))}"
    pdf_path = os.path.join(out_dir, base + ".pdf")
    png_path = os.path.join(out_dir, base + ".png")

    browsers = find_browsers()
    if browsers:
        render_dir = tempfile.mkdtemp(prefix="po_render_")
        try:
            html_path = os.path.join(render_dir, "render.html")
            with open(html_path, "w", encoding="utf-8") as handle:
                handle.write(render_document(data))
            for browser_name, browser_path in browsers:
                pdf_ok = browser_print_pdf(browser_name, browser_path, html_path, pdf_path)
                if pdf_ok and browser_screenshot(browser_name, browser_path, html_path, png_path):
                    return pdf_path, png_path, browser_name.lower()
        finally:
            for _attempt in range(3):
                shutil.rmtree(render_dir, ignore_errors=True)
                if not os.path.exists(render_dir):
                    break
                time.sleep(0.3)
    else:
        log_render("no supported browser found; using built-in renderer")

    render_native_document(data, pdf_path, png_path)
    return pdf_path, png_path, "builtin"


class DatePickerPopup(tk.Toplevel):
    ACCENT = "#2563eb"

    def __init__(self, master, variable, anchor):
        super().__init__(master)
        self.variable = variable
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.configure(bg="#ffffff", highlightthickness=1, highlightbackground="#8a8f98")

        today = datetime.date.today()
        current = parse_date_text(variable.get()) or today
        self.view_year = current.year
        self.view_month = current.month
        self.selected = current
        self.today = today

        self._build()
        self.update_idletasks()
        width, height = self.winfo_reqwidth(), self.winfo_reqheight()
        x = anchor.winfo_rootx()
        y = anchor.winfo_rooty() + anchor.winfo_height() + 2
        x = max(4, min(x, self.winfo_screenwidth() - width - 4))
        y = max(4, min(y, self.winfo_screenheight() - height - 4))
        self.geometry(f"+{x}+{y}")

        self.focus_force()
        self.bind("<Escape>", lambda _event: self.destroy())
        self.bind("<FocusOut>", self._on_focus_out)

    def _on_focus_out(self, _event):
        self.after(80, self._close_if_unfocused)

    def _close_if_unfocused(self):
        try:
            focused = self.focus_displayof()
            if focused is None:
                self.destroy()
                return
            widget = focused
            while widget is not None:
                if widget is self:
                    return
                widget = getattr(widget, "master", None)
            self.destroy()
        except tk.TclError:
            pass

    def _build(self):
        header = tk.Frame(self, bg=self.ACCENT)
        header.pack(fill="x")
        tk.Button(
            header,
            text="<",
            command=lambda: self._shift_month(-1),
            relief="flat",
            bd=0,
            bg=self.ACCENT,
            fg="#ffffff",
            activebackground="#1d4ed8",
            activeforeground="#ffffff",
            font=("Segoe UI", 10, "bold"),
            width=3,
        ).pack(side="left")
        self.month_label = tk.Label(
            header,
            text="",
            bg=self.ACCENT,
            fg="#ffffff",
            font=("Segoe UI", 10, "bold"),
            width=16,
        )
        self.month_label.pack(side="left")
        tk.Button(
            header,
            text=">",
            command=lambda: self._shift_month(1),
            relief="flat",
            bd=0,
            bg=self.ACCENT,
            fg="#ffffff",
            activebackground="#1d4ed8",
            activeforeground="#ffffff",
            font=("Segoe UI", 10, "bold"),
            width=3,
        ).pack(side="left")

        self.grid_frame = tk.Frame(self, bg="#ffffff", padx=6, pady=4)
        self.grid_frame.pack()

        footer = tk.Frame(self, bg="#ffffff", pady=4)
        footer.pack(fill="x")
        tk.Button(
            footer,
            text="Today",
            command=lambda: self._pick(self.today),
            relief="flat",
            bd=0,
            bg="#eef2ff",
            fg=self.ACCENT,
            activebackground="#dbeafe",
            font=("Segoe UI", 9),
        ).pack()

        self._render_month()

    def _shift_month(self, delta):
        month = self.view_month + delta
        year = self.view_year
        if month < 1:
            month, year = 12, year - 1
        elif month > 12:
            month, year = 1, year + 1
        self.view_month, self.view_year = month, year
        self._render_month()

    def _render_month(self):
        self.month_label.configure(text=f"{calendar.month_name[self.view_month]} {self.view_year}")
        for child in self.grid_frame.winfo_children():
            child.destroy()

        for column, day_name in enumerate(("Su", "Mo", "Tu", "We", "Th", "Fr", "Sa")):
            tk.Label(
                self.grid_frame,
                text=day_name,
                fg="#6b7280",
                bg="#ffffff",
                font=("Segoe UI", 8, "bold"),
                width=3,
            ).grid(row=0, column=column, padx=1)

        weeks = calendar.Calendar(firstweekday=6).monthdatescalendar(self.view_year, self.view_month)
        for row, week in enumerate(weeks, start=1):
            for column, day in enumerate(week):
                in_month = day.month == self.view_month
                if day == self.selected:
                    background, foreground = self.ACCENT, "#ffffff"
                elif day == self.today:
                    background, foreground = "#dbeafe", self.ACCENT
                else:
                    background = "#ffffff"
                    foreground = "#111827" if in_month else "#c3c7cd"
                tk.Button(
                    self.grid_frame,
                    text=str(day.day),
                    command=lambda d=day: self._pick(d),
                    relief="flat",
                    bd=0,
                    bg=background,
                    fg=foreground,
                    activebackground="#dbeafe",
                    activeforeground="#1e3a8a",
                    font=("Segoe UI", 9),
                    width=3,
                ).grid(row=row, column=column, padx=1, pady=1)

    def _pick(self, day):
        self.variable.set(f"{day.month}/{day.day}/{day.year}")
        self.destroy()


class PurchaseOrderApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Takara Purchase Order Generator")
        self.scale = self._apply_dpi_scaling()
        self.geometry(f"{int(1060 * self.scale)}x{int(880 * self.scale)}")
        self.minsize(int(980 * self.scale), int(780 * self.scale))
        self.configure(bg="#f0f1f3")
        self._set_icon()
        self._init_style()

        self.po_var = tk.StringVar()
        self.name_var = tk.StringVar()
        self.order_date_var = tk.StringVar(value=datetime.date.today().strftime("%m/%d/%Y"))
        self.delivery_date_var = tk.StringVar()
        self.approved_by_var = tk.StringVar(value=DEFAULT_APPROVER)
        self.approved_date_var = tk.StringVar()
        self.verified_by_var = tk.StringVar()
        self.verified_date_var = tk.StringVar()
        self.received_by_var = tk.StringVar()
        self.received_date_var = tk.StringVar()
        self.total_var = tk.StringVar(value="PHP 0.00")
        self.status_var = tk.StringVar(value="Ready.")

        self.item_rows = []
        self.last_pdf = None
        self.last_png = None
        self.dirty = False
        self.current_order_file = None
        self.current_order_meta = {}
        self.history_window = None

        for variable in (
            self.po_var,
            self.name_var,
            self.order_date_var,
            self.delivery_date_var,
            self.approved_by_var,
            self.approved_date_var,
            self.verified_by_var,
            self.verified_date_var,
            self.received_by_var,
            self.received_date_var,
        ):
            variable.trace_add("write", self._mark_dirty)

        self._build_ui()
        for _ in range(3):
            self.add_item_row()
        self._set_dirty(False)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _mark_dirty(self, *_args):
        if not self.dirty:
            self.dirty = True
            self._update_title()

    def _set_dirty(self, flag):
        self.dirty = bool(flag)
        self._update_title()

    def _update_title(self):
        suffix = " *" if self.dirty else ""
        self.title(f"Takara Purchase Order Generator{suffix}")

    def _apply_dpi_scaling(self):
        try:
            from ctypes import windll

            dpi = windll.user32.GetDpiForSystem()
            self.tk.call("tk", "scaling", dpi / 72.0)
            return dpi / 96.0
        except Exception:
            return 1.0

    def _set_icon(self):
        try:
            self.iconphoto(True, tk.PhotoImage(file=ICON_PATH))
        except tk.TclError:
            pass

    def _init_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Card.TLabelframe", background="#f0f1f3")
        style.configure("Card.TLabelframe.Label", background="#f0f1f3", font=("Segoe UI", 10, "bold"))
        style.configure("TLabel", background="#f0f1f3", font=("Segoe UI", 10))
        style.configure("Hint.TLabel", foreground="#666666", font=("Segoe UI", 9))
        style.configure("Total.TLabel", font=("Segoe UI", 12, "bold"))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))

    def _build_ui(self):
        root = ttk.Frame(self, padding=14)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)

        details = ttk.Labelframe(root, text="Order Details", style="Card.TLabelframe", padding=12)
        details.grid(row=0, column=0, sticky="ew")
        details.columnconfigure(1, weight=1)
        details.columnconfigure(3, weight=2)

        ttk.Label(details, text="P.O. No.:").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=4)
        ttk.Entry(details, textvariable=self.po_var, width=18).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Label(details, text="Name:").grid(row=0, column=2, sticky="w", padx=(18, 6), pady=4)
        ttk.Entry(details, textvariable=self.name_var).grid(row=0, column=3, sticky="ew", pady=4)

        ttk.Label(details, text="Order Date:").grid(row=1, column=0, sticky="w", padx=(0, 6), pady=4)
        self._date_field(details, self.order_date_var).grid(row=1, column=1, sticky="w", pady=4)
        ttk.Label(details, text="Delivery Date:").grid(row=1, column=2, sticky="w", padx=(18, 6), pady=4)
        self._date_field(details, self.delivery_date_var).grid(row=1, column=3, sticky="w", pady=4)

        items = ttk.Labelframe(root, text="Items", style="Card.TLabelframe", padding=12)
        items.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        items.columnconfigure(0, weight=1)
        items.rowconfigure(1, weight=1)

        toolbar = ttk.Frame(items)
        toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Button(toolbar, text="+ Add Item", command=self.add_item_row).pack(side="left")
        ttk.Label(
            toolbar,
            text="Amount and total are computed automatically from Qty x Rate.",
            style="Hint.TLabel",
        ).pack(side="left", padx=(12, 0))

        canvas = tk.Canvas(items, background="#ffffff", highlightthickness=1, highlightbackground="#c9ccd1")
        scrollbar = ttk.Scrollbar(items, orient="vertical", command=canvas.yview)
        self.items_frame = tk.Frame(canvas, background="#ffffff")
        self.items_window = canvas.create_window((0, 0), window=self.items_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=1, column=0, sticky="nsew")
        scrollbar.grid(row=1, column=1, sticky="ns")

        def on_frame_configure(_event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def on_canvas_configure(event):
            canvas.itemconfigure(self.items_window, width=event.width)

        self.items_frame.bind("<Configure>", on_frame_configure)
        canvas.bind("<Configure>", on_canvas_configure)
        canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-event.delta / 120), "units"))

        signatories = ttk.Labelframe(root, text="Signatories", style="Card.TLabelframe", padding=12)
        signatories.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        for column in (1, 3, 5):
            signatories.columnconfigure(column, weight=1)

        ttk.Label(signatories, text="Approved by:").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=4)
        ttk.Entry(signatories, textvariable=self.approved_by_var).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Label(signatories, text="Date:").grid(row=0, column=2, sticky="w", padx=(18, 6), pady=4)
        self._date_field(signatories, self.approved_date_var, width=12).grid(row=0, column=3, sticky="w", pady=4)
        ttk.Label(signatories, text="Received by:").grid(row=0, column=4, sticky="w", padx=(18, 6), pady=4)

        ttk.Label(signatories, text="Verified by:").grid(row=1, column=0, sticky="w", padx=(0, 6), pady=4)
        ttk.Entry(signatories, textvariable=self.verified_by_var).grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Label(signatories, text="Date:").grid(row=1, column=2, sticky="w", padx=(18, 6), pady=4)
        self._date_field(signatories, self.verified_date_var, width=12).grid(row=1, column=3, sticky="w", pady=4)
        ttk.Entry(signatories, textvariable=self.received_by_var).grid(row=0, column=5, sticky="ew", pady=4)
        ttk.Label(signatories, text="Date:").grid(row=1, column=4, sticky="w", padx=(18, 6), pady=4)
        self._date_field(signatories, self.received_date_var, width=12).grid(row=1, column=5, sticky="w", pady=4)

        footer = ttk.Frame(root)
        footer.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        ttk.Label(footer, text="Total Amount:").pack(side="left")
        ttk.Label(footer, textvariable=self.total_var, style="Total.TLabel").pack(side="left", padx=(8, 0))

        ttk.Button(footer, text="Clear Form", command=self.clear_form).pack(side="right")
        ttk.Button(footer, text="Print", command=self.print_document).pack(side="right", padx=(0, 8))
        ttk.Button(footer, text="Preview", command=self.preview_document).pack(side="right", padx=(0, 8))
        ttk.Button(
            footer, text="Generate PDF && PNG", style="Accent.TButton", command=self.generate_documents_clicked
        ).pack(side="right", padx=(0, 8))
        ttk.Button(footer, text="Save", command=self.save_order_clicked).pack(side="right", padx=(0, 8))
        ttk.Button(footer, text="Saved POs", command=self.open_history).pack(side="right", padx=(0, 8))
        ttk.Button(footer, text="Library Folder", command=self.open_library_folder).pack(side="right", padx=(0, 8))

        ttk.Label(root, textvariable=self.status_var, style="Hint.TLabel").grid(row=4, column=0, sticky="w", pady=(8, 0))

    def _date_field(self, parent, variable, width=18):
        frame = ttk.Frame(parent)
        entry = ttk.Entry(frame, textvariable=variable, width=width)
        entry.pack(side="left", fill="x", expand=True)
        ttk.Button(
            frame,
            text="\U0001F4C5",
            width=3,
            command=lambda: DatePickerPopup(self, variable, entry),
        ).pack(side="left", padx=(2, 0))
        return frame

    def add_item_row(self):
        row = {
            "item_no": tk.StringVar(),
            "description": tk.StringVar(),
            "qty": tk.StringVar(),
            "rate": tk.StringVar(),
            "uom": tk.StringVar(),
            "amount": tk.StringVar(),
        }
        row["qty"].trace_add("write", lambda *_args, r=row: self._update_row_amount(r))
        row["rate"].trace_add("write", lambda *_args, r=row: self._update_row_amount(r))
        for key in ("item_no", "description", "qty", "rate", "uom"):
            row[key].trace_add("write", self._mark_dirty)
        self.item_rows.append(row)
        self._rebuild_items_ui()

    def remove_item_row(self, row):
        if row in self.item_rows:
            self.item_rows.remove(row)
        if not self.item_rows:
            self.add_item_row()
            return
        self._rebuild_items_ui()

    def _rebuild_items_ui(self):
        for child in self.items_frame.winfo_children():
            child.destroy()

        headers = ["ITEM NO.", "DESCRIPTION", "QTY.", "RATE", "U/M", "AMOUNT", ""]
        for column, text in enumerate(headers):
            tk.Label(
                self.items_frame,
                text=text,
                font=("Segoe UI", 9, "bold"),
                background="#e9ebee",
                anchor="w",
                padx=6,
                pady=4,
            ).grid(row=0, column=column, sticky="ew", padx=1, pady=(0, 2))

        widths = [14, 40, 8, 12, 8, 14]
        for index, row in enumerate(self.item_rows, start=1):
            for column, (key, width) in enumerate(zip(["item_no", "description", "qty", "rate", "uom"], widths)):
                entry = tk.Entry(self.items_frame, textvariable=row[key], width=width)
                entry.grid(row=index, column=column, sticky="ew", padx=1, pady=2)
            tk.Label(
                self.items_frame,
                textvariable=row["amount"],
                background="#ffffff",
                anchor="e",
                padx=6,
                font=("Segoe UI", 9),
            ).grid(row=index, column=5, sticky="ew", padx=1, pady=2)
            tk.Button(
                self.items_frame,
                text="✕",
                width=2,
                command=lambda r=row: self.remove_item_row(r),
            ).grid(row=index, column=6, padx=(6, 6), pady=2)
        self.items_frame.columnconfigure(1, weight=1)
        self._update_row_amount(None)

    def _update_row_amount(self, _row):
        total = 0.0
        has_amount = False
        for row in self.item_rows:
            qty = parse_number(row["qty"].get())
            rate = parse_number(row["rate"].get())
            if qty is not None and rate is not None:
                amount = qty * rate
                row["amount"].set(fmt_money(amount))
                total += amount
                has_amount = True
            else:
                row["amount"].set("")
        self.total_var.set(f"PHP {fmt_money(total)}" if has_amount or self.item_rows else "PHP 0.00")

    def collect_data(self):
        po_no = self.po_var.get().strip()
        name = self.name_var.get().strip()
        if not po_no:
            messagebox.showwarning("Missing P.O. No.", "Please enter a P.O. No.")
            return None
        if not name:
            messagebox.showwarning("Missing Name", "Please enter a Name.")
            return None

        items = []
        for index, row in enumerate(self.item_rows, start=1):
            values = {
                "item_no": row["item_no"].get().strip(),
                "description": row["description"].get().strip(),
                "qty_text": row["qty"].get().strip(),
                "rate_text": row["rate"].get().strip(),
                "uom": row["uom"].get().strip(),
            }
            if not any(values.values()):
                continue
            qty = parse_number(values["qty_text"])
            rate = parse_number(values["rate_text"])
            if values["qty_text"] and qty is None:
                messagebox.showwarning("Invalid QTY", f"Row {index}: QTY must be a number.")
                return None
            if values["rate_text"] and rate is None:
                messagebox.showwarning("Invalid RATE", f"Row {index}: RATE must be a number.")
                return None
            items.append(
                {
                    "item_no": values["item_no"],
                    "description": values["description"],
                    "qty": qty,
                    "rate": rate,
                    "uom": values["uom"],
                }
            )

        return {
            "po_no": po_no,
            "name": name,
            "order_date": self.order_date_var.get().strip(),
            "delivery_date": self.delivery_date_var.get().strip(),
            "items": items,
            "approved_by": self.approved_by_var.get().strip(),
            "approved_date": self.approved_date_var.get().strip(),
            "verified_by": self.verified_by_var.get().strip(),
            "verified_date": self.verified_date_var.get().strip(),
            "received_by": self.received_by_var.get().strip(),
            "received_date": self.received_date_var.get().strip(),
        }

    def save_order_clicked(self):
        result = self.save_current()
        if isinstance(result, tuple):
            json_path, pdf_path, png_path = result
            detail = f"Record:  {os.path.basename(json_path)}"
            if pdf_path:
                detail += f"\nDocument:  {os.path.basename(pdf_path)}\nImage:  {os.path.basename(png_path)}"
            messagebox.showinfo("Saved", f"P.O. saved to the library:\n\n{detail}")

    def save_current(self, data=None, generate=True):
        if data is None:
            data = self.collect_data()
        if data is None:
            return None

        same_record = (
            self.current_order_file
            and self.current_order_meta.get("po_no") == data["po_no"]
            and self.current_order_meta.get("name") == data["name"]
        )
        candidate = order_file_name(data["po_no"], data["name"])
        if same_record:
            target = self.current_order_file
        elif os.path.exists(order_path(candidate)):
            existing = read_order(candidate) or {}
            choice = self._ask_duplicate_dialog(candidate, existing, data)
            if choice == "keep":
                target = next_copy_file_name(data["po_no"], data["name"])
            elif choice == "replace":
                target = candidate
            else:
                self.status_var.set("Save skipped - not stored in the P.O. library.")
                return "cancelled"
        else:
            target = candidate

        pdf_path = png_path = renderer = None
        if generate:
            self.status_var.set("Saving P.O. and generating document...")
            self.update_idletasks()
            try:
                pdf_path, png_path, renderer = generate_documents(
                    data, out_dir=DOCUMENTS_DIR, base_name=target[:-5]
                )
            except Exception as error:
                pdf_path = png_path = renderer = None
                messagebox.showwarning(
                    "Document not generated",
                    f"The P.O. was saved, but its PDF/PNG could not be generated:\n\n{error}",
                )

        order = normalize_order(data, existing=read_order(target))
        if pdf_path:
            order["document"] = {
                "pdf": os.path.basename(pdf_path),
                "png": os.path.basename(png_path),
                "renderer": renderer,
                "updated": datetime.datetime.now().isoformat(timespec="seconds"),
            }
        try:
            write_order(order, target)
        except OSError as error:
            messagebox.showerror("Save failed", f"Could not save the P.O.:\n\n{error}")
            return "failed"

        self.current_order_file = target
        self.current_order_meta = {"po_no": data["po_no"], "name": data["name"]}
        self._set_dirty(False)
        if pdf_path:
            note = " (built-in renderer)" if renderer == "builtin" else ""
            self.status_var.set(f"Saved: {target}  +  {os.path.basename(pdf_path)}{note}")
        else:
            self.status_var.set(f"Saved record: {target} (document missing)")
        self._refresh_history()
        return os.path.join(ORDERS_DIR, target), pdf_path, png_path

    def _ask_duplicate_dialog(self, file_name, existing, data):
        dialog = tk.Toplevel(self)
        dialog.title("Duplicate P.O.")
        dialog.transient(self)
        dialog.resizable(False, False)
        dialog.configure(bg="#f0f1f3")
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(
            frame, text="A saved P.O. already uses this file name:", font=("Segoe UI", 10, "bold")
        ).pack(anchor="w")
        ttk.Label(frame, text=file_name, style="Hint.TLabel").pack(anchor="w", pady=(2, 10))
        ttk.Label(
            frame,
            text=(
                f"SAVED:   {existing.get('po_no', '')} - {existing.get('name', '')}\n"
                f"                modified {format_timestamp(existing.get('modified', ''))}\n\n"
                f"NEW:       {data['po_no']} - {data['name']}"
            ),
            justify="left",
        ).pack(anchor="w")
        ttk.Label(
            frame,
            text="Keep both records, replace the saved record, or cancel this save?",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(10, 4))
        ttk.Label(
            frame,
            text=f'"Keep Both" saves the new copy as: {next_copy_file_name(data["po_no"], data["name"])}',
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(0, 10))

        choice = {"value": "cancel"}

        def pick(value):
            choice["value"] = value
            dialog.destroy()

        buttons = ttk.Frame(frame)
        buttons.pack(anchor="e")
        ttk.Button(buttons, text="Keep Both", style="Accent.TButton", command=lambda: pick("keep")).pack(side="left")
        ttk.Button(buttons, text="Replace", command=lambda: pick("replace")).pack(side="left", padx=8)
        ttk.Button(buttons, text="Cancel", command=lambda: pick("cancel")).pack(side="left")

        dialog.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - dialog.winfo_reqwidth()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - dialog.winfo_reqheight()) // 2
        dialog.geometry(f"+{max(0, x)}+{max(0, y)}")
        self.wait_window(dialog)
        return choice["value"]

    def open_history(self):
        if self.history_window and self.history_window.winfo_exists():
            self.history_window.lift()
            self.history_window.focus_force()
            self._refresh_history()
            return

        window = tk.Toplevel(self)
        self.history_window = window
        window.title("Saved Purchase Orders")
        window.transient(self)
        window.geometry(f"{int(900 * self.scale)}x{int(540 * self.scale)}")
        window.minsize(int(720 * self.scale), int(400 * self.scale))
        window.configure(bg="#f0f1f3")

        container = ttk.Frame(window, padding=12)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        search_bar = ttk.Frame(container)
        search_bar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ttk.Label(search_bar, text="Search:").pack(side="left")
        self.history_search_var = tk.StringVar()
        ttk.Entry(search_bar, textvariable=self.history_search_var).pack(
            side="left", fill="x", expand=True, padx=(6, 0)
        )
        self.history_search_var.trace_add("write", lambda *_args: self._refresh_history())

        columns = ("po", "name", "order", "delivery", "total", "document", "modified")
        self.history_tree = ttk.Treeview(
            container, columns=columns, show="headings", height=14, selectmode="browse"
        )
        headings = {
            "po": ("P.O. No.", 90, "w"),
            "name": ("Name", 220, "w"),
            "order": ("Order Date", 95, "center"),
            "delivery": ("Delivery Date", 95, "center"),
            "total": ("Total", 105, "e"),
            "document": ("Document", 95, "center"),
            "modified": ("Last Modified", 150, "center"),
        }
        for key, (text, width, anchor) in headings.items():
            self.history_tree.heading(key, text=text)
            self.history_tree.column(
                key, width=int(width * self.scale), anchor=anchor, stretch=(key == "name")
            )
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=scrollbar.set)
        self.history_tree.grid(row=1, column=0, sticky="nsew")
        scrollbar.grid(row=1, column=1, sticky="ns")
        self.history_tree.bind("<Double-1>", lambda _event: self.open_selected_order())
        self.history_tree.bind("<Return>", lambda _event: self.open_selected_order())

        self.history_status_var = tk.StringVar()
        actions = ttk.Frame(container)
        actions.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(actions, text="Open", style="Accent.TButton", command=self.open_selected_order).pack(side="left")
        ttk.Button(actions, text="Open Document", command=self.open_selected_document).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Duplicate", command=self.duplicate_selected_order).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Delete", command=self.delete_selected_order).pack(side="left", padx=(8, 0))
        ttk.Label(actions, textvariable=self.history_status_var, style="Hint.TLabel").pack(
            side="left", padx=(12, 0)
        )
        ttk.Button(actions, text="Close", command=window.destroy).pack(side="right")
        ttk.Button(actions, text="Open Folder", command=self.open_library_folder).pack(
            side="right", padx=(0, 8)
        )

        self._refresh_history()

    def _refresh_history(self):
        if not (self.history_window and self.history_window.winfo_exists()):
            return
        query = self.history_search_var.get().strip().lower()
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        count = 0
        for entry in list_orders():
            order = entry["order"]
            po_no = str(order.get("po_no", ""))
            name = str(order.get("name", ""))
            if query and query not in po_no.lower() and query not in name.lower():
                continue
            self.history_tree.insert(
                "",
                "end",
                iid=entry["file"],
                values=(
                    f"{po_no}{order_copy_suffix(entry['file'])}",
                    name,
                    order.get("order_date", ""),
                    order.get("delivery_date", ""),
                    f"PHP {fmt_money(entry['total'])}",
                    "PDF + PNG" if has_document(entry["file"]) else "missing",
                    format_timestamp(entry["modified"]),
                ),
            )
            count += 1
        self.history_status_var.set(f"{count} saved P.O.(s)")

    def _selected_history_file(self):
        selection = self.history_tree.selection()
        if not selection:
            messagebox.showinfo("Saved P.O.s", "Select a P.O. from the list first.")
            return None
        return selection[0]

    def open_selected_order(self):
        file_name = self._selected_history_file()
        if not file_name:
            return
        if self.dirty and not messagebox.askyesno(
            "Unsaved changes", "Discard unsaved changes and open this saved P.O.?"
        ):
            return
        self.open_order(file_name)

    def open_selected_document(self):
        file_name = self._selected_history_file()
        if not file_name:
            return
        pdf_path, _png_path = document_paths(file_name)
        if not os.path.isfile(pdf_path):
            if not messagebox.askyesno(
                "No document yet",
                "No PDF is stored for this P.O.\n\nGenerate it now?",
            ):
                return
            order = read_order(file_name)
            if order is None:
                messagebox.showerror("Open failed", f"Could not read:\n\n{file_name}")
                self._refresh_history()
                return
            self.status_var.set("Generating document...")
            self.update_idletasks()
            try:
                pdf_path, png_path, renderer = generate_documents(
                    order, out_dir=DOCUMENTS_DIR, base_name=file_name[:-5]
                )
            except Exception as error:
                messagebox.showerror("Generation failed", str(error))
                return
            record = read_order(file_name) or {}
            record["document"] = {
                "pdf": os.path.basename(pdf_path),
                "png": os.path.basename(png_path),
                "renderer": renderer,
                "updated": datetime.datetime.now().isoformat(timespec="seconds"),
            }
            try:
                write_order(record, file_name)
            except OSError:
                pass
            self._refresh_history()
            self.status_var.set(f"Generated: {os.path.basename(pdf_path)}")
        try:
            os.startfile(pdf_path)
        except OSError as error:
            messagebox.showerror("Open failed", f"Could not open the PDF:\n{error}")

    def open_order(self, file_name):
        order = read_order(file_name)
        if order is None:
            messagebox.showerror("Open failed", f"Could not read:\n\n{file_name}")
            self._refresh_history()
            return
        self.load_order_into_form(order, file_name)
        self.status_var.set(f"Opened saved P.O.: {file_name}")

    def load_order_into_form(self, order, file_name):
        self.po_var.set(order.get("po_no", ""))
        self.name_var.set(order.get("name", ""))
        self.order_date_var.set(order.get("order_date", ""))
        self.delivery_date_var.set(order.get("delivery_date", ""))
        self.approved_by_var.set(order.get("approved_by", ""))
        self.approved_date_var.set(order.get("approved_date", ""))
        self.verified_by_var.set(order.get("verified_by", ""))
        self.verified_date_var.set(order.get("verified_date", ""))
        self.received_by_var.set(order.get("received_by", ""))
        self.received_date_var.set(order.get("received_date", ""))

        self.item_rows = []
        for item in order.get("items", []):
            self.add_item_row()
            row = self.item_rows[-1]
            qty = item.get("qty")
            rate = item.get("rate")
            row["item_no"].set(item.get("item_no", ""))
            row["description"].set(item.get("description", ""))
            row["qty"].set(fmt_qty(qty) if qty is not None else "")
            row["rate"].set(fmt_money(rate) if rate is not None else "")
            row["uom"].set(item.get("uom", ""))
        while len(self.item_rows) < 3:
            self.add_item_row()

        self.current_order_file = file_name
        self.current_order_meta = {"po_no": order.get("po_no", ""), "name": order.get("name", "")}
        self._set_dirty(False)
        self._update_row_amount(None)

    def duplicate_selected_order(self):
        file_name = self._selected_history_file()
        if not file_name:
            return
        order = read_order(file_name)
        if order is None:
            messagebox.showerror("Duplicate failed", f"Could not read:\n\n{file_name}")
            self._refresh_history()
            return
        new_file = next_copy_file_name(order.get("po_no", ""), order.get("name", ""))
        copy = dict(order)
        now = datetime.datetime.now().isoformat(timespec="seconds")
        copy["created"] = now
        copy["modified"] = now
        try:
            write_order(copy, new_file)
        except OSError as error:
            messagebox.showerror("Duplicate failed", str(error))
            return
        source_pdf, source_png = document_paths(file_name)
        copy_pdf, copy_png = document_paths(new_file)
        if os.path.isfile(source_pdf):
            try:
                os.makedirs(DOCUMENTS_DIR, exist_ok=True)
                shutil.copy2(source_pdf, copy_pdf)
                shutil.copy2(source_png, copy_png)
                copy["document"] = {
                    "pdf": os.path.basename(copy_pdf),
                    "png": os.path.basename(copy_png),
                    "updated": now,
                }
                write_order(copy, new_file)
            except OSError:
                pass
        self._refresh_history()
        self.status_var.set(f"Duplicated as {new_file}")

    def delete_selected_order(self):
        file_name = self._selected_history_file()
        if not file_name:
            return
        if not messagebox.askyesno(
            "Delete saved P.O.",
            f"Delete this saved P.O. and its stored PDF/PNG?\n\n{file_name}",
        ):
            return
        delete_order(file_name)
        delete_order_documents(file_name)
        if self.current_order_file == file_name:
            self.current_order_file = None
            self.current_order_meta = {}
        self._refresh_history()
        self.status_var.set(f"Deleted {file_name}")

    def _on_close(self):
        if self.dirty:
            answer = messagebox.askyesnocancel("Unsaved changes", "Save this P.O. before closing?")
            if answer is None:
                return
            if answer:
                result = self.save_current()
                if not isinstance(result, tuple):
                    return
        if getattr(self, "_session_temp_dir", None):
            shutil.rmtree(self._session_temp_dir, ignore_errors=True)
        self.destroy()

    def _session_tmp(self):
        if not getattr(self, "_session_temp_dir", None):
            self._session_temp_dir = tempfile.mkdtemp(prefix="po_session_")
        return self._session_temp_dir

    def _ensure_outputs(self):
        data = self.collect_data()
        if data is None:
            return None
        outcome = self.save_current(data)
        if isinstance(outcome, tuple):
            _json_path, pdf_path, png_path = outcome
        else:
            self.status_var.set("Generating document...")
            self.update_idletasks()
            try:
                pdf_path, png_path, _renderer = generate_documents(data, out_dir=self._session_tmp())
            except Exception as error:
                self.status_var.set("Failed to generate document.")
                messagebox.showerror("Generation failed", str(error))
                return None
        self.last_pdf = pdf_path
        self.last_png = png_path
        self.status_var.set(f"Document: {os.path.basename(pdf_path)}  |  {os.path.basename(png_path)}")
        return pdf_path, png_path

    def generate_documents_clicked(self):
        outputs = self._ensure_outputs()
        if outputs:
            pdf_path, png_path = outputs
            messagebox.showinfo(
                "Document generated",
                "Saved with the P.O. library:\n\n"
                f"{os.path.basename(pdf_path)}\n{os.path.basename(png_path)}\n\n"
                f"Location:\n{DOCUMENTS_DIR}",
            )

    def preview_document(self):
        outputs = self._ensure_outputs()
        if not outputs:
            return
        try:
            os.startfile(outputs[0])
        except OSError as error:
            messagebox.showerror("Preview failed", f"Could not open the PDF:\n{error}")

    def print_document(self):
        outputs = self._ensure_outputs()
        if not outputs:
            return
        try:
            os.startfile(outputs[0], "print")
            self.status_var.set("Printing sent to the default printer.")
        except OSError:
            try:
                os.startfile(outputs[1])
                messagebox.showinfo(
                    "Print",
                    "No direct print handler was found.\nThe document image was opened instead - "
                    "print it from there, or use Preview and print from your PDF reader.",
                )
            except OSError as error:
                messagebox.showerror("Print failed", str(error))

    def open_library_folder(self):
        os.makedirs(LIBRARY_DIR, exist_ok=True)
        os.startfile(LIBRARY_DIR)

    def clear_form(self):
        if not messagebox.askyesno("Clear form", "Clear all fields and start over?"):
            return
        self.po_var.set("")
        self.name_var.set("")
        self.order_date_var.set(datetime.date.today().strftime("%m/%d/%Y"))
        self.delivery_date_var.set("")
        self.approved_by_var.set(DEFAULT_APPROVER)
        self.approved_date_var.set("")
        self.verified_by_var.set("")
        self.verified_date_var.set("")
        self.received_by_var.set("")
        self.received_date_var.set("")
        self.item_rows = []
        for _ in range(3):
            self.add_item_row()
        self.current_order_file = None
        self.current_order_meta = {}
        self._set_dirty(False)
        self.status_var.set("Form cleared.")


def run_selftest():
    global ORDERS_DIR, DOCUMENTS_DIR, BROWSER_PROFILE_DIR, RENDER_LOG_PATH
    temp_root = tempfile.mkdtemp(prefix="po_selftest_")
    ORDERS_DIR = os.path.join(temp_root, "orders")
    DOCUMENTS_DIR = os.path.join(temp_root, "documents")
    BROWSER_PROFILE_DIR = os.path.join(temp_root, ".browser")
    RENDER_LOG_PATH = os.path.join(temp_root, "po_render.log")
    try:
        data = {
            "po_no": "26-0003",
            "name": "STEEL WORLD MANUFACTURING CORP.",
            "order_date": "1/28/2026",
            "delivery_date": "1/29/2026",
            "items": [
                {"item_no": "GI WIRE 8", "description": "G.I. WIRE #8 (ROLL)", "qty": 1, "rate": 1925.00, "uom": "Roll"},
            ],
            "approved_by": DEFAULT_APPROVER,
            "approved_date": "",
            "verified_by": "",
            "verified_date": "",
            "received_by": "",
            "received_date": "",
        }
        pdf_path, png_path, renderer = generate_documents(data)
        print(pdf_path)
        print(png_path)
        print("renderer:", renderer)

        trusted_find_browsers = globals()["find_browsers"]
        globals()["find_browsers"] = lambda: []
        try:
            native_pdf, native_png, native_renderer = generate_documents(
                data, base_name="selftest_native"
            )
        finally:
            globals()["find_browsers"] = trusted_find_browsers
        if native_renderer != "builtin" or not valid_pdf(native_pdf) or not valid_png(native_png):
            raise RuntimeError("Built-in renderer self-test failed")
        os.remove(native_pdf)
        os.remove(native_png)
        print("built-in renderer OK")

        file_name = order_file_name(data["po_no"], data["name"])
        write_order(normalize_order(data), file_name)
        stored = read_order(file_name)
        if not stored or stored.get("po_no") != data["po_no"] or order_total(stored) != 1925.0:
            raise RuntimeError("Order store roundtrip failed")
        expected_pdf, expected_png = document_paths(file_name)
        if pdf_path != expected_pdf or png_path != expected_png:
            raise RuntimeError("Document paths do not match the record name")
        if not (has_document(file_name) and os.path.isfile(expected_png)):
            raise RuntimeError("Stored document files are missing")
        delete_order(file_name)
        delete_order_documents(file_name)
        if os.path.exists(expected_pdf) or os.path.exists(expected_png):
            raise RuntimeError("Document cleanup failed")
        print("store OK")

        library_dir = resolve_library_dir()
        os.makedirs(library_dir, exist_ok=True)
        probe = os.path.join(library_dir, ".write_test")
        with open(probe, "w", encoding="utf-8") as handle:
            handle.write("ok")
        os.remove(probe)
        print("library OK:", library_dir)
    finally:
        shutil.rmtree(temp_root, ignore_errors=True)


def main():
    if "--selftest" in sys.argv:
        run_selftest()
        return
    try:
        from ctypes import windll

        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    migrate_legacy_orders()
    cleanup_stale_render_dirs()
    app = PurchaseOrderApp()
    app.mainloop()


if __name__ == "__main__":
    main()
