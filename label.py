#!/usr/bin/env python3
"""
Simple label GUI for Brother QL-810W using brother_ql CLI.

Features:
- Full-width cut button grid (5 buttons per row)
- Edit DB window (add/edit/delete/reorder/save/load)
- Session window for farm/session metadata
- Live preview updates on field changes
"""

import json
import os
import re
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk

from PIL import Image, ImageDraw, ImageFont, ImageTk
import qrcode

try:
    import serial
except Exception:
    serial = None

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")
DEFAULT_WIDGET_SCALING = 1.35
ctk.set_widget_scaling(DEFAULT_WIDGET_SCALING)

PRINTER_MODEL = "QL-810W"
CONFIG_DIR = Path(__file__).with_name("config")
CONFIG_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_PATH = CONFIG_DIR / "printer_config.json"
CUT_DB_PATH = CONFIG_DIR / "cuts_db.json"
ANALYTICS_LOG_PATH = Path(__file__).with_name("print_log.jsonl")
LABEL_CONFIG_PATH = CONFIG_DIR / "label_config.json"
SESSION_DEFAULT_PATH = CONFIG_DIR / "session_default.json"
PRINTED_LABELS_DIR = Path(__file__).with_name("printed_labels")
LOG_ARCHIVE_DIR = Path(__file__).with_name("logs")

SUPPORTED_USB_BACKENDS = {"pyusb"}

LABEL_SIZE = "62x100"
LABEL_WIDTH = 1109
LABEL_HEIGHT = 696

SCALE_PORT = "/dev/ttyUSB0"
SCALE_BAUDRATE = 9600
SCALE_DEFAULT_STABLE_ITERATIONS = 10
SCALE_RECONNECT_DELAY_SECONDS = 2.0
SCALE_VALUE_RE = re.compile(r"([-+]?\d+(?:\.\d+)?)\s*kg", re.IGNORECASE)
MIN_PRINT_WEIGHT_KG = 0.01
PRINT_RETRY_ATTEMPTS = 4
PRINT_RETRY_BASE_DELAY_SECONDS = 0.35
MIN_WIDGET_SCALING = 0.80
MAX_WIDGET_SCALING = 2.20
WIDGET_SCALING_STEP = 0.05
FILENAME_PART_RE = re.compile(r"[^A-Za-z0-9._-]+")

SESSION_FIELDS = [
    ("farm_name", "Name of Farm"),
    ("logo_path", "Logo"),
    ("animal_number", "Animal Number"),
    ("farm_number", "Farm Number"),
    ("due_date_4_7", "Due date 4-7"),
    ("due_date_frozen", "Due date frozen"),
    ("birth_country", "Birth Country"),
    ("life_country", "Life Country"),
    ("slaughter_country", "Slaugther Country"),
    ("packaged_country", "Packaged Country"),
]

LABEL_FIELD_DEFS = [
    ("cut_name", "Cut"),
    ("weight_kg", "Weight KG"),
    ("price_per_kg", "Price / KG"),
    ("tax", "Tax"),
    ("total_price", "Total price"),
    ("farm_name", "Farm"),
    ("logo_path", "Logo"),
    ("animal_number", "Animal Number"),
    ("farm_number", "Farm Number"),
    ("due_date_4_7", "Due date 4-7"),
    ("due_date_frozen", "Due date frozen"),
    ("birth_country", "Birth Country"),
    ("life_country", "Life Country"),
    ("slaughter_country", "Slaugther Country"),
    ("packaged_country", "Packaged Country"),
]
LABEL_FIELD_LABELS = {key: label for key, label in LABEL_FIELD_DEFS}
EMPTY_LINE_KEY_PREFIX = "__empty_line__"
EMPTY_LINE_LABEL = "Free Text"
DEFAULT_LABEL_LINE_SPACING = 8


def default_label_field_config():
    defaults = []
    for key, label in LABEL_FIELD_DEFS:
        defaults.append(
            {
                "key": key,
                "print_name": label,
                "show": key != "logo_path",
                "font_size": 24,
            }
        )

    for entry in defaults:
        if entry["key"] == "cut_name":
            entry["font_size"] = 52
        elif entry["key"] in {"weight_kg", "total_price"}:
            entry["font_size"] = 34
        elif entry["key"] == "farm_name":
            entry["font_size"] = 28

    return defaults


def is_empty_line_key(key):
    return str(key).startswith(EMPTY_LINE_KEY_PREFIX)


def make_empty_line_entry(key=None):
    if not key:
        key = f"{EMPTY_LINE_KEY_PREFIX}{time.time_ns()}"
    return {
        "key": key,
        "print_name": "",
        "show": True,
        "font_size": 24,
    }


def normalize_label_line_spacing(raw_value):
    try:
        value = int(raw_value)
    except (TypeError, ValueError):
        value = DEFAULT_LABEL_LINE_SPACING
    return max(0, min(120, value))


def normalize_label_field_config(config_items):
    defaults_by_key = {entry["key"]: entry for entry in default_label_field_config()}
    normalized = []
    seen_defaults = set()
    seen_custom = set()

    if isinstance(config_items, list):
        for raw in config_items:
            if not isinstance(raw, dict):
                continue

            key = str(raw.get("key", "")).strip()
            if not key:
                continue

            if key in defaults_by_key:
                if key in seen_defaults:
                    continue
                seen_defaults.add(key)
                default_entry = defaults_by_key[key]
            elif is_empty_line_key(key):
                if key in seen_custom:
                    continue
                seen_custom.add(key)
                default_entry = make_empty_line_entry(key=key)
            else:
                continue

            if "print_name" in raw:
                print_name = str(raw.get("print_name", "")).strip()
            else:
                print_name = str(default_entry.get("print_name", "")).strip()
            show = bool(raw.get("show", default_entry.get("show", True)))
            try:
                font_size = int(raw.get("font_size", default_entry.get("font_size", 24)))
            except (TypeError, ValueError):
                font_size = int(default_entry.get("font_size", 24))
            font_size = max(8, min(120, font_size))

            normalized.append(
                {
                    "key": key,
                    "print_name": print_name,
                    "show": show,
                    "font_size": font_size,
                }
            )

    for key, _label in LABEL_FIELD_DEFS:
        if key not in seen_defaults:
            default_entry = defaults_by_key[key]
            normalized.append(
                {
                    "key": key,
                    "print_name": default_entry["print_name"],
                    "show": default_entry["show"],
                    "font_size": default_entry["font_size"],
                }
            )

    return normalized


def parse_label_config_payload(data):
    if isinstance(data, dict):
        config_items = data.get("fields", [])
        line_spacing = normalize_label_line_spacing(data.get("line_spacing", DEFAULT_LABEL_LINE_SPACING))
    else:
        config_items = data
        line_spacing = DEFAULT_LABEL_LINE_SPACING

    return normalize_label_field_config(config_items), line_spacing


def load_label_field_config(path: Path):
    if not path.exists():
        return default_label_field_config(), DEFAULT_LABEL_LINE_SPACING

    with path.open("r", encoding="utf-8") as f:
        data = json.load(f)
    return parse_label_config_payload(data)


def save_label_field_config(path: Path, config_items, line_spacing):
    normalized = normalize_label_field_config(config_items)
    spacing = normalize_label_line_spacing(line_spacing)
    payload = {
        "line_spacing": spacing,
        "fields": normalized,
    }
    with path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)


def load_printer_connection(config_path: Path):
    with config_path.open("r", encoding="utf-8") as f:
        config = json.load(f)

    mode = str(config.get("connection_mode", "usb")).strip().lower()
    if mode != "usb":
        raise ValueError("Config error: connection_mode must be 'usb'.")

    usb_cfg = config.get("usb", {})
    backend = str(usb_cfg.get("backend", "pyusb")).strip()
    identifier = str(usb_cfg.get("identifier", "")).strip()

    if backend not in SUPPORTED_USB_BACKENDS:
        raise ValueError(
            "Config error: usb.backend must be one of "
            f"{sorted(SUPPORTED_USB_BACKENDS)}."
        )
    if not identifier:
        raise ValueError("Config error: usb.identifier must be set for usb mode.")

    return backend, identifier


BACKEND, PRINTER_IDENTIFIER = load_printer_connection(CONFIG_PATH)


def normalize_item(raw_item):
    legacy_weight_as_price = str(raw_item.get("weight_kg", "")).strip()
    price_per_kg = str(raw_item.get("price_per_kg", "")).strip()
    if not price_per_kg:
        price_per_kg = legacy_weight_as_price

    return {
        "cut_name": str(raw_item.get("cut_name", "")).strip(),
        "price_per_kg": price_per_kg,
        "tax": str(raw_item.get("tax", "")).strip(),
    }


def load_cut_items(path: Path):
    if not path.exists():
        return []

    with path.open("r", encoding="utf-8") as f:
        data = json.load(f)

    if not isinstance(data, list):
        raise ValueError("DB file must contain a JSON list.")

    items = []
    for entry in data:
        if isinstance(entry, dict):
            item = normalize_item(entry)
            if item["cut_name"]:
                items.append(item)
    return items


def save_cut_items(path: Path, items):
    serializable = [normalize_item(item) for item in items]
    with path.open("w", encoding="utf-8") as f:
        json.dump(serializable, f, indent=2, ensure_ascii=False)


FONT_CACHE = {}
FONT_CANDIDATES = [
    "OpenSans-Light.ttf",
    "OpenSans-Regular.ttf",
    "DejaVuSans.ttf",
    "LiberationSans-Regular.ttf",
    "NotoSans-Regular.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/liberation/LiberationSans-Regular.ttf",
    "/usr/share/fonts/google-noto/NotoSans-Regular.ttf",
]

try:
    RESAMPLE_NEAREST = Image.Resampling.NEAREST
    RESAMPLE_LANCZOS = Image.Resampling.LANCZOS
except AttributeError:
    RESAMPLE_NEAREST = Image.NEAREST
    RESAMPLE_LANCZOS = Image.LANCZOS


def load_font(size: int):
    size = int(size)
    if size in FONT_CACHE:
        return FONT_CACHE[size]

    font = None
    for candidate in FONT_CANDIDATES:
        try:
            font = ImageFont.truetype(candidate, size)
            break
        except OSError:
            continue

    if font is None:
        font = ImageFont.load_default()

    FONT_CACHE[size] = font
    return font


FONT_TITLE = load_font(54)
FONT_MAIN = load_font(34)
FONT_SMALL = load_font(24)


def build_label_image(label_values, label_field_config, line_spacing):
    img = Image.new("RGB", (LABEL_WIDTH, LABEL_HEIGHT), "white")
    draw = ImageDraw.Draw(img)

    margin = 28
    logo_box_w = 250
    logo_box_h = 180

    config_by_key = {
        str(entry.get("key", "")).strip(): entry
        for entry in label_field_config
        if isinstance(entry, dict)
    }

    logo_entry = config_by_key.get("logo_path", {})
    show_logo = bool(logo_entry.get("show", False))
    logo_path = str(label_values.get("logo_path", "")).strip()
    if show_logo and logo_path and os.path.exists(logo_path):
        try:
            logo = Image.open(logo_path).convert("RGBA")
            logo.thumbnail((logo_box_w, logo_box_h), RESAMPLE_LANCZOS)
            logo_x = LABEL_WIDTH - margin - logo.width
            logo_y = margin
            img.paste(logo, (logo_x, logo_y), logo)
        except OSError:
            pass

    spacing = normalize_label_line_spacing(line_spacing)
    y = margin
    for entry in label_field_config:
        if not entry.get("show", True):
            continue

        key = str(entry.get("key", "")).strip()
        try:
            font_size = int(entry.get("font_size", 24))
        except (TypeError, ValueError):
            font_size = 24
        font_size = max(8, min(120, font_size))

        if is_empty_line_key(key):
            free_text = str(entry.get("print_name", "")).strip()
            if free_text:
                draw.text((margin, y), free_text, fill="black", font=load_font(font_size))
            y += font_size + spacing
            continue

        if key == "logo_path":
            continue

        value = str(label_values.get(key, "")).strip()
        if not value:
            continue

        print_name = str(entry.get("print_name", LABEL_FIELD_LABELS.get(key, key))).strip()
        font = load_font(font_size)

        line = f"{print_name}: {value}" if print_name else value
        draw.text((margin, y), line, fill="black", font=font)
        y += font_size + spacing

    qr_payload = {k: v for k, v in label_values.items() if k != "logo_path"}
    qr_text = json.dumps(qr_payload, ensure_ascii=False)

    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=1,
    )
    qr.add_data(qr_text)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")

    qr_size = 240
    qr_img = qr_img.resize((qr_size, qr_size), RESAMPLE_NEAREST)
    qr_x = LABEL_WIDTH - margin - qr_size
    qr_y = LABEL_HEIGHT - margin - qr_size
    img.paste(qr_img, (qr_x, qr_y))

    return img


def sanitize_filename_part(value, fallback):
    text = str(value).strip()
    if not text:
        return fallback
    text = FILENAME_PART_RE.sub("_", text)
    text = re.sub(r"_+", "_", text).strip("._")
    return text or fallback


def build_printed_label_path(cut_name, weight):
    PRINTED_LABELS_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    cut_part = sanitize_filename_part(cut_name, "cut")
    if weight is None:
        weight_part = "unknownkg"
    else:
        weight_part = sanitize_filename_part(f"{weight:.4f}kg", "unknownkg")

    base_name = f"{timestamp}_{cut_part}_{weight_part}"
    target = PRINTED_LABELS_DIR / f"{base_name}.png"
    suffix = 2
    while target.exists():
        target = PRINTED_LABELS_DIR / f"{base_name}_{suffix}.png"
        suffix += 1
    return target


def is_resource_busy_error(error_text):
    text = str(error_text).lower()
    return (
        "resource busy" in text
        or "errno 16" in text
        or "usb.core.usberror" in text
    )


def print_via_brother_cli(pil_image, printer_identifier, image_path: Path, cut_paper: bool):
    image_path.parent.mkdir(parents=True, exist_ok=True)
    pil_image.save(image_path, format="PNG")

    cmd = [
        sys.executable,
        "-m",
        "brother_ql.cli",
        "-b",
        "pyusb",
        "-m",
        PRINTER_MODEL,
        "-p",
        printer_identifier,
        "print",
        "-l",
        LABEL_SIZE,
    ]
    if not cut_paper:
        cmd.append("--no-cut")

    cmd.append(str(image_path))
    #print(f"Would call:{cmd}")
    #Disabled so I don't print such much right now, but you can uncomment to enable CLI printing.
    last_output = ""
    for attempt in range(1, PRINT_RETRY_ATTEMPTS + 1):
        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=False,
        )

        if result.returncode == 0:
            return

        output = (result.stderr or "").strip() or (result.stdout or "").strip()
        last_output = output or "brother_ql CLI print failed."

        if attempt < PRINT_RETRY_ATTEMPTS and is_resource_busy_error(last_output):
            time.sleep(PRINT_RETRY_BASE_DELAY_SECONDS * attempt)
            continue

        raise RuntimeError(last_output)

    raise RuntimeError(last_output or "brother_ql CLI print failed.")


def send_to_printer(pil_image, image_path: Path, cut_paper: bool):
    if BACKEND != "pyusb":
        raise ValueError("Only usb+pyusb is supported by this app.")
    print_via_brother_cli(pil_image, PRINTER_IDENTIFIER, image_path, cut_paper)


def open_scale_serial(port=SCALE_PORT, baudrate=SCALE_BAUDRATE):
    if serial is None:
        raise RuntimeError("pyserial is not installed. Install it with: pip install pyserial")

    return serial.Serial(
        port,
        baudrate,
        bytesize=serial.EIGHTBITS,
        parity=serial.PARITY_NONE,
        stopbits=serial.STOPBITS_ONE,
        xonxoff=False,
        rtscts=False,
        timeout=0.5,
    )


class ItemEditorDialog(ctk.CTkToplevel):
    def __init__(self, parent, title, item=None):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.result = None

        self.cut_var = tk.StringVar(value=(item or {}).get("cut_name", ""))
        self.price_var = tk.StringVar(value=(item or {}).get("price_per_kg", ""))
        self.tax_var = tk.StringVar(value=(item or {}).get("tax", ""))

        body = ctk.CTkFrame(self)
        body.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        ctk.CTkLabel(body, text="Cut name").grid(row=0, column=0, sticky="w", pady=(0, 6))
        ctk.CTkEntry(body, textvariable=self.cut_var, width=32).grid(
            row=0, column=1, sticky="ew", pady=(0, 6)
        )

        ctk.CTkLabel(body, text="Price / KG").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ctk.CTkEntry(body, textvariable=self.price_var, width=32).grid(
            row=1, column=1, sticky="ew", pady=(6, 0)
        )

        ctk.CTkLabel(body, text="Tax").grid(row=2, column=0, sticky="w", pady=(6, 0))
        ctk.CTkEntry(body, textvariable=self.tax_var, width=32).grid(
            row=2, column=1, sticky="ew", pady=(6, 0)
        )

        btns = ctk.CTkFrame(body)
        btns.grid(row=3, column=0, columnspan=2, pady=(10, 0), sticky="e")
        ctk.CTkButton(btns, text="Cancel", command=self.on_cancel).grid(row=0, column=0, padx=(0, 6))
        ctk.CTkButton(btns, text="Save", command=self.on_save).grid(row=0, column=1)

        self.bind("<Return>", lambda _e: self.on_save())
        self.bind("<Escape>", lambda _e: self.on_cancel())
        self.transient(parent)
        self.after(0, self._activate_modal)

    def _activate_modal(self, attempts_left=20):
        if not self.winfo_exists():
            return

        try:
            self.lift()
            self.focus_force()
            self.grab_set()
        except tk.TclError:
            if attempts_left > 0:
                self.after(25, lambda: self._activate_modal(attempts_left - 1))

    def on_save(self):
        cut_name = self.cut_var.get().strip()
        price = self.price_var.get().strip()
        tax = self.tax_var.get().strip()

        if not cut_name:
            messagebox.showerror("Invalid item", "Cut name cannot be empty.", parent=self)
            return

        self.result = {
            "cut_name": cut_name,
            "price_per_kg": price,
            "tax": tax,
        }
        self.destroy()

    def on_cancel(self):
        self.result = None
        self.destroy()


class DatabaseEditorWindow(ctk.CTkToplevel):
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Edit DB")
        self.geometry("680x420")
        self.controls_layout_job = None
        self.control_buttons = []
        self.last_controls_layout = None
        self.last_controls_width = None

        main = ctk.CTkFrame(self)
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        main.rowconfigure(0, weight=1)
        main.columnconfigure(0, weight=1)

        list_frame = ctk.CTkFrame(main)
        list_frame.grid(row=0, column=0, sticky="nsew")
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        self.listbox = tk.Listbox(list_frame)
        self.listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.listbox.configure(yscrollcommand=scrollbar.set)

        self.controls_frame = ctk.CTkFrame(main)
        self.controls_frame.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        self.controls_frame.bind("<Configure>", lambda _e: self.schedule_controls_layout())

        button_specs = [
            ("Add", self.on_add),
            ("Edit", self.on_edit),
            ("Delete", self.on_delete),
            ("Move Up", self.on_move_up),
            ("Move Down", self.on_move_down),
            ("Save List", self.on_save_list),
            ("Load List", self.on_load_list),
            ("Close", self.destroy),
        ]
        for text, command in button_specs:
            btn = ctk.CTkButton(self.controls_frame, text=text, command=command)
            self.control_buttons.append(btn)

        self.after(0, self.layout_controls)

        self.refresh_list()

    def schedule_controls_layout(self):
        current_width = self.controls_frame.winfo_width()
        if self.last_controls_width == current_width and self.last_controls_layout is not None:
            return

        if self.controls_layout_job is not None:
            try:
                self.after_cancel(self.controls_layout_job)
            except Exception:
                pass
        self.controls_layout_job = self.after(40, self.layout_controls)

    def layout_controls(self):
        self.controls_layout_job = None
        frame_width = self.controls_frame.winfo_width()
        if frame_width <= 1:
            self.after(40, self.layout_controls)
            return

        row = 0
        col = 0
        used_width = 0
        pad_x = 6
        pad_y = 6
        layout = []

        for btn in self.control_buttons:
            btn.update_idletasks()
            btn_width = btn.winfo_reqwidth()
            needed = btn_width if col == 0 else btn_width + pad_x

            if col > 0 and used_width + needed > frame_width:
                row += 1
                col = 0
                used_width = 0
                needed = btn_width

            layout.append((row, col))
            used_width += needed
            col += 1

        if layout == self.last_controls_layout:
            self.last_controls_width = frame_width
            return

        self.last_controls_layout = layout
        self.last_controls_width = frame_width

        for btn in self.control_buttons:
            btn.grid_forget()

        for btn, (btn_row, btn_col) in zip(self.control_buttons, layout):
            btn.grid(row=btn_row, column=btn_col, sticky="w", padx=(0, pad_x), pady=(0, pad_y))

    def selected_index(self):
        sel = self.listbox.curselection()
        if not sel:
            return None
        return sel[0]

    def refresh_list(self, select_index=None):
        self.listbox.delete(0, tk.END)
        for idx, item in enumerate(self.app.items, start=1):
            tax_text = str(item.get("tax", "")).strip()
            tax_suffix = f" | Tax {tax_text}" if tax_text else ""
            self.listbox.insert(
                tk.END,
                f"{idx:02d}. {item.get('cut_name', '')} | "
                f"{item.get('price_per_kg', '')} /KG{tax_suffix}",
            )

        if select_index is not None and self.app.items:
            select_index = max(0, min(select_index, len(self.app.items) - 1))
            self.listbox.selection_set(select_index)
            self.listbox.see(select_index)

    def on_add(self):
        dialog = ItemEditorDialog(self, "Add item")
        self.wait_window(dialog)
        if dialog.result:
            self.app.items.append(dialog.result)
            self.app.refresh_item_buttons()
            self.app.save_default_db()
            self.refresh_list(select_index=len(self.app.items) - 1)

    def on_edit(self):
        idx = self.selected_index()
        if idx is None:
            return

        dialog = ItemEditorDialog(self, "Edit item", self.app.items[idx])
        self.wait_window(dialog)
        if dialog.result:
            self.app.items[idx] = dialog.result
            self.app.refresh_item_buttons()
            self.app.save_default_db()
            self.refresh_list(select_index=idx)

    def on_delete(self):
        idx = self.selected_index()
        if idx is None:
            return
        if not messagebox.askyesno("Delete item", "Delete selected item?", parent=self):
            return

        del self.app.items[idx]
        self.app.refresh_item_buttons()
        self.app.save_default_db()
        self.refresh_list(select_index=idx)

    def on_move_up(self):
        idx = self.selected_index()
        if idx is None or idx == 0:
            return

        self.app.items[idx - 1], self.app.items[idx] = self.app.items[idx], self.app.items[idx - 1]
        self.app.refresh_item_buttons()
        self.app.save_default_db()
        self.refresh_list(select_index=idx - 1)

    def on_move_down(self):
        idx = self.selected_index()
        if idx is None or idx >= len(self.app.items) - 1:
            return

        self.app.items[idx + 1], self.app.items[idx] = self.app.items[idx], self.app.items[idx + 1]
        self.app.refresh_item_buttons()
        self.app.save_default_db()
        self.refresh_list(select_index=idx + 1)

    def on_save_list(self):
        file_path = filedialog.asksaveasfilename(
            parent=self,
            title="Save cut list",
            defaultextension=".json",
            filetypes=[("JSON", "*.json"), ("All files", "*.*")],
        )
        if not file_path:
            return
        try:
            save_cut_items(Path(file_path), self.app.items)
        except Exception as e:
            messagebox.showerror("Save failed", f"Could not save list:\n{e}", parent=self)

    def on_load_list(self):
        file_path = filedialog.askopenfilename(
            parent=self,
            title="Load cut list",
            filetypes=[("JSON", "*.json"), ("All files", "*.*")],
        )
        if not file_path:
            return

        try:
            loaded = load_cut_items(Path(file_path))
        except Exception as e:
            messagebox.showerror("Load failed", f"Could not load list:\n{e}", parent=self)
            return

        self.app.items = loaded
        self.app.refresh_item_buttons()
        self.app.save_default_db()
        self.refresh_list(select_index=0)


class SessionWindow(ctk.CTkToplevel):
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Session")
        self.resizable(False, False)

        main = ctk.CTkFrame(self)
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        for row_idx, (key, label_text) in enumerate(SESSION_FIELDS):
            ctk.CTkLabel(main, text=label_text).grid(row=row_idx, column=0, sticky="w", pady=3)
            entry = ctk.CTkEntry(main, textvariable=self.app.session_vars[key], width=44)
            entry.grid(row=row_idx, column=1, sticky="ew", pady=3)

            if key == "logo_path":
                ctk.CTkButton(
                    main,
                    text="Browse",
                    command=self.on_browse_logo,
                ).grid(row=row_idx, column=2, padx=(6, 0), pady=3)

        button_row = ctk.CTkFrame(main)
        button_row.grid(row=len(SESSION_FIELDS), column=0, columnspan=3, pady=(10, 0), sticky="e")
        ctk.CTkButton(button_row, text="Load Session", command=self.on_load_session).grid(
            row=0, column=0, padx=(0, 6)
        )
        ctk.CTkButton(button_row, text="Save Session", command=self.on_save_session).grid(
            row=0, column=1, padx=(0, 6)
        )
        ctk.CTkButton(button_row, text="Close", command=self.destroy).grid(row=0, column=2)

    def on_browse_logo(self):
        file_path = filedialog.askopenfilename(
            parent=self,
            title="Select logo image",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.bmp *.gif"),
                ("All files", "*.*"),
            ],
        )
        if file_path:
            self.app.session_vars["logo_path"].set(file_path)

    def on_save_session(self):
        file_path = filedialog.asksaveasfilename(
            parent=self,
            title="Save session",
            defaultextension=".json",
            filetypes=[("JSON", "*.json"), ("All files", "*.*")],
        )
        if not file_path:
            return

        try:
            self.app.save_session_to_file(Path(file_path))
        except Exception as e:
            messagebox.showerror("Save failed", f"Could not save session:\n{e}", parent=self)

    def on_load_session(self):
        file_path = filedialog.askopenfilename(
            parent=self,
            title="Load session",
            filetypes=[("JSON", "*.json"), ("All files", "*.*")],
        )
        if not file_path:
            return

        try:
            self.app.load_session_from_file(Path(file_path))
        except Exception as e:
            messagebox.showerror("Load failed", f"Could not load session:\n{e}", parent=self)


class LabelConfigWindow(ctk.CTkToplevel):
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Configure label")
        self.geometry("1020x680")

        self.row_models = []
        self.row_widgets = []
        self.line_spacing_var = tk.StringVar(value=str(self.app.label_line_spacing))

        main = ctk.CTkFrame(self)
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)
        main.columnconfigure(0, weight=1)

        controls = ctk.CTkFrame(main)
        controls.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ctk.CTkButton(controls, text="Apply", command=self.on_apply).grid(row=0, column=0, padx=(0, 6))
        ctk.CTkButton(controls, text="Save JSON", command=self.on_save_json).grid(row=0, column=1, padx=(0, 6))
        ctk.CTkButton(controls, text="Load JSON", command=self.on_load_json).grid(row=0, column=2, padx=(0, 6))
        ctk.CTkButton(controls, text="Reset Default", command=self.on_reset_default).grid(
            row=0, column=3, padx=(0, 6)
        )
        ctk.CTkButton(controls, text="Add Free Text", command=self.on_add_empty_line).grid(
            row=0, column=4, padx=(0, 6)
        )
        ctk.CTkButton(controls, text="Close", command=self.destroy).grid(row=0, column=5)
        ctk.CTkLabel(controls, text="Line spacing").grid(row=0, column=6, padx=(16, 6))
        ctk.CTkEntry(controls, textvariable=self.line_spacing_var, width=72).grid(row=0, column=7)

        self.scroll_frame = ctk.CTkScrollableFrame(main)
        self.scroll_frame.grid(row=1, column=0, sticky="nsew")
        self.scroll_frame.columnconfigure(1, weight=1)

        self.set_config(self.app.label_field_config, self.app.label_line_spacing)

    def set_config(self, config_items, line_spacing=None):
        self.row_models = []
        if line_spacing is not None:
            self.line_spacing_var.set(str(normalize_label_line_spacing(line_spacing)))
        normalized = normalize_label_field_config(config_items)
        for entry in normalized:
            model = {
                "key": entry["key"],
                "print_name": str(entry.get("print_name", "")),
                "show": bool(entry.get("show", True)),
                "font_size": str(entry.get("font_size", 24)),
            }
            self.row_models.append(model)
        self.rebuild_rows()

    def rebuild_rows(self):
        self.row_widgets = []
        for child in self.scroll_frame.winfo_children():
            child.destroy()

        headers = ["Field", "Print name", "Show", "Font size", "Order"]
        for col, text in enumerate(headers):
            ctk.CTkLabel(self.scroll_frame, text=text).grid(
                row=0,
                column=col,
                padx=6,
                pady=(4, 8),
                sticky="w",
            )

        for idx, model in enumerate(self.row_models, start=1):
            key = model["key"]
            if is_empty_line_key(key):
                field_label = EMPTY_LINE_LABEL
                field_hint = ""
            else:
                field_label = LABEL_FIELD_LABELS.get(key, key)
                field_hint = f" ({key})"

            ctk.CTkLabel(self.scroll_frame, text=f"{field_label}{field_hint}").grid(
                row=idx,
                column=0,
                padx=6,
                pady=4,
                sticky="w",
            )
            print_name_entry = ctk.CTkEntry(self.scroll_frame)
            print_name_entry.grid(row=idx, column=1, padx=6, pady=4, sticky="ew")
            print_name_entry.insert(0, model["print_name"])

            show_checkbox = ctk.CTkCheckBox(
                self.scroll_frame,
                text="",
                width=24,
            )
            show_checkbox.grid(row=idx, column=2, padx=6, pady=4)
            if model["show"]:
                show_checkbox.select()
            else:
                show_checkbox.deselect()

            font_size_entry = ctk.CTkEntry(
                self.scroll_frame,
                width=70,
            )
            font_size_entry.grid(row=idx, column=3, padx=6, pady=4, sticky="w")
            font_size_entry.insert(0, str(model["font_size"]))

            order_buttons = ctk.CTkFrame(self.scroll_frame)
            order_buttons.grid(row=idx, column=4, padx=6, pady=4, sticky="w")
            ctk.CTkButton(
                order_buttons,
                text="Up",
                width=54,
                command=lambda i=idx - 1: self.move_up(i),
            ).grid(row=0, column=0, padx=(0, 4))
            ctk.CTkButton(
                order_buttons,
                text="Down",
                width=64,
                command=lambda i=idx - 1: self.move_down(i),
            ).grid(row=0, column=1)
            if is_empty_line_key(key):
                ctk.CTkButton(
                    order_buttons,
                    text="Delete",
                    width=70,
                    command=lambda i=idx - 1: self.delete_row(i),
                ).grid(row=0, column=2, padx=(4, 0))

            self.row_widgets.append(
                {
                    "print_name": print_name_entry,
                    "show": show_checkbox,
                    "font_size": font_size_entry,
                }
            )

    def _sync_models_from_widgets(self):
        if len(self.row_widgets) != len(self.row_models):
            return

        for model, widgets in zip(self.row_models, self.row_widgets):
            try:
                model["print_name"] = widgets["print_name"].get().strip()
            except tk.TclError:
                continue
            try:
                model["show"] = bool(widgets["show"].get())
            except tk.TclError:
                continue
            try:
                model["font_size"] = widgets["font_size"].get().strip()
            except tk.TclError:
                continue

    def serialize_config(self):
        self._sync_models_from_widgets()
        serialized = []
        for model in self.row_models:
            try:
                font_size = int(model["font_size"])
            except (TypeError, ValueError):
                font_size = 24

            serialized.append(
                {
                    "key": model["key"],
                    "print_name": str(model["print_name"]).strip(),
                    "show": bool(model["show"]),
                    "font_size": max(8, min(120, font_size)),
                }
            )
        return serialized

    def current_line_spacing(self):
        return normalize_label_line_spacing(self.line_spacing_var.get())

    def apply_to_app(self, save_default):
        self.app.set_label_field_config(
            self.serialize_config(),
            self.current_line_spacing(),
            save_default=save_default,
        )

    def move_up(self, index):
        if index <= 0:
            return
        self._sync_models_from_widgets()
        self.row_models[index - 1], self.row_models[index] = self.row_models[index], self.row_models[index - 1]
        self.rebuild_rows()
        self.apply_to_app(save_default=False)

    def move_down(self, index):
        if index >= len(self.row_models) - 1:
            return
        self._sync_models_from_widgets()
        self.row_models[index + 1], self.row_models[index] = self.row_models[index], self.row_models[index + 1]
        self.rebuild_rows()
        self.apply_to_app(save_default=False)

    def delete_row(self, index):
        if index < 0 or index >= len(self.row_models):
            return
        if not is_empty_line_key(self.row_models[index]["key"]):
            return
        self._sync_models_from_widgets()
        del self.row_models[index]
        self.rebuild_rows()
        self.apply_to_app(save_default=False)

    def on_add_empty_line(self):
        self._sync_models_from_widgets()
        self.row_models.append(make_empty_line_entry())
        self.rebuild_rows()
        self.apply_to_app(save_default=False)

    def on_apply(self):
        self.apply_to_app(save_default=True)

    def on_reset_default(self):
        self.set_config(default_label_field_config(), DEFAULT_LABEL_LINE_SPACING)
        self.apply_to_app(save_default=False)

    def on_save_json(self):
        file_path = filedialog.asksaveasfilename(
            parent=self,
            title="Save label config",
            defaultextension=".json",
            filetypes=[("JSON", "*.json"), ("All files", "*.*")],
        )
        if not file_path:
            return
        try:
            save_label_field_config(
                Path(file_path),
                self.serialize_config(),
                self.current_line_spacing(),
            )
        except Exception as e:
            messagebox.showerror("Save failed", f"Could not save label config:\n{e}", parent=self)

    def on_load_json(self):
        file_path = filedialog.askopenfilename(
            parent=self,
            title="Load label config",
            filetypes=[("JSON", "*.json"), ("All files", "*.*")],
        )
        if not file_path:
            return
        try:
            loaded_config, loaded_line_spacing = load_label_field_config(Path(file_path))
        except Exception as e:
            messagebox.showerror("Load failed", f"Could not load label config:\n{e}", parent=self)
            return

        self.set_config(loaded_config, loaded_line_spacing)
        self.apply_to_app(save_default=False)


def parse_decimal(value_text):
    text = str(value_text).strip().replace(",", ".")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def load_print_logs(path: Path):
    if not path.exists():
        return []

    entries = []
    with path.open("r", encoding="utf-8") as f:
        for raw_line in f:
            line = raw_line.strip()
            if not line:
                continue
            try:
                entry = json.loads(line)
            except json.JSONDecodeError:
                continue
            if isinstance(entry, dict):
                entries.append(entry)
    return entries


def append_print_log(path: Path, entry):
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False))
        f.write("\n")


def rotate_print_log(path: Path, archive_dir: Path):
    archive_dir.mkdir(parents=True, exist_ok=True)

    archived_path = None
    if path.exists():
        if not path.is_file():
            raise ValueError(f"Log path is not a file: {path}")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        candidate = archive_dir / f"{timestamp}.jsonl"
        suffix = 2
        while candidate.exists():
            candidate = archive_dir / f"{timestamp}_{suffix}.jsonl"
            suffix += 1

        path.replace(candidate)
        archived_path = candidate

    path.touch(exist_ok=True)
    return archived_path


class AnalyticsWindow(ctk.CTkToplevel):
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Analytics")
        self.geometry("980x600")

        main = ctk.CTkFrame(self)
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)
        main.columnconfigure(0, weight=1)

        top = ctk.CTkFrame(main)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ctk.CTkButton(top, text="Refresh", command=self.refresh_data).grid(row=0, column=0, padx=(0, 6))
        ctk.CTkButton(top, text="Reset", command=self.on_reset_log).grid(row=0, column=1)

        self.tabview = ctk.CTkTabview(main)
        self.tabview.grid(row=1, column=0, sticky="nsew")

        self._build_summary_tab()
        self._build_logs_tab()
        self._build_totals_tab()
        self.refresh_data()

    def _build_summary_tab(self):
        self.tabview.add("By Cut")
        tab = self.tabview.tab("By Cut")
        tab.rowconfigure(0, weight=1)
        tab.columnconfigure(0, weight=1)

        self.summary_tree = ttk.Treeview(
            tab,
            columns=("cut_name", "total_kg", "total_price"),
            show="headings",
        )
        self.summary_tree.heading("cut_name", text="Cut name")
        self.summary_tree.heading("total_kg", text="Total KG")
        self.summary_tree.heading("total_price", text="Total Price")
        self.summary_tree.column("cut_name", width=320, anchor="w")
        self.summary_tree.column("total_kg", width=150, anchor="e")
        self.summary_tree.column("total_price", width=180, anchor="e")
        self.summary_tree.grid(row=0, column=0, sticky="nsew")

        sc = ttk.Scrollbar(tab, orient="vertical", command=self.summary_tree.yview)
        sc.grid(row=0, column=1, sticky="ns")
        self.summary_tree.configure(yscrollcommand=sc.set)

    def _build_logs_tab(self):
        self.tabview.add("Log")
        tab = self.tabview.tab("Log")
        tab.rowconfigure(0, weight=1)
        tab.columnconfigure(0, weight=1)

        self.log_tree = ttk.Treeview(
            tab,
            columns=("time", "cut_name", "weight_kg", "price_per_kg", "total_price"),
            show="headings",
        )
        self.log_tree.heading("time", text="Time")
        self.log_tree.heading("cut_name", text="Cut name")
        self.log_tree.heading("weight_kg", text="Weight KG")
        self.log_tree.heading("price_per_kg", text="Price/KG")
        self.log_tree.heading("total_price", text="Total Price")
        self.log_tree.column("time", width=200, anchor="w")
        self.log_tree.column("cut_name", width=260, anchor="w")
        self.log_tree.column("weight_kg", width=120, anchor="e")
        self.log_tree.column("price_per_kg", width=120, anchor="e")
        self.log_tree.column("total_price", width=140, anchor="e")
        self.log_tree.grid(row=0, column=0, sticky="nsew")

        sc = ttk.Scrollbar(tab, orient="vertical", command=self.log_tree.yview)
        sc.grid(row=0, column=1, sticky="ns")
        self.log_tree.configure(yscrollcommand=sc.set)

    def _build_totals_tab(self):
        self.tabview.add("Totals")
        tab = self.tabview.tab("Totals")
        tab.columnconfigure(1, weight=1)

        self.total_weight_var = tk.StringVar(value="0.0000")
        self.total_price_var = tk.StringVar(value="0.00")
        self.total_count_var = tk.StringVar(value="0")

        ctk.CTkLabel(tab, text="Total prints").grid(row=0, column=0, sticky="w", pady=(0, 10))
        ctk.CTkLabel(tab, textvariable=self.total_count_var).grid(row=0, column=1, sticky="w", pady=(0, 10))
        ctk.CTkLabel(tab, text="Total weight (KG)").grid(row=1, column=0, sticky="w", pady=(0, 10))
        ctk.CTkLabel(tab, textvariable=self.total_weight_var).grid(row=1, column=1, sticky="w", pady=(0, 10))
        ctk.CTkLabel(tab, text="Total price").grid(row=2, column=0, sticky="w")
        ctk.CTkLabel(tab, textvariable=self.total_price_var).grid(row=2, column=1, sticky="w")

    def refresh_data(self):
        entries = load_print_logs(ANALYTICS_LOG_PATH)

        for child in self.summary_tree.get_children():
            self.summary_tree.delete(child)
        for child in self.log_tree.get_children():
            self.log_tree.delete(child)

        per_cut = {}
        total_weight = 0.0
        total_price = 0.0

        for entry in entries:
            cut_name = str(entry.get("cut_name", "")).strip() or "(empty)"
            weight = parse_decimal(entry.get("weight_kg"))
            price = parse_decimal(entry.get("total_price"))
            price_per_kg = parse_decimal(entry.get("price_per_kg"))
            time_text = str(entry.get("time", ""))

            if weight is None:
                weight = 0.0
            if price is None:
                price = 0.0

            total_weight += weight
            total_price += price

            if cut_name not in per_cut:
                per_cut[cut_name] = {"kg": 0.0, "price": 0.0}
            per_cut[cut_name]["kg"] += weight
            per_cut[cut_name]["price"] += price

            self.log_tree.insert(
                "",
                tk.END,
                values=(
                    time_text,
                    cut_name,
                    f"{weight:.4f}",
                    "" if price_per_kg is None else f"{price_per_kg:.2f}",
                    f"{price:.2f}",
                ),
            )

        for cut_name in sorted(per_cut):
            summary = per_cut[cut_name]
            self.summary_tree.insert(
                "",
                tk.END,
                values=(cut_name, f"{summary['kg']:.4f}", f"{summary['price']:.2f}"),
            )

        self.total_count_var.set(str(len(entries)))
        self.total_weight_var.set(f"{total_weight:.4f}")
        self.total_price_var.set(f"{total_price:.2f}")

    def on_reset_log(self):
        if not messagebox.askyesno(
            "Reset analytics log",
            "Archive current log and start a new empty log?",
            parent=self,
        ):
            return

        try:
            archived_path = rotate_print_log(ANALYTICS_LOG_PATH, LOG_ARCHIVE_DIR)
        except Exception as e:
            messagebox.showerror("Reset failed", f"Could not reset analytics log:\n{e}", parent=self)
            return

        self.refresh_data()
        if archived_path is None:
            messagebox.showinfo(
                "Analytics reset",
                "No existing log file was found. Created a new empty log.",
                parent=self,
            )
        else:
            messagebox.showinfo(
                "Analytics reset",
                f"Archived to:\n{archived_path}\n\nCreated new empty log:\n{ANALYTICS_LOG_PATH}",
                parent=self,
            )


class LabelApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("The Open Source Labeling Scale Project")
        self.geometry("1200x780")
        self.minsize(900, 620)

        self.items = []
        self.current_label_image = None
        self.preview_photo = None

        self.cut_name_var = tk.StringVar(value="")
        self.current_weight_var = tk.StringVar(value="")
        self.price_per_kg_var = tk.StringVar(value="")
        self.tax_var = tk.StringVar(value="")
        self.total_price_var = tk.StringVar(value="")
        self.scale_status_var = tk.StringVar(value="Scale: starting...")
        self.widget_scaling = DEFAULT_WIDGET_SCALING
        self.zoom_text_var = tk.StringVar(value="")
        self.auto_print_enabled_var = tk.BooleanVar(value=False)
        self.cut_paper_var = tk.BooleanVar(value=False)
        self.stable_iterations_var = tk.StringVar(value=str(SCALE_DEFAULT_STABLE_ITERATIONS))
        self.session_vars = {key: tk.StringVar(value="") for key, _label in SESSION_FIELDS}
        try:
            self.label_field_config, self.label_line_spacing = load_label_field_config(LABEL_CONFIG_PATH)
        except Exception as e:
            messagebox.showwarning(
                "Label config warning",
                f"Could not load {LABEL_CONFIG_PATH.name}; using defaults.\n\n{e}",
            )
            self.label_field_config = default_label_field_config()
            self.label_line_spacing = DEFAULT_LABEL_LINE_SPACING
        self.scale_stop_event = threading.Event()
        self.last_scale_value = None
        self.same_value_iterations = 0
        self.last_auto_printed_value = None
        self.auto_print_in_progress = False
        self.preview_update_job = None
        self.preview_resize_job = None
        self.session_save_job = None

        self.db_window = None
        self.session_window = None
        self.analytics_window = None
        self.label_config_window = None

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self._build_widgets()
        self.load_default_db()
        self.refresh_item_buttons()
        self.load_default_session()
        self.bind_preview_traces()
        self.update_total_price()
        self.update_preview()
        self.start_scale_monitor()

    def _build_widgets(self):
        self.rowconfigure(1, weight=1)
        self.columnconfigure(0, weight=1)

        top_bar = ctk.CTkFrame(self)
        top_bar.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        top_bar.columnconfigure(4, weight=1)

        ctk.CTkButton(top_bar, text="Edit DB", command=self.open_db_editor).grid(row=0, column=0, padx=(0, 6))
        ctk.CTkButton(top_bar, text="Session", command=self.open_session_window).grid(row=0, column=1, padx=(0, 6))
        ctk.CTkButton(top_bar, text="Analytics", command=self.open_analytics_window).grid(
            row=0, column=2, padx=(0, 6)
        )
        ctk.CTkButton(top_bar, text="Configure label", command=self.open_label_config_window).grid(
            row=0, column=3, padx=(0, 6)
        )
        ctk.CTkButton(top_bar, text="Exit", command=self.on_close).grid(row=0, column=5)
        self._update_zoom_text()

        self.main_split = tk.PanedWindow(
            self,
            orient=tk.VERTICAL,
            sashrelief=tk.RAISED,
            sashwidth=8,
            bd=0,
        )
        self.main_split.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 6))

        status_bar = ctk.CTkFrame(self, corner_radius=0)
        status_bar.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 10))
        status_bar.columnconfigure(0, weight=1)
        ctk.CTkLabel(status_bar, textvariable=self.scale_status_var, anchor="w").grid(
            row=0,
            column=0,
            sticky="w",
            padx=8,
            pady=4,
        )
        ctk.CTkLabel(status_bar, textvariable=self.zoom_text_var, anchor="e").grid(
            row=0,
            column=1,
            sticky="e",
            padx=(0, 6),
            pady=4,
        )
        ctk.CTkButton(status_bar, text="Zoom -", width=70, command=self.on_zoom_out).grid(
            row=0,
            column=2,
            padx=(0, 6),
            pady=2,
        )
        ctk.CTkButton(status_bar, text="Zoom +", width=70, command=self.on_zoom_in).grid(
            row=0,
            column=3,
            pady=2,
        )

        top_pane = ctk.CTkFrame(self.main_split)
        bottom_pane = ctk.CTkFrame(self.main_split)
        top_pane.rowconfigure(0, weight=1)
        top_pane.columnconfigure(0, weight=1)
        bottom_pane.rowconfigure(0, weight=1)
        bottom_pane.columnconfigure(0, weight=1)

        self.main_split.add(top_pane, minsize=140, stretch="always")
        self.main_split.add(bottom_pane, minsize=220, stretch="always")
        self.after(0, self._set_initial_splitter_position)

        self.buttons_frame = ctk.CTkFrame(top_pane)
        self.buttons_frame.grid(row=0, column=0, sticky="nsew")

        bottom = ctk.CTkFrame(bottom_pane)
        bottom.grid(row=0, column=0, sticky="nsew")
        bottom.rowconfigure(0, weight=1)
        bottom.columnconfigure(0, weight=1)
        bottom.columnconfigure(1, weight=1)

        form = ctk.CTkFrame(bottom)
        form.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        form.columnconfigure(1, weight=1)
        ctk.CTkLabel(form, text="Selected Item").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))

        ctk.CTkLabel(form, text="Cut name").grid(row=1, column=0, sticky="w", pady=(0, 8))
        ctk.CTkEntry(form, textvariable=self.cut_name_var).grid(row=1, column=1, sticky="ew", pady=(0, 8))

        ctk.CTkLabel(form, text="Price / KG").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ctk.CTkEntry(form, textvariable=self.price_per_kg_var).grid(row=2, column=1, sticky="ew", pady=(8, 0))

        ctk.CTkLabel(form, text="Tax").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ctk.CTkEntry(form, textvariable=self.tax_var).grid(row=3, column=1, sticky="ew", pady=(8, 0))

        ctk.CTkLabel(form, text="Current weight").grid(row=4, column=0, sticky="w", pady=(8, 0))
        ctk.CTkEntry(
            form,
            textvariable=self.current_weight_var,
            state="disabled",
        ).grid(row=4, column=1, sticky="ew", pady=(8, 0))

        ctk.CTkLabel(form, text="Total price").grid(row=5, column=0, sticky="w", pady=(8, 0))
        ctk.CTkEntry(
            form,
            textvariable=self.total_price_var,
            state="disabled",
        ).grid(row=5, column=1, sticky="ew", pady=(8, 0))

        preview_box = ctk.CTkFrame(bottom)
        preview_box.grid(row=0, column=1, sticky="nsew")
        preview_box.columnconfigure(0, weight=1)
        preview_box.rowconfigure(1, weight=1)
        ctk.CTkLabel(preview_box, text="Preview").grid(row=0, column=0, sticky="w")
        self.preview_viewport = ctk.CTkFrame(preview_box)
        self.preview_viewport.grid(row=1, column=0, sticky="nsew")
        self.preview_viewport.grid_propagate(False)
        self.preview_viewport.bind("<Configure>", lambda _e: self.request_preview_resize())
        self.preview_label = tk.Label(self.preview_viewport, bd=0, highlightthickness=0)
        self.preview_label.place(relx=0.5, rely=0.5, anchor="center")

        preview_controls = ctk.CTkFrame(preview_box)
        preview_controls.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        preview_controls.columnconfigure(5, weight=1)

        ctk.CTkButton(
            preview_controls,
            text="Print",
            command=self.on_print,
            width=220,
            height=86,
            font=ctk.CTkFont(size=28, weight="bold"),
        ).grid(row=0, column=0, rowspan=2, padx=(0, 10), pady=(0, 2), sticky="ns")
        ctk.CTkCheckBox(
            preview_controls,
            text="Print when stable",
            variable=self.auto_print_enabled_var,
        ).grid(row=0, column=1, padx=(0, 8))
        ctk.CTkCheckBox(
            preview_controls,
            text="Cut paper",
            variable=self.cut_paper_var,
        ).grid(row=1, column=1, padx=(0, 8), sticky="w")
        ctk.CTkLabel(preview_controls, text="Stable iterations").grid(row=0, column=2, padx=(0, 6))
        ttk.Spinbox(
            preview_controls,
            from_=1,
            to=1000,
            width=6,
            textvariable=self.stable_iterations_var,
        ).grid(row=0, column=3, padx=(0, 8))
    def _set_initial_splitter_position(self):
        try:
            total_height = self.main_split.winfo_height()
            if total_height > 1:
                self.main_split.sashpos(0, int(total_height * 0.55))
        except Exception:
            pass

    def _update_zoom_text(self):
        self.zoom_text_var.set(f"Zoom {int(round(self.widget_scaling * 100))}%")

    def set_widget_scaling_live(self, new_value):
        clamped = max(MIN_WIDGET_SCALING, min(MAX_WIDGET_SCALING, float(new_value)))
        clamped = round(clamped, 2)
        if abs(clamped - self.widget_scaling) < 1e-9:
            return

        self.widget_scaling = clamped
        ctk.set_widget_scaling(self.widget_scaling)
        self._update_zoom_text()
        self.request_preview_resize()

    def on_zoom_in(self):
        self.set_widget_scaling_live(self.widget_scaling + WIDGET_SCALING_STEP)

    def on_zoom_out(self):
        self.set_widget_scaling_live(self.widget_scaling - WIDGET_SCALING_STEP)

    def bind_preview_traces(self):
        self.cut_name_var.trace_add("write", lambda *_: self.request_preview_update())
        self.price_per_kg_var.trace_add("write", lambda *_: self.update_total_price())
        self.tax_var.trace_add("write", lambda *_: self.request_preview_update())
        self.current_weight_var.trace_add("write", lambda *_: self.update_total_price())
        self.current_weight_var.trace_add("write", lambda *_: self.request_preview_update())
        self.total_price_var.trace_add("write", lambda *_: self.request_preview_update())
        for var in self.session_vars.values():
            var.trace_add("write", lambda *_: self.request_preview_update())
            var.trace_add("write", lambda *_: self.request_session_autosave())

    def request_preview_update(self):
        if self.preview_update_job is not None:
            return
        self.preview_update_job = self.after(75, self._run_preview_update)

    def _run_preview_update(self):
        self.preview_update_job = None
        self.update_preview()

    def request_preview_resize(self):
        if self.preview_resize_job is not None:
            return
        self.preview_resize_job = self.after(60, self._run_preview_resize)

    def _run_preview_resize(self):
        self.preview_resize_job = None
        self.update_preview_display()

    def request_session_autosave(self):
        if self.session_save_job is not None:
            return
        self.session_save_job = self.after(400, self._run_session_autosave)

    def _run_session_autosave(self):
        self.session_save_job = None
        self.save_default_session(show_errors=False)

    def load_default_db(self):
        try:
            self.items = load_cut_items(CUT_DB_PATH)
        except Exception as e:
            messagebox.showerror("DB load error", f"Could not load {CUT_DB_PATH.name}:\n{e}")
            self.items = []

    def save_default_db(self):
        try:
            save_cut_items(CUT_DB_PATH, self.items)
        except Exception as e:
            messagebox.showerror("DB save error", f"Could not save {CUT_DB_PATH.name}:\n{e}")

    def load_default_session(self):
        if not SESSION_DEFAULT_PATH.exists():
            return
        try:
            self.load_session_from_file(SESSION_DEFAULT_PATH)
        except Exception as e:
            messagebox.showwarning(
                "Session load warning",
                f"Could not load {SESSION_DEFAULT_PATH.name}:\n{e}",
            )

    def save_default_session(self, show_errors):
        try:
            self.save_session_to_file(SESSION_DEFAULT_PATH)
        except Exception as e:
            if show_errors:
                messagebox.showwarning(
                    "Session save warning",
                    f"Could not save {SESSION_DEFAULT_PATH.name}:\n{e}",
                )

    def refresh_item_buttons(self):
        for child in self.buttons_frame.winfo_children():
            child.destroy()

        for col in range(5):
            self.buttons_frame.columnconfigure(col, weight=1)

        if not self.items:
            ctk.CTkLabel(
                self.buttons_frame,
                text="No items in DB. Click 'Edit DB' to add cuts.",
            ).grid(row=0, column=0, sticky="w")
            return

        for idx, item in enumerate(self.items):
            row = idx // 5
            col = idx % 5
            tax_text = str(item.get("tax", "")).strip()
            button_text = (
                f"{item.get('cut_name', '')}\n"
                f"{item.get('price_per_kg', '')} /KG"
            )
            if tax_text:
                button_text += f"\nTax {tax_text}"
            btn = ctk.CTkButton(
                self.buttons_frame,
                text=button_text,
                command=lambda i=idx: self.on_item_button(i),
            )
            btn.grid(row=row, column=col, sticky="nsew", padx=4, pady=4, ipadx=12, ipady=12)

    def on_item_button(self, index):
        item = self.items[index]
        self.cut_name_var.set(item.get("cut_name", ""))
        self.price_per_kg_var.set(item.get("price_per_kg", ""))
        self.tax_var.set(item.get("tax", ""))

    def start_scale_monitor(self):
        thread = threading.Thread(target=self._scale_monitor_worker, daemon=True)
        thread.start()

    def _scale_monitor_worker(self):
        while not self.scale_stop_event.is_set():
            try:
                ser = open_scale_serial()
            except Exception as e:
                self.after(0, lambda err=str(e): self._set_scale_offline_status(err))
                time.sleep(SCALE_RECONNECT_DELAY_SECONDS)
                continue

            self.after(0, lambda: self.scale_status_var.set("Scale: connected"))

            try:
                while not self.scale_stop_event.is_set():
                    raw = ser.readline()
                    if not raw:
                        continue

                    line = raw.decode(errors="ignore").strip()
                    if not line:
                        continue

                    match = SCALE_VALUE_RE.search(line)
                    if not match:
                        continue

                    value = float(match.group(1))
                    self.after(0, lambda v=value: self._on_live_scale_value(v))
            except Exception as e:
                self.after(0, lambda err=str(e): self._set_scale_offline_status(err))
            finally:
                try:
                    ser.close()
                except Exception:
                    pass

            time.sleep(SCALE_RECONNECT_DELAY_SECONDS)

    def _set_scale_offline_status(self, _error_text):
        self.scale_status_var.set("Scale: offline")
        self.last_scale_value = None
        self.same_value_iterations = 0
        self.last_auto_printed_value = None
        self.current_weight_var.set("n/a")

    def get_required_stable_iterations(self):
        try:
            value = int(self.stable_iterations_var.get())
        except (TypeError, ValueError):
            value = SCALE_DEFAULT_STABLE_ITERATIONS
        return max(1, value)

    def _trigger_auto_print(self):
        if self.auto_print_in_progress:
            return False

        weight = parse_decimal(self.current_weight_var.get())
        if weight is None or abs(weight) < MIN_PRINT_WEIGHT_KG:
            self.scale_status_var.set(f"Scale: waiting (>= {MIN_PRINT_WEIGHT_KG:.3f} kg)")
            return False

        self.auto_print_in_progress = True
        try:
            self.print_label()
            return True
        except Exception as e:
            self.scale_status_var.set("Scale: auto-print failed")
            messagebox.showerror("Auto print error", f"Could not auto print label:\n{e}")
            return False
        finally:
            self.auto_print_in_progress = False

    def _on_live_scale_value(self, value):
        self.scale_status_var.set("Scale: live")
        formatted = f"{value:.4f}"
        self.current_weight_var.set(formatted)

        if formatted == self.last_scale_value:
            self.same_value_iterations += 1
        else:
            self.last_scale_value = formatted
            self.same_value_iterations = 1
            self.last_auto_printed_value = None

        if not self.auto_print_enabled_var.get():
            return

        required = self.get_required_stable_iterations()
        if (
            self.same_value_iterations >= required
            and self.last_auto_printed_value != formatted
        ):
            if self._trigger_auto_print():
                self.last_auto_printed_value = formatted

    def update_total_price(self):
        weight = parse_decimal(self.current_weight_var.get())
        price_per_kg = parse_decimal(self.price_per_kg_var.get())

        if weight is None or price_per_kg is None:
            self.total_price_var.set("")
            return

        total = weight * price_per_kg
        self.total_price_var.set(f"{total:.2f}")

    def open_db_editor(self):
        if self.db_window and self.db_window.winfo_exists():
            self.db_window.lift()
            return
        self.db_window = DatabaseEditorWindow(self)

    def open_session_window(self):
        if self.session_window and self.session_window.winfo_exists():
            self.session_window.lift()
            return
        self.session_window = SessionWindow(self)

    def open_analytics_window(self):
        if self.analytics_window and self.analytics_window.winfo_exists():
            self.analytics_window.lift()
            self.analytics_window.refresh_data()
            return
        self.analytics_window = AnalyticsWindow(self)

    def set_label_field_config(self, config_items, line_spacing, save_default):
        self.label_field_config = normalize_label_field_config(config_items)
        self.label_line_spacing = normalize_label_line_spacing(line_spacing)
        if save_default:
            save_label_field_config(
                LABEL_CONFIG_PATH,
                self.label_field_config,
                self.label_line_spacing,
            )
        self.request_preview_update()

    def open_label_config_window(self):
        if self.label_config_window and self.label_config_window.winfo_exists():
            self.label_config_window.lift()
            return
        self.label_config_window = LabelConfigWindow(self)

    def current_session_data(self):
        return {key: var.get().strip() for key, var in self.session_vars.items()}

    def save_session_to_file(self, path: Path):
        with path.open("w", encoding="utf-8") as f:
            json.dump(self.current_session_data(), f, indent=2, ensure_ascii=False)

    def load_session_from_file(self, path: Path):
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)

        if not isinstance(data, dict):
            raise ValueError("Session file must contain a JSON object.")

        for key, _label in SESSION_FIELDS:
            self.session_vars[key].set(str(data.get(key, "")).strip())

    def update_preview(self):
        current_weight = parse_decimal(self.current_weight_var.get())
        weight_text = "" if current_weight is None else f"{current_weight:.4f}"
        total_price = parse_decimal(self.total_price_var.get())
        total_price_text = "" if total_price is None else f"{total_price:.2f}"
        price_per_kg = parse_decimal(self.price_per_kg_var.get())
        price_per_kg_text = "" if price_per_kg is None else f"{price_per_kg:.2f}"

        label_values = self.current_session_data()
        label_values.update(
            {
                "cut_name": self.cut_name_var.get().strip(),
                "weight_kg": weight_text,
                "price_per_kg": price_per_kg_text,
                "tax": self.tax_var.get().strip(),
                "total_price": total_price_text,
            }
        )

        self.current_label_image = build_label_image(
            label_values,
            self.label_field_config,
            self.label_line_spacing,
        )
        self.update_preview_display()

    def update_preview_display(self):
        if self.current_label_image is None:
            return

        max_w = max(1, self.preview_viewport.winfo_width())
        max_h = max(1, self.preview_viewport.winfo_height())

        # Wait until the viewport has a real layout size to avoid clipped first paint.
        if max_w <= 1 or max_h <= 1:
            self.request_preview_resize()
            return

        scale = min(max_w / LABEL_WIDTH, max_h / LABEL_HEIGHT)
        if scale <= 0:
            scale = 0.33
        scale = min(scale, 1.0)

        preview_size = (
            max(1, int(LABEL_WIDTH * scale)),
            max(1, int(LABEL_HEIGHT * scale)),
        )
        preview_img = self.current_label_image.resize(preview_size, RESAMPLE_LANCZOS)
        self.preview_photo = ImageTk.PhotoImage(preview_img)
        self.preview_label.configure(
            image=self.preview_photo,
            width=preview_size[0],
            height=preview_size[1],
        )

    def on_close(self):
        self.scale_stop_event.set()
        if self.preview_update_job is not None:
            try:
                self.after_cancel(self.preview_update_job)
            except Exception:
                pass
            self.preview_update_job = None
        if self.preview_resize_job is not None:
            try:
                self.after_cancel(self.preview_resize_job)
            except Exception:
                pass
            self.preview_resize_job = None
        if self.session_save_job is not None:
            try:
                self.after_cancel(self.session_save_job)
            except Exception:
                pass
            self.session_save_job = None
        self.save_default_session(show_errors=False)
        self.destroy()

    def build_print_log_entry(self):
        cut_name = self.cut_name_var.get().strip()
        weight = parse_decimal(self.current_weight_var.get())
        price_per_kg = parse_decimal(self.price_per_kg_var.get())
        total_price = parse_decimal(self.total_price_var.get())

        if total_price is None and weight is not None and price_per_kg is not None:
            total_price = weight * price_per_kg

        return {
            "time": datetime.now().isoformat(timespec="seconds"),
            "cut_name": cut_name,
            "weight_kg": None if weight is None else round(weight, 4),
            "price_per_kg": None if price_per_kg is None else round(price_per_kg, 4),
            "total_price": None if total_price is None else round(total_price, 2),
        }

    def log_successful_print(self):
        entry = self.build_print_log_entry()
        try:
            append_print_log(ANALYTICS_LOG_PATH, entry)
        except Exception as e:
            messagebox.showwarning("Analytics log warning", f"Printed, but log save failed:\n{e}")
            return

        if self.analytics_window and self.analytics_window.winfo_exists():
            self.analytics_window.refresh_data()

    def print_label(self):
        if self.current_label_image is None:
            self.update_preview()

        weight = parse_decimal(self.current_weight_var.get())
        if weight is None or abs(weight) < MIN_PRINT_WEIGHT_KG:
            raise ValueError(
                f"Current weight is too close to 0. Minimum is {MIN_PRINT_WEIGHT_KG:.3f} kg."
            )

        label_path = build_printed_label_path(
            self.cut_name_var.get().strip(),
            weight,
        )
        send_to_printer(
            self.current_label_image,
            label_path,
            cut_paper=bool(self.cut_paper_var.get()),
        )
        self.log_successful_print()

    def on_print(self):
        try:
            self.print_label()
        except Exception as e:
            err = str(e)
            if "resource busy" in err.lower():
                messagebox.showerror(
                    "Printer busy",
                    "Printer USB interface is busy.\n\n"
                    "Common fix:\n"
                    "  sudo systemctl stop ipp-usb\n\n"
                    f"Original error:\n{err}",
                )
                return
            messagebox.showerror("Print error", f"Could not print label:\n{err}")


if __name__ == "__main__":
    app = LabelApp()
    app.mainloop()
