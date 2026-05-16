import csv
import json
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

import qrcode
from django.conf import settings
from django.db.models import Count, Sum
from django.http import HttpResponse
from django.utils import timezone
from PIL import Image, ImageDraw, ImageFont

from .models import (
    AppSetting,
    CutPreset,
    LabelFieldConfig,
    PrintArchive,
    PrintLog,
    ScaleState,
    SessionSettings,
    WorkstationState,
)

WEB_ROOT = settings.BASE_DIR
LEGACY_ROOT = WEB_ROOT.parent
DATA_DIR = settings.DATA_DIR
PRINTED_LABELS_DIR = DATA_DIR / "printed_labels"
LOG_ARCHIVE_DIR = DATA_DIR / "log_archives"

PRINTER_MODEL = "QL-810W"
LABEL_SIZE = "62x100"
LABEL_WIDTH = 1109
LABEL_HEIGHT = 696
MIN_PRINT_WEIGHT_KG = Decimal("0.010")
PRINT_RETRY_ATTEMPTS = 4
PRINT_RETRY_BASE_DELAY_SECONDS = 0.35

SCALE_PORT_DEFAULT = "/dev/ttyUSB0"
SCALE_BAUDRATE_DEFAULT = "9600"
SCALE_RECONNECT_DELAY_SECONDS = 2.0
SCALE_VALUE_RE = re.compile(r"([-+]?\d+(?:\.\d+)?)\s*kg", re.IGNORECASE)
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
    ("identity_mark", "Identity Mark"),
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
    ("identity_mark", "Identity Mark"),
]
LABEL_FIELD_LABELS = {key: label for key, label in LABEL_FIELD_DEFS}
EMPTY_LINE_KEY_PREFIX = "__empty_line__"
EMPTY_LINE_LABEL = "Free Text"
DEFAULT_LABEL_LINE_SPACING = 8

try:
    RESAMPLE_NEAREST = Image.Resampling.NEAREST
    RESAMPLE_LANCZOS = Image.Resampling.LANCZOS
except AttributeError:
    RESAMPLE_NEAREST = Image.NEAREST
    RESAMPLE_LANCZOS = Image.LANCZOS

FONT_CACHE = {}
FONT_CANDIDATES = [
    str(WEB_ROOT / "fonts" / "OpenSans-Light.ttf"),
    str(LEGACY_ROOT / "OpenSans-Light.ttf"),
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/liberation/LiberationSans-Regular.ttf",
    "/usr/share/fonts/google-noto/NotoSans-Regular.ttf",
]


def ensure_data_dirs():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    PRINTED_LABELS_DIR.mkdir(parents=True, exist_ok=True)
    LOG_ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)
    settings.MEDIA_ROOT.mkdir(parents=True, exist_ok=True)
    (settings.MEDIA_ROOT / "logos").mkdir(parents=True, exist_ok=True)


def parse_decimal(value_text):
    text = str(value_text or "").strip().replace(",", ".")
    if not text:
        return None
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def format_decimal(value, places):
    if value is None:
        return ""
    return f"{Decimal(value):.{places}f}"


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
        elif entry["key"] == "identity_mark":
            entry["font_size"] = 28
            entry["show"] = False

    return defaults


def is_empty_line_key(key):
    return str(key).startswith(EMPTY_LINE_KEY_PREFIX)


def make_empty_line_entry():
    return {
        "key": f"{EMPTY_LINE_KEY_PREFIX}{time.time_ns()}",
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
                default_entry = {"key": key, "print_name": "", "show": True, "font_size": 24}
            else:
                continue

            try:
                font_size = int(raw.get("font_size", default_entry["font_size"]))
            except (TypeError, ValueError):
                font_size = int(default_entry["font_size"])
            normalized.append(
                {
                    "key": key,
                    "print_name": str(raw.get("print_name", default_entry["print_name"])).strip(),
                    "show": bool(raw.get("show", default_entry["show"])),
                    "font_size": max(8, min(120, font_size)),
                }
            )

    for key, _label in LABEL_FIELD_DEFS:
        if key not in seen_defaults:
            normalized.append(defaults_by_key[key])

    return normalized


def save_label_config_to_db(config_items, line_spacing):
    normalized = normalize_label_field_config(config_items)
    LabelFieldConfig.objects.all().delete()
    for index, entry in enumerate(normalized):
        LabelFieldConfig.objects.create(sort_order=index, **entry)
    AppSetting.set_value("label_line_spacing", normalize_label_line_spacing(line_spacing))


def load_label_config_from_db():
    if not LabelFieldConfig.objects.exists():
        save_label_config_to_db(default_label_field_config(), DEFAULT_LABEL_LINE_SPACING)
    rows = [
        {
            "id": row.id,
            "key": row.key,
            "print_name": row.print_name,
            "show": row.show,
            "font_size": row.font_size,
        }
        for row in LabelFieldConfig.objects.all()
    ]
    return rows, int(AppSetting.get_value("label_line_spacing", DEFAULT_LABEL_LINE_SPACING))


def load_font(size):
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


def resolve_logo_path(raw_path):
    logo_path = str(raw_path or "").strip()
    if not logo_path:
        return None
    candidate = Path(logo_path).expanduser()
    if candidate.is_absolute():
        return candidate
    for base in (WEB_ROOT, settings.MEDIA_ROOT, LEGACY_ROOT):
        resolved = base / candidate
        if resolved.exists():
            return resolved
    return WEB_ROOT / candidate


def _draw_identity_mark(draw, x, y, font_size, lines, spacing):
    if not lines:
        return 0

    font = load_font(font_size)
    max_text_w = 0
    for line in lines:
        bbox = font.getbbox(line)
        max_text_w = max(max_text_w, bbox[2] - bbox[0])

    oval_pad_h = max(16, font_size // 2)
    oval_pad_v = max(12, font_size // 3)
    inner_h = font_size * len(lines) + spacing * max(0, len(lines) - 1)
    oval_w = max_text_w + oval_pad_h * 2
    oval_h = inner_h + oval_pad_v * 2

    draw.ellipse([x, y, x + oval_w, y + oval_h], outline="black", width=3)
    text_y = y + oval_pad_v
    for line in lines:
        bbox = font.getbbox(line)
        text_w = bbox[2] - bbox[0]
        draw.text((x + (oval_w - text_w) // 2, text_y), line, fill="black", font=font)
        text_y += font_size + spacing
    return oval_h


def build_label_image(label_values, label_field_config, line_spacing):
    img = Image.new("RGB", (LABEL_WIDTH, LABEL_HEIGHT), "white")
    draw = ImageDraw.Draw(img)

    margin = 28
    logo_box_w = 250
    logo_box_h = 180
    config_by_key = {str(entry.get("key", "")).strip(): entry for entry in label_field_config}

    logo_entry = config_by_key.get("logo_path", {})
    logo_path = resolve_logo_path(label_values.get("logo_path", ""))
    if bool(logo_entry.get("show", False)) and logo_path and logo_path.exists():
        try:
            logo = Image.open(logo_path).convert("RGBA")
            logo.thumbnail((logo_box_w, logo_box_h), RESAMPLE_LANCZOS)
            img.paste(logo, (LABEL_WIDTH - margin - logo.width, margin), logo)
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
        if key == "identity_mark":
            oval_h = _draw_identity_mark(draw, margin, y, font_size, value.split()[:3], spacing)
            y += oval_h + spacing
            continue

        print_name = str(entry.get("print_name", LABEL_FIELD_LABELS.get(key, key))).strip()
        line = f"{print_name}: {value}" if print_name else value
        draw.text((margin, y), line, fill="black", font=load_font(font_size))
        y += font_size + spacing

    qr_payload = {k: v for k, v in label_values.items() if k != "logo_path"}
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=1,
    )
    qr.add_data(json.dumps(qr_payload, ensure_ascii=False))
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    qr_img = qr_img.resize((240, 240), RESAMPLE_NEAREST)
    img.paste(qr_img, (LABEL_WIDTH - margin - 240, LABEL_HEIGHT - margin - 240))
    return img


def selected_label_values():
    session = SessionSettings.load()
    workstation = WorkstationState.load()
    scale = ScaleState.load()
    weight = scale.current_weight_kg
    price_per_kg = workstation.price_per_kg
    total_price = None
    if weight is not None and price_per_kg is not None:
        total_price = Decimal(weight) * Decimal(price_per_kg)

    values = session.as_label_values()
    values.update(
        {
            "cut_name": workstation.cut_name,
            "weight_kg": format_decimal(weight, 4),
            "price_per_kg": format_decimal(price_per_kg, 2),
            "tax": workstation.tax,
            "total_price": format_decimal(total_price, 2),
        }
    )
    return values


def build_current_label_image():
    config, line_spacing = load_label_config_from_db()
    return build_label_image(selected_label_values(), config, line_spacing)


def sanitize_filename_part(value, fallback):
    text = str(value or "").strip()
    if not text:
        return fallback
    text = FILENAME_PART_RE.sub("_", text)
    text = re.sub(r"_+", "_", text).strip("._")
    return text or fallback


def build_printed_label_path(cut_name, weight):
    ensure_data_dirs()
    timestamp = timezone.localtime().strftime("%Y%m%d_%H%M%S")
    cut_part = sanitize_filename_part(cut_name, "cut")
    weight_part = sanitize_filename_part(f"{Decimal(weight):.4f}kg", "unknownkg")
    base_name = f"{timestamp}_{cut_part}_{weight_part}"
    target = PRINTED_LABELS_DIR / f"{base_name}.png"
    suffix = 2
    while target.exists():
        target = PRINTED_LABELS_DIR / f"{base_name}_{suffix}.png"
        suffix += 1
    return target


def is_resource_busy_error(error_text):
    text = str(error_text).lower()
    return "resource busy" in text or "errno 16" in text or "usb.core.usberror" in text


def printer_identifier():
    return AppSetting.get_value("printer_identifier", "usb://0x04f9:0x209c")


def print_via_brother_cli(pil_image, image_path, cut_paper):
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
        printer_identifier(),
        "print",
        "-l",
        LABEL_SIZE,
    ]
    if not cut_paper:
        cmd.append("--no-cut")
    cmd.append(str(image_path))

    last_output = ""
    for attempt in range(1, PRINT_RETRY_ATTEMPTS + 1):
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, check=False)
        if result.returncode == 0:
            return
        output = (result.stderr or "").strip() or (result.stdout or "").strip()
        last_output = output or "brother_ql CLI print failed."
        if attempt < PRINT_RETRY_ATTEMPTS and is_resource_busy_error(last_output):
            time.sleep(PRINT_RETRY_BASE_DELAY_SECONDS * attempt)
            continue
        raise RuntimeError(last_output)
    raise RuntimeError(last_output or "brother_ql CLI print failed.")


def print_current_label(trigger="manual"):
    workstation = WorkstationState.load()
    scale = ScaleState.load()
    weight = scale.current_weight_kg
    if weight is None or abs(Decimal(weight)) < MIN_PRINT_WEIGHT_KG:
        raise ValueError(f"Current weight is too close to 0. Minimum is {MIN_PRINT_WEIGHT_KG:.3f} kg.")

    label_values = selected_label_values()
    image = build_current_label_image()
    label_path = build_printed_label_path(workstation.cut_name, weight)
    print_via_brother_cli(image, label_path, cut_paper=workstation.cut_paper)

    price_per_kg = workstation.price_per_kg
    total_price = None
    if price_per_kg is not None:
        total_price = Decimal(weight) * Decimal(price_per_kg)

    log = PrintLog.objects.create(
        cut_name=workstation.cut_name,
        weight_kg=weight,
        price_per_kg=price_per_kg,
        tax=workstation.tax,
        total_price=total_price,
        label_image_path=str(label_path.relative_to(WEB_ROOT)),
        label_values=label_values,
        trigger=trigger,
    )
    return log


def analytics_queryset():
    return PrintLog.objects.filter(archive__isnull=True)


def analytics_context():
    logs = analytics_queryset()
    totals = logs.aggregate(
        total_weight=Sum("weight_kg"),
        total_price=Sum("total_price"),
        total_count=Count("id"),
    )
    by_cut = (
        logs.values("cut_name")
        .annotate(total_kg=Sum("weight_kg"), total_price=Sum("total_price"), count=Count("id"))
        .order_by("cut_name")
    )
    return {
        "logs": logs[:300],
        "by_cut": by_cut,
        "total_weight": totals["total_weight"] or Decimal("0"),
        "total_price": totals["total_price"] or Decimal("0"),
        "total_count": totals["total_count"] or 0,
    }


def export_analytics_csv():
    response = HttpResponse(content_type="text/csv")
    response["Content-Disposition"] = 'attachment; filename="osls_analytics.csv"'
    writer = csv.writer(response)
    writer.writerow(["Cut name", "Total KG", "Total Price", "Labels"])
    for row in analytics_context()["by_cut"]:
        writer.writerow([row["cut_name"], row["total_kg"] or 0, row["total_price"] or 0, row["count"]])
    return response


def archive_active_logs():
    archive = PrintArchive.objects.create(note="Manual analytics reset")
    PrintLog.objects.filter(archive__isnull=True).update(archive=archive)
    return archive


def seed_from_legacy_files(force=False):
    ensure_data_dirs()
    legacy_config = LEGACY_ROOT / "config"

    cuts_path = legacy_config / "cuts_db.json"
    if force or not CutPreset.objects.exists():
        CutPreset.objects.all().delete()
        if cuts_path.exists():
            data = json.loads(cuts_path.read_text(encoding="utf-8"))
            for index, item in enumerate(data):
                CutPreset.objects.create(
                    cut_name=str(item.get("cut_name", "")).strip(),
                    price_per_kg=parse_decimal(item.get("price_per_kg")) or Decimal("0.00"),
                    tax=str(item.get("tax", "")).strip(),
                    sort_order=index,
                )

    session_path = legacy_config / "session_default.json"
    if force or not SessionSettings.objects.exists():
        session = SessionSettings.load()
        if session_path.exists():
            data = json.loads(session_path.read_text(encoding="utf-8"))
            for field, _label in SESSION_FIELDS:
                value = str(data.get(field, "")).strip()
                if field == "logo_path" and value == "logo.png":
                    value = "logos/logo.png"
                setattr(session, field, value)
            session.save()

    legacy_logo = LEGACY_ROOT / "logo.png"
    target_logo = settings.MEDIA_ROOT / "logos" / "logo.png"
    if legacy_logo.exists() and (force or not target_logo.exists()):
        shutil.copy2(legacy_logo, target_logo)

    label_path = legacy_config / "label_config.json"
    if force or not LabelFieldConfig.objects.exists():
        if label_path.exists():
            data = json.loads(label_path.read_text(encoding="utf-8"))
            fields = data.get("fields", [])
            line_spacing = data.get("line_spacing", DEFAULT_LABEL_LINE_SPACING)
            save_label_config_to_db(fields, line_spacing)
        else:
            save_label_config_to_db(default_label_field_config(), DEFAULT_LABEL_LINE_SPACING)

    printer_path = legacy_config / "printer_config.json"
    if printer_path.exists():
        data = json.loads(printer_path.read_text(encoding="utf-8"))
        usb = data.get("usb", {})
        AppSetting.set_value("printer_identifier", usb.get("identifier", "usb://0x04f9:0x209c"))
    AppSetting.set_value("scale_port", AppSetting.get_value("scale_port", SCALE_PORT_DEFAULT))
    AppSetting.set_value("scale_baudrate", AppSetting.get_value("scale_baudrate", SCALE_BAUDRATE_DEFAULT))

    legacy_log = LEGACY_ROOT / "print_log.jsonl"
    if (force or not PrintLog.objects.exists()) and legacy_log.exists():
        if force:
            PrintLog.objects.all().delete()
        for raw_line in legacy_log.read_text(encoding="utf-8").splitlines():
            try:
                entry = json.loads(raw_line)
            except json.JSONDecodeError:
                continue
            printed_at = timezone.now()
            time_text = entry.get("time")
            if time_text:
                try:
                    printed_at = datetime.fromisoformat(time_text)
                    if timezone.is_naive(printed_at):
                        printed_at = timezone.make_aware(printed_at)
                except ValueError:
                    printed_at = timezone.now()
            PrintLog.objects.create(
                printed_at=printed_at,
                cut_name=str(entry.get("cut_name", "")).strip(),
                weight_kg=parse_decimal(entry.get("weight_kg")),
                price_per_kg=parse_decimal(entry.get("price_per_kg")),
                total_price=parse_decimal(entry.get("total_price")),
                label_values=entry,
                trigger="import",
            )
