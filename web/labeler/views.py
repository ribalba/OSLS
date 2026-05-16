from decimal import Decimal
from io import BytesIO

from django.conf import settings
from django.contrib import messages
from django.core.files.storage import FileSystemStorage
from django.db.models import Max
from django.http import HttpResponse
from django.shortcuts import redirect, render
from django.views.decorators.http import require_GET, require_http_methods

from .forms import SessionSettingsForm, WorkstationForm
from .models import AppSetting, CutPreset, LabelFieldConfig, ScaleState, SessionSettings, WorkstationState
from .services import (
    DEFAULT_LABEL_LINE_SPACING,
    EMPTY_LINE_LABEL,
    LABEL_FIELD_LABELS,
    analytics_context,
    archive_active_logs,
    build_current_label_image,
    export_analytics_csv,
    format_decimal,
    is_empty_line_key,
    load_label_config_from_db,
    make_empty_line_entry,
    normalize_label_field_config,
    normalize_label_line_spacing,
    parse_decimal,
    print_current_label,
    save_label_config_to_db,
    seed_from_legacy_files,
)


def ensure_initial_data():
    if not CutPreset.objects.exists() or not LabelFieldConfig.objects.exists():
        seed_from_legacy_files(force=False)


def resequence_cuts():
    for index, cut in enumerate(CutPreset.objects.order_by("sort_order", "id")):
        if cut.sort_order != index:
            cut.sort_order = index
            cut.save(update_fields=["sort_order"])


def index(request):
    ensure_initial_data()
    workstation = WorkstationState.load()
    scale = ScaleState.load()

    if request.method == "POST":
        action = request.POST.get("action", "")
        if action == "select_cut":
            cut = CutPreset.objects.filter(pk=request.POST.get("cut_id"), active=True).first()
            if cut:
                workstation.selected_cut = cut
                workstation.cut_name = cut.cut_name
                workstation.price_per_kg = cut.price_per_kg
                workstation.tax = cut.tax
                workstation.save()
                messages.success(request, f"Selected {cut.cut_name}.")
            return redirect("index")

        form = WorkstationForm(request.POST, instance=workstation)
        if form.is_valid():
            workstation = form.save(commit=False)
            workstation.selected_cut = None
            workstation.save()
        else:
            messages.error(request, "Selected item settings are invalid.")
            return redirect("index")

        if action == "print":
            try:
                log = print_current_label(trigger="manual")
            except Exception as exc:
                text = str(exc)
                if "resource busy" in text.lower():
                    text = f"{text} Common fix on Linux: sudo systemctl stop ipp-usb"
                messages.error(request, f"Print failed: {text}")
            else:
                messages.success(request, f"Printed label #{log.id}.")
        else:
            messages.success(request, "Selection updated.")
        return redirect("index")

    weight = scale.current_weight_kg
    total_price = None
    if weight is not None and workstation.price_per_kg is not None:
        total_price = Decimal(weight) * Decimal(workstation.price_per_kg)

    return render(
        request,
        "labeler/index.html",
        {
            "cuts": CutPreset.objects.filter(active=True),
            "workstation": workstation,
            "scale": scale,
            "current_weight": format_decimal(weight, 4) or "n/a",
            "total_price": format_decimal(total_price, 2),
        },
    )


@require_GET
def preview_png(request):
    ensure_initial_data()
    image = build_current_label_image()
    buffer = BytesIO()
    image.save(buffer, format="PNG")
    return HttpResponse(buffer.getvalue(), content_type="image/png")


@require_http_methods(["GET", "POST"])
def cuts(request):
    ensure_initial_data()
    if request.method == "POST":
        action = request.POST.get("action", "")
        cut = CutPreset.objects.filter(pk=request.POST.get("cut_id")).first()

        if action == "add":
            name = request.POST.get("cut_name", "").strip()
            price = parse_decimal(request.POST.get("price_per_kg")) or Decimal("0.00")
            if name:
                next_order = (CutPreset.objects.aggregate(Max("sort_order"))["sort_order__max"] or 0) + 1
                CutPreset.objects.create(
                    cut_name=name,
                    price_per_kg=price,
                    tax=request.POST.get("tax", "").strip(),
                    sort_order=next_order,
                    active=True,
                )
                messages.success(request, "Cut added.")
            else:
                messages.error(request, "Cut name cannot be empty.")
        elif cut and action == "update":
            cut.cut_name = request.POST.get("cut_name", "").strip()
            cut.price_per_kg = parse_decimal(request.POST.get("price_per_kg")) or Decimal("0.00")
            cut.tax = request.POST.get("tax", "").strip()
            cut.active = request.POST.get("active") == "on"
            cut.save()
            messages.success(request, "Cut updated.")
        elif cut and action == "delete":
            cut.delete()
            resequence_cuts()
            messages.success(request, "Cut deleted.")
        elif cut and action in {"move_up", "move_down"}:
            ordered = list(CutPreset.objects.order_by("sort_order", "id"))
            index_pos = next((idx for idx, item in enumerate(ordered) if item.id == cut.id), None)
            if index_pos is not None:
                target_pos = index_pos - 1 if action == "move_up" else index_pos + 1
                if 0 <= target_pos < len(ordered):
                    other = ordered[target_pos]
                    cut.sort_order, other.sort_order = other.sort_order, cut.sort_order
                    cut.save(update_fields=["sort_order"])
                    other.save(update_fields=["sort_order"])
        return redirect("cuts")

    return render(request, "labeler/cuts.html", {"cuts": CutPreset.objects.all()})


@require_http_methods(["GET", "POST"])
def session_settings(request):
    ensure_initial_data()
    session = SessionSettings.load()
    if request.method == "POST":
        form = SessionSettingsForm(request.POST, instance=session)
        if form.is_valid():
            session = form.save(commit=False)
            upload = request.FILES.get("logo_upload")
            if upload:
                storage = FileSystemStorage(location=settings.MEDIA_ROOT / "logos")
                filename = storage.save(upload.name, upload)
                session.logo_path = f"logos/{filename}"
            session.save()
            messages.success(request, "Session saved.")
            return redirect("session_settings")
        messages.error(request, "Session settings are invalid.")
    else:
        form = SessionSettingsForm(instance=session)
    return render(request, "labeler/session.html", {"form": form, "session": session})


def _label_rows_from_post(request):
    rows = []
    for row_id in request.POST.getlist("row_id"):
        rows.append(
            {
                "id": int(row_id),
                "key": request.POST.get(f"key_{row_id}", ""),
                "print_name": request.POST.get(f"print_name_{row_id}", ""),
                "show": request.POST.get(f"show_{row_id}") == "on",
                "font_size": request.POST.get(f"font_size_{row_id}", "24"),
            }
        )
    return rows


def _save_label_rows_from_post(request):
    rows = _label_rows_from_post(request)
    spacing = normalize_label_line_spacing(request.POST.get("line_spacing", DEFAULT_LABEL_LINE_SPACING))
    save_label_config_to_db(rows, spacing)


@require_http_methods(["GET", "POST"])
def label_config(request):
    ensure_initial_data()
    if request.method == "POST":
        action = request.POST.get("action", "save")
        if action == "reset":
            save_label_config_to_db(normalize_label_field_config([]), DEFAULT_LABEL_LINE_SPACING)
            messages.success(request, "Label config reset.")
            return redirect("label_config")

        _save_label_rows_from_post(request)
        if action == "add_free_text":
            rows, spacing = load_label_config_from_db()
            rows.append(make_empty_line_entry())
            save_label_config_to_db(rows, spacing)
            messages.success(request, "Free text row added.")
        elif action in {"move_up", "move_down"}:
            row = LabelFieldConfig.objects.filter(pk=request.POST.get("target_id")).first()
            if row:
                ordered = list(LabelFieldConfig.objects.all())
                index_pos = next((idx for idx, item in enumerate(ordered) if item.id == row.id), None)
                target_pos = index_pos - 1 if action == "move_up" else index_pos + 1
                if index_pos is not None and 0 <= target_pos < len(ordered):
                    other = ordered[target_pos]
                    row.sort_order, other.sort_order = other.sort_order, row.sort_order
                    row.save(update_fields=["sort_order"])
                    other.save(update_fields=["sort_order"])
        elif action == "delete":
            row = LabelFieldConfig.objects.filter(pk=request.POST.get("target_id")).first()
            if row and is_empty_line_key(row.key):
                row.delete()
                for index, item in enumerate(LabelFieldConfig.objects.all()):
                    item.sort_order = index
                    item.save(update_fields=["sort_order"])
                messages.success(request, "Free text row deleted.")
        else:
            messages.success(request, "Label config saved.")
        return redirect("label_config")

    rows, spacing = load_label_config_from_db()
    for row in rows:
        row["field_label"] = EMPTY_LINE_LABEL if is_empty_line_key(row["key"]) else LABEL_FIELD_LABELS.get(row["key"], row["key"])
        row["is_free_text"] = is_empty_line_key(row["key"])
    return render(request, "labeler/label_config.html", {"rows": rows, "line_spacing": spacing})


@require_http_methods(["GET", "POST"])
def analytics(request):
    ensure_initial_data()
    if request.method == "POST":
        action = request.POST.get("action", "")
        if action == "export":
            return export_analytics_csv()
        if action == "reset":
            archive = archive_active_logs()
            messages.success(request, f"Analytics reset. Archived old logs as archive #{archive.id}.")
            return redirect("analytics")
    return render(request, "labeler/analytics.html", analytics_context())


@require_http_methods(["GET", "POST"])
def hardware_settings(request):
    ensure_initial_data()
    keys = {
        "printer_identifier": "usb://0x04f9:0x209c",
        "scale_port": "/dev/ttyUSB0",
        "scale_baudrate": "9600",
    }
    if request.method == "POST":
        for key in keys:
            AppSetting.set_value(key, request.POST.get(key, keys[key]).strip())
        messages.success(request, "Hardware settings saved.")
        return redirect("hardware_settings")

    values = {key: AppSetting.get_value(key, default) for key, default in keys.items()}
    return render(request, "labeler/hardware_settings.html", {"settings_values": values})
