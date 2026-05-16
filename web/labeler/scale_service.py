import threading
import time
from decimal import Decimal, InvalidOperation

from django.db import close_old_connections
from django.utils import timezone

from .models import AppSetting, ScaleState, WorkstationState
from .services import (
    MIN_PRINT_WEIGHT_KG,
    SCALE_BAUDRATE_DEFAULT,
    SCALE_PORT_DEFAULT,
    SCALE_RECONNECT_DELAY_SECONDS,
    SCALE_VALUE_RE,
    print_current_label,
)

try:
    import serial
except Exception:
    serial = None

_thread = None
_thread_lock = threading.Lock()
_print_lock = threading.Lock()


def start_scale_monitor():
    global _thread
    with _thread_lock:
        if _thread and _thread.is_alive():
            return
        _thread = threading.Thread(target=_scale_monitor_worker, name="osls-scale-monitor", daemon=True)
        _thread.start()


def _open_scale_serial():
    if serial is None:
        raise RuntimeError("pyserial is not installed. Install dependencies from web/requirements.txt.")

    port = AppSetting.get_value("scale_port", SCALE_PORT_DEFAULT)
    try:
        baudrate = int(AppSetting.get_value("scale_baudrate", SCALE_BAUDRATE_DEFAULT))
    except ValueError:
        baudrate = int(SCALE_BAUDRATE_DEFAULT)

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


def _set_offline_status(error_text):
    ScaleState.objects.update_or_create(
        pk=1,
        defaults={
            "connected": False,
            "status": "Scale: offline",
            "current_weight_kg": None,
            "stable_weight_kg": None,
            "same_value_iterations": 0,
            "last_error": str(error_text),
            "updated_at": timezone.now(),
        },
    )
    workstation = WorkstationState.load()
    workstation.last_auto_printed_weight = None
    workstation.must_zero_before_next_print = False
    workstation.auto_print_in_progress = False
    workstation.save(
        update_fields=[
            "last_auto_printed_weight",
            "must_zero_before_next_print",
            "auto_print_in_progress",
            "updated_at",
        ]
    )


def _handle_weight(value):
    try:
        weight = Decimal(str(value)).quantize(Decimal("0.0001"))
    except (InvalidOperation, ValueError):
        return

    scale = ScaleState.load()
    previous = scale.current_weight_kg
    if previous == weight:
        same_count = scale.same_value_iterations + 1
    else:
        same_count = 1

    stable_iterations = WorkstationState.load().stable_iterations
    stable_weight = weight if same_count >= stable_iterations else scale.stable_weight_kg
    scale.connected = True
    scale.status = "Scale: live"
    scale.current_weight_kg = weight
    scale.stable_weight_kg = stable_weight
    scale.same_value_iterations = same_count
    scale.last_error = ""
    scale.updated_at = timezone.now()
    scale.save()

    workstation = WorkstationState.load()
    if abs(weight) < MIN_PRINT_WEIGHT_KG:
        workstation.must_zero_before_next_print = False
        workstation.last_auto_printed_weight = None
        workstation.status_message = ""
        workstation.save(
            update_fields=[
                "must_zero_before_next_print",
                "last_auto_printed_weight",
                "status_message",
                "updated_at",
            ]
        )
        return

    if not workstation.auto_print_enabled:
        return
    if workstation.must_zero_before_next_print:
        workstation.status_message = "Return scale to zero before next auto-print."
        workstation.save(update_fields=["status_message", "updated_at"])
        return
    if same_count < max(1, workstation.stable_iterations):
        return
    if workstation.last_auto_printed_weight == weight:
        return

    with _print_lock:
        workstation = WorkstationState.load()
        if workstation.auto_print_in_progress:
            return
        workstation.auto_print_in_progress = True
        workstation.status_message = "Auto-printing..."
        workstation.save(update_fields=["auto_print_in_progress", "status_message", "updated_at"])
        try:
            print_current_label(trigger="auto")
        except Exception as exc:
            workstation.status_message = f"Auto-print failed: {exc}"
        else:
            workstation.last_auto_printed_weight = weight
            workstation.must_zero_before_next_print = True
            workstation.status_message = "Auto-print done."
        finally:
            workstation.auto_print_in_progress = False
            workstation.save()


def _scale_monitor_worker():
    while True:
        close_old_connections()
        try:
            ser = _open_scale_serial()
        except Exception as exc:
            _set_offline_status(exc)
            time.sleep(SCALE_RECONNECT_DELAY_SECONDS)
            continue

        ScaleState.objects.update_or_create(
            pk=1,
            defaults={
                "connected": True,
                "status": "Scale: connected",
                "last_error": "",
                "updated_at": timezone.now(),
            },
        )

        try:
            while True:
                raw = ser.readline()
                if not raw:
                    close_old_connections()
                    continue
                line = raw.decode(errors="ignore").strip()
                match = SCALE_VALUE_RE.search(line)
                if not match:
                    continue
                _handle_weight(match.group(1))
        except Exception as exc:
            _set_offline_status(exc)
        finally:
            try:
                ser.close()
            except Exception:
                pass
        time.sleep(SCALE_RECONNECT_DELAY_SECONDS)
