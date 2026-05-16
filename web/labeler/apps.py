import os
import sys

from django.apps import AppConfig


class LabelerConfig(AppConfig):
    default_auto_field = "django.db.models.BigAutoField"
    name = "labeler"

    def ready(self):
        disabled = os.environ.get("OSLS_DISABLE_SCALE_MONITOR") == "1"
        management_only = {
            "makemigrations",
            "migrate",
            "collectstatic",
            "shell",
            "createsuperuser",
            "import_legacy_data",
            "check",
            "test",
        }
        if disabled or any(command in sys.argv for command in management_only):
            return

        if "runserver" in sys.argv and os.environ.get("RUN_MAIN") != "true":
            return

        from .scale_service import start_scale_monitor

        start_scale_monitor()
