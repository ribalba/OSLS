from decimal import Decimal

from django.db import models
from django.utils import timezone


class AppSetting(models.Model):
    key = models.CharField(max_length=80, unique=True)
    value = models.TextField(blank=True, default="")

    def __str__(self):
        return self.key

    @classmethod
    def get_value(cls, key, default=""):
        setting = cls.objects.filter(key=key).first()
        if setting is None:
            return default
        return setting.value

    @classmethod
    def set_value(cls, key, value):
        cls.objects.update_or_create(key=key, defaults={"value": str(value)})


class CutPreset(models.Model):
    cut_name = models.CharField(max_length=200)
    price_per_kg = models.DecimalField(max_digits=10, decimal_places=2, default=Decimal("0.00"))
    tax = models.CharField(max_length=80, blank=True, default="")
    sort_order = models.PositiveIntegerField(default=0)
    active = models.BooleanField(default=True)

    class Meta:
        ordering = ["sort_order", "cut_name", "id"]

    def __str__(self):
        return self.cut_name


class SessionSettings(models.Model):
    farm_name = models.CharField(max_length=200, blank=True, default="")
    logo_path = models.CharField(max_length=500, blank=True, default="")
    animal_number = models.CharField(max_length=200, blank=True, default="")
    farm_number = models.CharField(max_length=200, blank=True, default="")
    due_date_4_7 = models.CharField(max_length=200, blank=True, default="")
    due_date_frozen = models.CharField(max_length=200, blank=True, default="")
    birth_country = models.CharField(max_length=200, blank=True, default="")
    life_country = models.CharField(max_length=200, blank=True, default="")
    slaughter_country = models.CharField(max_length=200, blank=True, default="")
    packaged_country = models.CharField(max_length=200, blank=True, default="")
    identity_mark = models.CharField(max_length=200, blank=True, default="")

    @classmethod
    def load(cls):
        obj, _created = cls.objects.get_or_create(pk=1)
        return obj

    def as_label_values(self):
        return {
            "farm_name": self.farm_name,
            "logo_path": self.logo_path,
            "animal_number": self.animal_number,
            "farm_number": self.farm_number,
            "due_date_4_7": self.due_date_4_7,
            "due_date_frozen": self.due_date_frozen,
            "birth_country": self.birth_country,
            "life_country": self.life_country,
            "slaughter_country": self.slaughter_country,
            "packaged_country": self.packaged_country,
            "identity_mark": self.identity_mark,
        }


class LabelFieldConfig(models.Model):
    key = models.CharField(max_length=120)
    print_name = models.CharField(max_length=300, blank=True, default="")
    show = models.BooleanField(default=True)
    font_size = models.PositiveIntegerField(default=24)
    sort_order = models.PositiveIntegerField(default=0)

    class Meta:
        ordering = ["sort_order", "id"]

    def __str__(self):
        return self.key


class ScaleState(models.Model):
    connected = models.BooleanField(default=False)
    status = models.CharField(max_length=200, default="Scale: starting...")
    current_weight_kg = models.DecimalField(max_digits=10, decimal_places=4, null=True, blank=True)
    stable_weight_kg = models.DecimalField(max_digits=10, decimal_places=4, null=True, blank=True)
    same_value_iterations = models.PositiveIntegerField(default=0)
    last_error = models.TextField(blank=True, default="")
    updated_at = models.DateTimeField(default=timezone.now)

    @classmethod
    def load(cls):
        obj, _created = cls.objects.get_or_create(pk=1)
        return obj


class WorkstationState(models.Model):
    selected_cut = models.ForeignKey(CutPreset, null=True, blank=True, on_delete=models.SET_NULL)
    cut_name = models.CharField(max_length=200, blank=True, default="")
    price_per_kg = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True)
    tax = models.CharField(max_length=80, blank=True, default="")
    auto_print_enabled = models.BooleanField(default=False)
    cut_paper = models.BooleanField(default=False)
    stable_iterations = models.PositiveIntegerField(default=10)
    last_auto_printed_weight = models.DecimalField(max_digits=10, decimal_places=4, null=True, blank=True)
    must_zero_before_next_print = models.BooleanField(default=False)
    auto_print_in_progress = models.BooleanField(default=False)
    status_message = models.CharField(max_length=300, blank=True, default="")
    updated_at = models.DateTimeField(auto_now=True)

    @classmethod
    def load(cls):
        obj, _created = cls.objects.get_or_create(pk=1)
        return obj


class PrintArchive(models.Model):
    created_at = models.DateTimeField(default=timezone.now)
    note = models.CharField(max_length=200, blank=True, default="")

    class Meta:
        ordering = ["-created_at", "-id"]

    def __str__(self):
        return self.created_at.isoformat(timespec="seconds")


class PrintLog(models.Model):
    printed_at = models.DateTimeField(default=timezone.now)
    cut_name = models.CharField(max_length=200, blank=True, default="")
    weight_kg = models.DecimalField(max_digits=10, decimal_places=4, null=True, blank=True)
    price_per_kg = models.DecimalField(max_digits=10, decimal_places=4, null=True, blank=True)
    tax = models.CharField(max_length=80, blank=True, default="")
    total_price = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True)
    label_image_path = models.CharField(max_length=500, blank=True, default="")
    label_values = models.JSONField(default=dict, blank=True)
    trigger = models.CharField(max_length=20, default="manual")
    archive = models.ForeignKey(PrintArchive, null=True, blank=True, on_delete=models.SET_NULL)

    class Meta:
        ordering = ["-printed_at", "-id"]

    def __str__(self):
        return f"{self.printed_at:%Y-%m-%d %H:%M:%S} {self.cut_name}"
