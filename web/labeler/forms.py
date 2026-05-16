from decimal import Decimal

from django import forms

from .models import CutPreset, SessionSettings, WorkstationState


class CutPresetForm(forms.ModelForm):
    class Meta:
        model = CutPreset
        fields = ["cut_name", "price_per_kg", "tax", "active"]
        widgets = {
            "cut_name": forms.TextInput(attrs={"class": "ui input"}),
            "price_per_kg": forms.NumberInput(attrs={"step": "0.01"}),
            "tax": forms.TextInput(),
        }


class SessionSettingsForm(forms.ModelForm):
    class Meta:
        model = SessionSettings
        fields = [
            "farm_name",
            "logo_path",
            "animal_number",
            "farm_number",
            "due_date_4_7",
            "due_date_frozen",
            "birth_country",
            "life_country",
            "slaughter_country",
            "packaged_country",
            "identity_mark",
        ]


class WorkstationForm(forms.ModelForm):
    class Meta:
        model = WorkstationState
        fields = [
            "cut_name",
            "price_per_kg",
            "tax",
            "auto_print_enabled",
            "cut_paper",
            "stable_iterations",
        ]

    def clean_stable_iterations(self):
        value = self.cleaned_data["stable_iterations"]
        return max(1, min(1000, int(value)))

    def clean_price_per_kg(self):
        value = self.cleaned_data["price_per_kg"]
        if value is None:
            return None
        return Decimal(value).quantize(Decimal("0.01"))
