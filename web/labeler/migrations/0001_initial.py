from decimal import Decimal

import django.db.models.deletion
from django.db import migrations, models
from django.utils import timezone


class Migration(migrations.Migration):
    initial = True

    dependencies = []

    operations = [
        migrations.CreateModel(
            name="AppSetting",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("key", models.CharField(max_length=80, unique=True)),
                ("value", models.TextField(blank=True, default="")),
            ],
        ),
        migrations.CreateModel(
            name="CutPreset",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("cut_name", models.CharField(max_length=200)),
                ("price_per_kg", models.DecimalField(decimal_places=2, default=Decimal("0.00"), max_digits=10)),
                ("tax", models.CharField(blank=True, default="", max_length=80)),
                ("sort_order", models.PositiveIntegerField(default=0)),
                ("active", models.BooleanField(default=True)),
            ],
            options={"ordering": ["sort_order", "cut_name", "id"]},
        ),
        migrations.CreateModel(
            name="LabelFieldConfig",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("key", models.CharField(max_length=120)),
                ("print_name", models.CharField(blank=True, default="", max_length=300)),
                ("show", models.BooleanField(default=True)),
                ("font_size", models.PositiveIntegerField(default=24)),
                ("sort_order", models.PositiveIntegerField(default=0)),
            ],
            options={"ordering": ["sort_order", "id"]},
        ),
        migrations.CreateModel(
            name="PrintArchive",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("created_at", models.DateTimeField(default=timezone.now)),
                ("note", models.CharField(blank=True, default="", max_length=200)),
            ],
            options={"ordering": ["-created_at", "-id"]},
        ),
        migrations.CreateModel(
            name="ScaleState",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("connected", models.BooleanField(default=False)),
                ("status", models.CharField(default="Scale: starting...", max_length=200)),
                ("current_weight_kg", models.DecimalField(blank=True, decimal_places=4, max_digits=10, null=True)),
                ("stable_weight_kg", models.DecimalField(blank=True, decimal_places=4, max_digits=10, null=True)),
                ("same_value_iterations", models.PositiveIntegerField(default=0)),
                ("last_error", models.TextField(blank=True, default="")),
                ("updated_at", models.DateTimeField(default=timezone.now)),
            ],
        ),
        migrations.CreateModel(
            name="SessionSettings",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("farm_name", models.CharField(blank=True, default="", max_length=200)),
                ("logo_path", models.CharField(blank=True, default="", max_length=500)),
                ("animal_number", models.CharField(blank=True, default="", max_length=200)),
                ("farm_number", models.CharField(blank=True, default="", max_length=200)),
                ("due_date_4_7", models.CharField(blank=True, default="", max_length=200)),
                ("due_date_frozen", models.CharField(blank=True, default="", max_length=200)),
                ("birth_country", models.CharField(blank=True, default="", max_length=200)),
                ("life_country", models.CharField(blank=True, default="", max_length=200)),
                ("slaughter_country", models.CharField(blank=True, default="", max_length=200)),
                ("packaged_country", models.CharField(blank=True, default="", max_length=200)),
                ("identity_mark", models.CharField(blank=True, default="", max_length=200)),
            ],
        ),
        migrations.CreateModel(
            name="PrintLog",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("printed_at", models.DateTimeField(default=timezone.now)),
                ("cut_name", models.CharField(blank=True, default="", max_length=200)),
                ("weight_kg", models.DecimalField(blank=True, decimal_places=4, max_digits=10, null=True)),
                ("price_per_kg", models.DecimalField(blank=True, decimal_places=4, max_digits=10, null=True)),
                ("tax", models.CharField(blank=True, default="", max_length=80)),
                ("total_price", models.DecimalField(blank=True, decimal_places=2, max_digits=10, null=True)),
                ("label_image_path", models.CharField(blank=True, default="", max_length=500)),
                ("label_values", models.JSONField(blank=True, default=dict)),
                ("trigger", models.CharField(default="manual", max_length=20)),
                ("archive", models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, to="labeler.printarchive")),
            ],
            options={"ordering": ["-printed_at", "-id"]},
        ),
        migrations.CreateModel(
            name="WorkstationState",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("cut_name", models.CharField(blank=True, default="", max_length=200)),
                ("price_per_kg", models.DecimalField(blank=True, decimal_places=2, max_digits=10, null=True)),
                ("tax", models.CharField(blank=True, default="", max_length=80)),
                ("auto_print_enabled", models.BooleanField(default=False)),
                ("cut_paper", models.BooleanField(default=False)),
                ("stable_iterations", models.PositiveIntegerField(default=10)),
                ("last_auto_printed_weight", models.DecimalField(blank=True, decimal_places=4, max_digits=10, null=True)),
                ("must_zero_before_next_print", models.BooleanField(default=False)),
                ("auto_print_in_progress", models.BooleanField(default=False)),
                ("status_message", models.CharField(blank=True, default="", max_length=300)),
                ("updated_at", models.DateTimeField(auto_now=True)),
                ("selected_cut", models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, to="labeler.cutpreset")),
            ],
        ),
    ]
