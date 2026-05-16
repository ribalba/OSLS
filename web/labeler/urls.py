from django.urls import path

from . import views

urlpatterns = [
    path("", views.index, name="index"),
    path("preview.png", views.preview_png, name="preview_png"),
    path("cuts/", views.cuts, name="cuts"),
    path("session/", views.session_settings, name="session_settings"),
    path("label-config/", views.label_config, name="label_config"),
    path("analytics/", views.analytics, name="analytics"),
    path("settings/", views.hardware_settings, name="hardware_settings"),
]
