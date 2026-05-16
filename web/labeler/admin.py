from django.contrib import admin

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


admin.site.register(AppSetting)
admin.site.register(CutPreset)
admin.site.register(SessionSettings)
admin.site.register(LabelFieldConfig)
admin.site.register(ScaleState)
admin.site.register(WorkstationState)
admin.site.register(PrintArchive)
admin.site.register(PrintLog)
