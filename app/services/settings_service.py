import os

from PySide6.QtCore import QSettings

from app.models import AppSettings


class SettingsService:
    def __init__(self, settings: QSettings | None = None):
        self.settings = settings or QSettings("settings.ini", QSettings.IniFormat)

    def load(self) -> AppSettings:
        save_path = self.settings.value("savePath", "")
        if not save_path or not os.path.isdir(save_path):
            save_path = os.path.join(os.path.expanduser("~"), "Desktop")

        author = self.settings.value("author", "")
        excel_checked = self.settings.value("excelChecked", "true") == "true"
        format_compare_enabled = self.settings.value("formatCompareEnabled", "false") == "true"
        return AppSettings(
            save_path=save_path,
            author=author,
            excel_checked=excel_checked,
            format_compare_enabled=format_compare_enabled,
        )

    def save(self, app_settings: AppSettings) -> None:
        self.settings.setValue("savePath", app_settings.save_path)
        self.settings.setValue("author", app_settings.author)
        self.settings.setValue(
            "excelChecked",
            "true" if app_settings.excel_checked else "false",
        )
        self.settings.setValue(
            "formatCompareEnabled",
            "true" if app_settings.format_compare_enabled else "false",
        )
        self.settings.sync()
