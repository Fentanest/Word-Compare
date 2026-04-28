import os

from PySide6.QtCore import Qt
from PySide6.QtGui import QStandardItem

from app.models import FilePair


class FileListManager:
    @staticmethod
    def append_file_item(model, file_path: str) -> None:
        item = QStandardItem(os.path.basename(file_path))
        item.setData(file_path, Qt.UserRole)
        item.setFlags(item.flags() & ~Qt.ItemIsDropEnabled)
        model.appendRow(item)

    @staticmethod
    def sort_model(model) -> None:
        items = []
        for row in range(model.rowCount()):
            item = model.item(row)
            items.append((item.text(), item.data(Qt.UserRole)))

        items.sort(key=lambda item: item[0])
        model.clear()

        for text, user_role_data in items:
            item = QStandardItem(text)
            item.setData(user_role_data, Qt.UserRole)
            item.setFlags(item.flags() & ~Qt.ItemIsDropEnabled)
            model.appendRow(item)

    @staticmethod
    def build_file_pairs(model_before, model_after) -> list[FilePair]:
        file_pairs: list[FilePair] = []
        row_count = min(model_before.rowCount(), model_after.rowCount())

        for row in range(row_count):
            before_item = model_before.item(row)
            after_item = model_after.item(row)
            file_pairs.append(
                FilePair(
                    before_path=before_item.data(Qt.UserRole),
                    after_path=after_item.data(Qt.UserRole),
                )
            )

        return file_pairs
