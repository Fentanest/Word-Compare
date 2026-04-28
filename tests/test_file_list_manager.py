import os
import unittest

from PySide6.QtCore import Qt
from PySide6.QtGui import QStandardItemModel
from PySide6.QtWidgets import QApplication

from app.ui.file_list_manager import FileListManager


class FileListManagerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_append_file_item_stores_basename_and_path(self):
        model = QStandardItemModel()

        FileListManager.append_file_item(model, "/tmp/example.docx")

        self.assertEqual(model.rowCount(), 1)
        item = model.item(0)
        self.assertEqual(item.text(), "example.docx")
        self.assertEqual(item.data(Qt.UserRole), "/tmp/example.docx")
        self.assertFalse(bool(item.flags() & Qt.ItemIsDropEnabled))

    def test_sort_model_orders_items_by_visible_filename(self):
        model = QStandardItemModel()
        FileListManager.append_file_item(model, "/tmp/zeta.docx")
        FileListManager.append_file_item(model, "/tmp/alpha.docx")

        FileListManager.sort_model(model)

        self.assertEqual(model.item(0).text(), "alpha.docx")
        self.assertEqual(model.item(1).text(), "zeta.docx")

    def test_build_file_pairs_matches_rows_by_index(self):
        before_model = QStandardItemModel()
        after_model = QStandardItemModel()
        FileListManager.append_file_item(before_model, "/tmp/before-1.docx")
        FileListManager.append_file_item(before_model, "/tmp/before-2.docx")
        FileListManager.append_file_item(after_model, "/tmp/after-1.docx")
        FileListManager.append_file_item(after_model, "/tmp/after-2.docx")

        file_pairs = FileListManager.build_file_pairs(before_model, after_model)

        self.assertEqual(len(file_pairs), 2)
        self.assertEqual(file_pairs[0].before_path, "/tmp/before-1.docx")
        self.assertEqual(file_pairs[0].after_path, "/tmp/after-1.docx")
        self.assertEqual(file_pairs[1].source_name, os.path.basename("/tmp/after-2.docx"))


if __name__ == "__main__":
    unittest.main()
