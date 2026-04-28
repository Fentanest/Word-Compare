import os

from PySide6.QtCore import QEvent, Qt, QUrl
from PySide6.QtGui import QAction, QDesktopServices, QIcon, QStandardItem, QStandardItemModel
from PySide6.QtWidgets import QApplication, QAbstractItemView, QFileDialog, QMainWindow

from app.models import AppSettings, CompareOptions, FilePair
from app.resources import resource_path
from app.services.settings_service import SettingsService
from app.services.word_compare_service import WordCompareService
from main_ui import Ui_MainWindow
from version import __version__


class WordCompareApp(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowIcon(QIcon(resource_path("logo.png")))
        self.setWindowTitle(f"Word Compare Tool v{__version__}")

        self.settings_service = SettingsService()
        self.word_compare_service = WordCompareService()

        self.model_before = QStandardItemModel()
        self.model_after = QStandardItemModel()
        self.listViewbefore.setModel(self.model_before)
        self.listViewafter.setModel(self.model_after)

        self._configure_list_view(self.listViewbefore)
        self._configure_list_view(self.listViewafter)

        self.btnStart.clicked.connect(self.start_compare)
        self.btnBrowsePath.clicked.connect(self.browse_path)
        self.btnOpenPath.clicked.connect(self.open_path)

        self.listViewbefore.installEventFilter(self)
        self.listViewafter.installEventFilter(self)

        self.actionGithub.triggered.connect(self.open_github_link)
        self.actionBlog.triggered.connect(self.open_blog_link)
        self.actionSorting.triggered.connect(self.sort_list_views)

        version_action = QAction(f"Version: {__version__}", self)
        version_action.setEnabled(False)
        self.menuMade_by_Fentanest.addAction(version_action)

        self.load_settings()

    def _configure_list_view(self, list_view) -> None:
        list_view.setDragDropMode(QAbstractItemView.DragDrop)
        list_view.setDefaultDropAction(Qt.MoveAction)
        list_view.setAcceptDrops(True)

    def load_settings(self) -> None:
        settings = self.settings_service.load()
        self.lineEditSavePath.setText(settings.save_path)
        self.textEditauthor.setPlainText(settings.author)
        self.checkBoxExcel.setChecked(settings.excel_checked)

    def save_settings(self) -> None:
        self.settings_service.save(
            AppSettings(
                save_path=self.lineEditSavePath.text(),
                author=self.textEditauthor.toPlainText(),
                excel_checked=self.checkBoxExcel.isChecked(),
            )
        )

    def closeEvent(self, event):
        self.save_settings()
        super().closeEvent(event)

    def eventFilter(self, source, event):
        if event.type() == QEvent.KeyPress and event.key() == Qt.Key_Delete:
            if source is self.listViewbefore:
                self.remove_selected_items(self.listViewbefore)
                return True
            if source is self.listViewafter:
                self.remove_selected_items(self.listViewafter)
                return True
        return super().eventFilter(source, event)

    def remove_selected_items(self, list_view) -> None:
        model = list_view.model()
        for index in reversed(sorted(list_view.selectedIndexes())):
            model.removeRow(index.row())

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        if not event.mimeData().hasUrls():
            return

        target_list_view = None
        if self.listViewbefore.geometry().contains(event.position().toPoint()):
            target_list_view = self.listViewbefore
        elif self.listViewafter.geometry().contains(event.position().toPoint()):
            target_list_view = self.listViewafter

        if not target_list_view:
            return

        model = target_list_view.model()
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.lower().endswith((".doc", ".docx")):
                self._append_file_item(model, file_path)

    def _append_file_item(self, model, file_path: str) -> None:
        item = QStandardItem(os.path.basename(file_path))
        item.setData(file_path, Qt.UserRole)
        item.setFlags(item.flags() & ~Qt.ItemIsDropEnabled)
        model.appendRow(item)

    def browse_path(self):
        path = QFileDialog.getExistingDirectory(
            self,
            "저장할 폴더 선택",
            self.lineEditSavePath.text(),
        )
        if path:
            self.lineEditSavePath.setText(path)
            self.save_settings()

    def open_path(self):
        path = self.lineEditSavePath.text()
        if os.path.isdir(path):
            os.startfile(path)
        else:
            self.log(f"경로를 열 수 없습니다: {path}")

    def log(self, message: str) -> None:
        self.txtLogOutput.append(message)
        QApplication.processEvents()

    def open_github_link(self):
        QDesktopServices.openUrl(QUrl("https://github.com/Fentanest/Word-Compare"))

    def open_blog_link(self):
        QDesktopServices.openUrl(QUrl("https://hb.worklazy.net/word-compare"))
        self.log("블로그 링크를 열었습니다.")

    def sort_list_views(self):
        self._sort_model(self.model_before)
        self._sort_model(self.model_after)
        self.log("리스트를 파일 이름으로 오름차순 정렬했습니다.")

    def _sort_model(self, model) -> None:
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

    def start_compare(self):
        before_count = self.model_before.rowCount()
        after_count = self.model_after.rowCount()

        if before_count == 0 or after_count == 0:
            self.log("오류: 비교할 파일이 없습니다. 파일을 리스트에 추가해주세요.")
            return

        if before_count != after_count:
            self.log(
                f"오류: '전' 파일 ({before_count}개)과 '후' 파일 ({after_count}개)의 개수가 일치하지 않습니다."
            )
            return

        save_dir = self.lineEditSavePath.text()
        if not os.path.isdir(save_dir):
            try:
                os.makedirs(save_dir)
                self.log(f"'{save_dir}' 폴더를 생성했습니다.")
            except OSError as error:
                self.log(f"오류: 저장 폴더를 생성할 수 없습니다. {error}")
                return

        file_pairs = self._build_file_pairs()
        options = CompareOptions(
            save_dir=save_dir,
            author_name=self.textEditauthor.toPlainText(),
            generate_excel=self.checkBoxExcel.isChecked(),
        )
        self.word_compare_service.compare_pairs(file_pairs, options, self.log)

    def _build_file_pairs(self) -> list[FilePair]:
        file_pairs: list[FilePair] = []
        row_count = min(self.model_before.rowCount(), self.model_after.rowCount())

        for row in range(row_count):
            before_item = self.model_before.item(row)
            after_item = self.model_after.item(row)
            file_pairs.append(
                FilePair(
                    before_path=before_item.data(Qt.UserRole),
                    after_path=after_item.data(Qt.UserRole),
                )
            )

        return file_pairs
