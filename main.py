
import sys
import os
import shutil
import tempfile
import unicodedata

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

import xlsxwriter
import win32com.client as win32

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QListWidgetItem, QAbstractItemView
)
from PySide6.QtCore import Qt, QUrl, QEvent, QSettings
from PySide6.QtGui import QStandardItemModel, QStandardItem, QIcon

from main_ui import Ui_MainWindow
from excel_generator import create_excel_report

class WordCompareApp(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowIcon(QIcon(resource_path('logo.png')))
        self.setWindowTitle("Word Compare Tool")
        self.settings = QSettings("MyCompany", "WordCompareTool")

        # 1. 모델 생성 및 리스트뷰에 설정
        self.model_before = QStandardItemModel()
        self.model_after = QStandardItemModel()
        self.listViewbefore.setModel(self.model_before)
        self.listViewafter.setModel(self.model_after)

        # 2. 드래그 앤 드롭 활성화
        self.listViewbefore.setDragDropMode(QAbstractItemView.DragDrop)
        self.listViewafter.setDragDropMode(QAbstractItemView.DragDrop)
        self.listViewbefore.setDefaultDropAction(Qt.MoveAction)
        self.listViewafter.setDefaultDropAction(Qt.MoveAction)
        self.listViewbefore.setAcceptDrops(True)
        self.listViewafter.setAcceptDrops(True)

        # 3. 버튼 시그널 연결
        self.btnStart.clicked.connect(self.start_compare)
        self.btnBrowsePath.clicked.connect(self.browse_path)
        self.btnOpenPath.clicked.connect(self.open_path)

        # 4. 저장 경로 불러오기 또는 기본값 설정
        self.load_initial_path()

        # 5. 리스트뷰 이벤트 필터 설치 (키 삭제용)
        self.listViewbefore.installEventFilter(self)
        self.listViewafter.installEventFilter(self)

    def load_initial_path(self):
        saved_path = self.settings.value("savePath", "")
        if saved_path and os.path.isdir(saved_path):
            self.lineEditSavePath.setText(saved_path)
        else:
            self.lineEditSavePath.setText(os.path.join(os.path.expanduser("~"), "Desktop"))



    def eventFilter(self, source, event):
        if event.type() == QEvent.KeyPress and event.key() == Qt.Key_Delete:
            if source is self.listViewbefore:
                self.remove_selected_items(self.listViewbefore)
                return True
            elif source is self.listViewafter:
                self.remove_selected_items(self.listViewafter)
                return True
        return super().eventFilter(source, event)

    def remove_selected_items(self, list_view):
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

        urls = event.mimeData().urls()
        
        target_list_view = None
        if self.listViewbefore.geometry().contains(event.position().toPoint()):
             target_list_view = self.listViewbefore
        elif self.listViewafter.geometry().contains(event.position().toPoint()):
             target_list_view = self.listViewafter

        if target_list_view:
            model = target_list_view.model()
            for url in urls:
                file_path = url.toLocalFile()
                if file_path.lower().endswith(('.doc', '.docx')):
                    file_name = os.path.basename(file_path)
                    item = QStandardItem(file_name)
                    item.setData(file_path, Qt.UserRole)
                    item.setFlags(item.flags() & ~Qt.ItemIsDropEnabled)
                    model.appendRow(item)
    
    def browse_path(self):
        path = QFileDialog.getExistingDirectory(self, "저장할 폴더 선택", self.lineEditSavePath.text())
        if path:
            self.lineEditSavePath.setText(path)
            self.settings.setValue("savePath", path)

    def open_path(self):
        path = self.lineEditSavePath.text()
        if os.path.isdir(path):
            os.startfile(path)
        else:
            self.log(f"경로를 열 수 없습니다: {path}")

    def log(self, message):
        self.txtLogOutput.append(message)
        QApplication.processEvents()

    def start_compare(self):
        before_count = self.model_before.rowCount()
        after_count = self.model_after.rowCount()

        if before_count == 0 or after_count == 0:
            self.log("오류: 비교할 파일이 없습니다. 파일을 리스트에 추가해주세요.")
            return

        if before_count != after_count:
            self.log(f"오류: '전' 파일 ({before_count}개)과 '후' 파일 ({after_count}개)의 개수가 일치하지 않습니다.")
            return

        save_dir = self.lineEditSavePath.text()
        if not os.path.isdir(save_dir):
            try:
                os.makedirs(save_dir)
                self.log(f"'{save_dir}' 폴더를 생성했습니다.")
            except OSError as e:
                self.log(f"오류: 저장 폴더를 생성할 수 없습니다. {e}")
                return
        
        self.log("비교 작업을 시작합니다...")
        
        word_app = None
        try:
            word_app = win32.gencache.EnsureDispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = 0

            for i in range(before_count):
                before_item = self.model_before.item(i)
                after_item = self.model_after.item(i)
                
                before_path_raw = before_item.data(Qt.UserRole)
                after_path_raw = after_item.data(Qt.UserRole)

                if not before_path_raw: before_path_raw = before_item.text()
                if not after_path_raw: after_path_raw = after_item.text()

                before_path = os.path.abspath(os.path.normpath(before_path_raw))
                after_path = os.path.abspath(os.path.normpath(after_path_raw))
                before_filename = os.path.basename(before_path)
                after_filename = os.path.basename(after_path) # Define after_filename here

                doc1, doc2, result_doc = None, None, None
                try:
                    self.log(f"'{before_filename}' 파일 여는 중 (변경 내용 적용)...")

                    # Open in R/W mode, accept revisions in memory, then compare
                    doc1 = word_app.Documents.Open(before_path)
                    doc1.Revisions.AcceptAll()

                    self.log(f"'{after_filename}' 파일 여는 중 (변경 내용 적용)...") # New log message

                    doc2 = word_app.Documents.Open(after_path)
                    doc2.Revisions.AcceptAll()

                    self.log(f"'{before_filename}'과 '{after_filename}' 파일 비교 중...") # Modify this line
                    
                    author_name = self.textEditauthor.toPlainText()
                    if not author_name.strip():
                        author_name = "Administrator"

                    result_doc = word_app.CompareDocuments(
                        OriginalDocument=doc1,
                        RevisedDocument=doc2,
                        Destination=2,
                        Granularity=1,
                        CompareMoves=True,
                        RevisedAuthor=author_name
                    )
                    
                    result_filename = f"비교_결과_{before_filename}"
                    result_save_path = os.path.join(save_dir, result_filename)
                    result_doc.SaveAs(os.path.abspath(result_save_path))
                    self.log(f"-> '비교 결과 문서' 저장: {result_save_path}")

                    if self.checkBoxExcel.isChecked():
                        excel_filename = f"변경내용_{os.path.splitext(before_filename)[0]}.xlsx"
                        excel_save_path = os.path.join(save_dir, excel_filename)
                        try:
                            create_excel_report(before_path, after_path, excel_save_path, self.log)
                        except Exception as e:
                            self.log(f"-> Excel 보고서 생성 중 오류 발생: {e}")
                    
                except Exception as e:
                    self.log(f"'{before_filename}' 비교 중 오류 발생: {e}")
                finally:
                    # Close original documents without saving changes
                    if doc1: doc1.Close(SaveChanges=False)
                    if doc2: doc2.Close(SaveChanges=False)
                    if result_doc: result_doc.Close(SaveChanges=False)

        except Exception as e:
            self.log(f"오류: Microsoft Word 처리 중 문제가 발생했습니다. ({e})")
        finally:
            if word_app:
                word_app.DisplayAlerts = -1
                word_app.Quit(SaveChanges=False)
        
        self.log("모든 비교 작업을 완료했습니다.")


if __name__ == '__main__':
    try:
        from ctypes import windll
        windll.ole32.CoInitialize(None)
    except ImportError:
        pass

    app = QApplication(sys.argv)
    window = WordCompareApp()
    window.show()
    sys.exit(app.exec())
