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

from version import __version__

import xlsxwriter
import win32com.client as win32

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QListWidgetItem, QAbstractItemView
)
from PySide6.QtCore import Qt, QUrl, QEvent, QSettings
from PySide6.QtGui import QStandardItemModel, QStandardItem, QIcon, QDesktopServices, QAction

from main_ui import Ui_MainWindow
from excel_generator import create_excel_report

class WordCompareApp(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowIcon(QIcon(resource_path('logo.png')))
        self.setWindowTitle(f"Word Compare Tool v{__version__}")
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

        # 6. Action 연결
        self.actionGithub.triggered.connect(self.open_github_link)
        self.actionBlog.triggered.connect(self.open_blog_link)
        self.actionSorting.triggered.connect(self.sort_list_views)

        # 7. 'Made by Fentanest' 메뉴에 버전 정보 Action 추가
        version_action = QAction(f"Version: {__version__}", self)
        version_action.setEnabled(False) # Make it unclickable
        self.menuMade_by_Fentanest.addAction(version_action)

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

    def open_github_link(self):
        github_url = "https://github.com/Fentanest/Word-Compare"
        QDesktopServices.openUrl(QUrl(github_url))

    def open_blog_link(self):
        blog_url = "https://hb.worklazy.net/word-compare" # Placeholder: Please update with your actual blog link
        QDesktopServices.openUrl(QUrl(blog_url))
        self.log("블로그 링크를 열었습니다.")

    def sort_list_views(self):
        self._sort_model(self.model_before)
        self._sort_model(self.model_after)
        self.log("리스트를 파일 이름으로 오름차순 정렬했습니다.")

    def _sort_model(self, model):
        # Extract items, sort them, and re-populate the model
        items = []
        for row in range(model.rowCount()):
            item = model.item(row)
            items.append((item.text(), item.data(Qt.UserRole))) # Store (filename, path)

        # Sort in ascending order by filename
        items.sort(key=lambda x: x[0], reverse=False)

        # Clear existing model and add sorted items
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
            # word_app.AutomationSecurity = 3 # Removed as IgnoreAllComparisonWarnings should be sufficient
            
            for i in range(before_count):
                before_item = self.model_before.item(i)
                after_item = self.model_after.item(i)
                
                original_before_path = os.path.abspath(before_item.data(Qt.UserRole))
                original_after_path = os.path.abspath(after_item.data(Qt.UserRole))
                original_filename = os.path.basename(original_after_path)
                
                doc1, doc2, result_doc = None, None, None
                excel_temp_before_path = None
                excel_temp_after_path = None

                try:
                    self.log(f"'{original_filename}' 파일 처리 중...")

                    # Open original documents
                    doc1 = word_app.Documents.Open(original_before_path)
                    doc2 = word_app.Documents.Open(original_after_path)
                    
                    # Accept revisions and turn off tracking in memory
                    doc1.Revisions.AcceptAll()
                    doc1.TrackRevisions = False
                    doc2.Revisions.AcceptAll()
                    doc2.TrackRevisions = False

                    self.log(f"'{original_filename}' 비교 중...")
                    author_name = self.textEditauthor.toPlainText()
                    if not author_name.strip():
                        author_name = "Administrator"
                    
                    result_doc = word_app.CompareDocuments(
                        OriginalDocument=doc1,
                        RevisedDocument=doc2,
                        Destination=2, # Create new document
                        Granularity=1, # Word by word
                        CompareMoves=True,
                        RevisedAuthor=author_name,
                        IgnoreAllComparisonWarnings=True # New parameter to suppress warnings
                    )
                    
                    result_filename = f"비교_결과_{original_filename}"
                    result_save_path = os.path.join(save_dir, result_filename)
                    result_doc.SaveAs(os.path.abspath(result_save_path))
                    self.log(f"-> '비교 결과 문서' 저장: {result_save_path}")

                    if self.checkBoxExcel.isChecked():
                        excel_filename = f"변경내용_{os.path.splitext(original_filename)[0]}.xlsx"
                        excel_save_path = os.path.join(save_dir, excel_filename)
                        
                        self.log(f"-> '{original_filename}' Excel 보고서 생성 준비 중 (임시 파일 생성)...")
                        # Save in-memory cleaned docs to temp files for excel_generator
                        fd_excel_before, excel_temp_before_path = tempfile.mkstemp(suffix=".docx", prefix="excel_before_")
                        os.close(fd_excel_before)
                        doc1.SaveAs(os.path.abspath(excel_temp_before_path), FileFormat=12) # Save the cleaned doc1
                        
                        fd_excel_after, excel_temp_after_path = tempfile.mkstemp(suffix=".docx", prefix="excel_after_")
                        os.close(fd_excel_after)
                        doc2.SaveAs(os.path.abspath(excel_temp_after_path), FileFormat=12) # Save the cleaned doc2

                        try:
                            # Use the temporary clean files for Excel report
                            create_excel_report(excel_temp_before_path, excel_temp_after_path, excel_save_path, self.log)
                        except Exception as e:
                            self.log(f"-> Excel 보고서 생성 중 오류 발생: {e}")
                    
                except Exception as e:
                    self.log(f"'{original_filename}' 처리 중 오류 발생: {e}")
                finally:
                    # Close all documents opened in the loop
                    if doc1: doc1.Close(SaveChanges=False)
                    if doc2: doc2.Close(SaveChanges=False)
                    if result_doc: result_doc.Close(SaveChanges=False)
                    
                    # Clean up temporary files for Excel report
                    if excel_temp_before_path and os.path.exists(excel_temp_before_path):
                        try: os.remove(excel_temp_before_path)
                        except OSError as e: self.log(f"임시 파일 삭제 오류 '{excel_temp_before_path}': {e}")
                    if excel_temp_after_path and os.path.exists(excel_temp_after_path):
                        try: os.remove(excel_temp_after_path)
                        except OSError as e: self.log(f"임시 파일 삭제 오류 '{excel_temp_after_path}': {e}")

        except Exception as e:
            self.log(f"오류: Microsoft Word 처리 중 문제가 발생했습니다. ({e})")
        finally:
            if word_app:
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