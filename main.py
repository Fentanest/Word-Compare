
import sys
import os
import shutil
import tempfile
import unicodedata
import xlsxwriter
import win32com.client as win32
import re

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QListWidgetItem, QAbstractItemView
)
from PySide6.QtCore import Qt, QUrl, QEvent, QSettings
from PySide6.QtGui import QStandardItemModel, QStandardItem

from main_ui import Ui_MainWindow

class KoreanCleaner:
    @classmethod
    def _normalize_numbers(cls, text):
        # Implement number normalization logic here if provided
        return text

    @classmethod
    def _normalize_english_text(cls, text):
        # Implement English text normalization logic here if provided
        return text

    @classmethod
    def normalize_text(cls, text):
        text = text.strip()
        text = cls._normalize_numbers(text)
        text = cls._normalize_english_text(text)
        return text

class WordCompareApp(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
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
            word_app.Visible = True
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

                try:
                    self.log(f"'{before_filename}' 파일 비교 중...")
                    
                    doc1 = word_app.Documents.Open(before_path, ReadOnly=True)
                    doc2 = word_app.Documents.Open(after_path, ReadOnly=True)

                    # 사용자 지정 author 이름 사용
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
                    
                    doc1.Close(SaveChanges=False)
                    doc2.Close(SaveChanges=False)

                    result_filename = f"비교_결과_{before_filename}"
                    result_save_path = os.path.join(save_dir, result_filename)
                    result_doc.SaveAs(os.path.abspath(result_save_path))
                    self.log(f"-> '비교 결과 문서' 저장: {result_save_path}")

                    if self.checkBoxExcel.isChecked():
                        self.create_excel_report(result_doc, before_filename, save_dir)
                    
                    result_doc.Close(SaveChanges=False)

                except Exception as e:
                    self.log(f"'{before_filename}' 비교 중 오류 발생: {e}")

        except Exception as e:
            self.log(f"오류: Microsoft Word 처리 중 문제가 발생했습니다. ({e})")
        finally:
            if word_app:
                word_app.DisplayAlerts = -1
                word_app.Quit(SaveChanges=False)
        
        self.log("모든 비교 작업을 완료했습니다.")


    def _sanitize_text(self, text):
        if not isinstance(text, str):
            return ""
        # Simply remove carriage returns and strip whitespace
        return text.replace('\r', '').strip()

    def create_excel_report(self, doc, original_filename, save_dir):
        excel_filename = f"변경내용_{os.path.splitext(original_filename)[0]}.xlsx"
        excel_save_path = os.path.join(save_dir, excel_filename)
        self.log(f"-> Excel 보고서 생성 중 (빠른 모드)...")

        try:
            workbook = xlsxwriter.Workbook(excel_save_path)
            worksheet = workbook.add_worksheet("변경 내용")

            # 서식 정의
            header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
            deleted_format = workbook.add_format({'font_color': 'blue', 'font_strikeout': True})
            inserted_format = workbook.add_format({'font_color': 'red', 'bold': True})
            
            # 헤더 설정
            worksheet.write('A1', '위치', header_format)
            worksheet.write('B1', '수정 전', header_format)
            worksheet.write('C1', '수정 후', header_format)
            worksheet.set_column('A:A', 15)
            worksheet.set_column('B:B', 50)
            worksheet.set_column('C:C', 50)
            worksheet.freeze_panes(1, 0)
            
            row = 1
            
            revisions = doc.Revisions
            if revisions.Count == 0:
                self.log("텍스트 변경 사항이 없어 Excel 보고서를 생성하지 않습니다.")
                workbook.close()
                os.remove(excel_save_path) # 빈 엑셀 파일 삭제
                return

            self.log(f"총 {revisions.Count}개의 변경점을 처리합니다.")
            
            # 문서의 모든 변경 사항을 직접 순회 (가장 빠름)
            for i, rev in enumerate(revisions):
                if (i + 1) % 10 == 0:
                    self.log(f"변경점 처리 중... ({i + 1}/{revisions.Count})")

                # 텍스트 변경(삽입, 삭제)만 처리
                if rev.Type == 1 or rev.Type == 2:
                    page = rev.Range.Information(3)
                    line = rev.Range.Information(10)
                    location_str = f"Page {page}, Line {line}"
                    
                    worksheet.write(row, 0, location_str)

                    sanitized_text = self._sanitize_text(rev.Range.Text)
                    normalized_text = KoreanCleaner.normalize_text(sanitized_text)

                    if rev.Type == 1:  # 삽입
                        worksheet.write(row, 2, normalized_text, inserted_format)
                    elif rev.Type == 2:  # 삭제
                        worksheet.write(row, 1, normalized_text, deleted_format)
                    
                    row += 1

            workbook.close()
            self.log(f"-> Excel 보고서 저장: {excel_save_path}")
        except Exception as e:
            import traceback
            self.log(f"-> Excel 저장 실패: {e}")
            self.log(traceback.format_exc())


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
