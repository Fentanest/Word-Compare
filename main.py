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
        
        # Windows 단일 파일 빌드 환경에서 더 안정적인 IniFormat 사용
        self.settings = QSettings("settings.ini", QSettings.IniFormat)

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

        # 4. 저장 경로, 작업자명, 엑셀 체크 여부 등 설정 불러오기
        self.load_settings()

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

    def load_settings(self):
        # 저장 경로
        saved_path = self.settings.value("savePath", "")
        if saved_path and os.path.isdir(saved_path):
            self.lineEditSavePath.setText(saved_path)
        else:
            self.lineEditSavePath.setText(os.path.join(os.path.expanduser("~"), "Desktop"))
        
        # 작업자명
        saved_author = self.settings.value("author", "")
        self.textEditauthor.setPlainText(saved_author)
        
        # 엑셀 보고서 생성 체크 여부 (기본값 True)
        saved_excel_checked = self.settings.value("excelChecked", "true")
        self.checkBoxExcel.setChecked(saved_excel_checked == "true")

    def save_settings(self):
        self.settings.setValue("savePath", self.lineEditSavePath.text())
        self.settings.setValue("author", self.textEditauthor.toPlainText())
        self.settings.setValue("excelChecked", "true" if self.checkBoxExcel.isChecked() else "false")
        self.settings.sync() # 디스크에 즉시 저장 강제

    def closeEvent(self, event):
        self.save_settings()
        super().closeEvent(event)

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

    def extract_data_hybrid(self, doc, doc_name=""):
        """
        [Strategy 2.0] Hybrid XML Extraction.
        1. Finalize auto-numbers in Word.
        2. Save as temp docx.
        3. Use python-docx to parse XML elements (100% accuracy, High speed).
        """
        import tempfile
        from docx import Document as DocxReader
        
        try:
            self.log(f"-> '{doc_name}' 데이터 분석 및 고속 추출 준비 중...")
            # 1. 자동 번호를 텍스트로 확정 (Word 엔진 활용)
            doc.Content.ListFormat.ConvertNumbersToText()
            
            # 2. 임시 파일로 저장 (XML 직접 접근을 위함)
            fd, temp_path = tempfile.mkstemp(suffix=".docx", prefix="extract_")
            os.close(fd)
            doc.SaveAs(os.path.abspath(temp_path), FileFormat=12)
            
            # 3. python-docx로 파일 직접 열기 (Word 엔진 통신 없음 - 초고속)
            reader = DocxReader(temp_path)
            
            paras = []
            is_table_flags = []
            tables_data = []
            
            # 4. 문서 요소(문단/표) 순차 스캔
            # document.element.body를 직접 순회하여 문서 내 순서를 100% 보존
            from docx.document import Document as _Document
            from docx.table import Table as _Table
            from docx.text.paragraph import Paragraph as _Paragraph
            
            # Accessing internal elements to preserve order
            for child in reader.element.body:
                if child.tag.endswith('p'): # Paragraph
                    p_obj = _Paragraph(child, reader)
                    text = p_obj.text.strip()
                    # if text: # 빈 문단도 인덱스 유지를 위해 포함할 수 있음 (사용자 선택)
                    paras.append(text)
                    is_table_flags.append(False)
                elif child.tag.endswith('tbl'): # Table
                    t_obj = _Table(child, reader)
                    is_table_flags.append(True)
                    paras.append("[TABLE_MARKER]")
                    
                    # [ULTIMATE ROBUST EXTRACTION] 
                    # 병합된 셀이 많은 복잡한 표에서도 데이터 누락을 원천 차단하는 방식
                    table_grid = []
                    try:
                        # 각 행(row)이 가진 실제 셀(cell)들을 하나도 빠짐없이 리스트로 변환
                        # 병합된 셀의 경우, 워드는 해당 영역의 모든 칸에 동일한 텍스트를 할당함
                        for row in t_obj.rows:
                            row_data = [cell.text.replace('\r', '\n').strip() for cell in row.cells]
                            table_grid.append(row_data)
                        
                        # 추출된 표의 행/열 크기 로그 출력 (디버깅용)
                        # self.log(f"-> 표 추출 완료: {len(table_grid)}행 x {len(table_grid[0]) if table_grid else 0}열")
                        tables_data.append(table_grid)
                    except Exception as te:
                        self.log(f"-> 표 추출 중 오류: {te}")
                        tables_data.append([["[데이터 추출 실패]"]])
            
            # 임시 파일 삭제
            try: os.remove(temp_path)
            except: pass
            
            self.log(f"-> '{doc_name}' 데이터 추출 완료 (표 {len(tables_data)}개 발견)")
            return paras, is_table_flags, tables_data
            
        except Exception as e:
            self.log(f"-> 하이브리드 추출 오류: {e}")
            # Fallback (Very Slow)
            return [p.Range.Text for p in doc.Paragraphs], [False]*doc.Paragraphs.Count, []

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
            word_app.ScreenUpdating = False # [Strategy 3] 화면 갱신 억제로 속도 향상
            
            for i in range(before_count):
                # 매 파일마다 Visibility 재확인 (일부 상황에서 자동으로 True가 되는 것을 방지)
                if word_app.Visible: word_app.Visible = False
                
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

                    # 문서 열기 전에 다시 한 번 체크
                    if word_app.Visible: word_app.Visible = False
                    
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
                        
                        try:
                            # Use Strategy 2.0: Hybrid XML Extraction (Faster & Accurate)
                            paras_before, flags_b, tables_before = self.extract_data_hybrid(doc1, "수정 전 문서")
                            paras_after, flags_a, tables_after = self.extract_data_hybrid(doc2, "수정 후 문서")
                            
                            def get_loc_info(idx, is_before):
                                # Note: Hybrid mode might have slightly different paragraph count than doc.Paragraphs.Count
                                # because XML treats every <w:p> as a paragraph, including inside tables.
                                # This is a simpler fallback for location.
                                return f"{idx+1}행"

                            create_excel_report(
                                None, None, excel_save_path, self.log, 
                                paras_before, paras_after, get_loc_info,
                                flags_b, flags_a, tables_before, tables_after
                            )
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
                try:
                    word_app.ScreenUpdating = True # 설정 복원
                except:
                    pass
                word_app.Quit(SaveChanges=False)
        
        self.log("모든 비교 작업을 완료했습니다.")


if __name__ == '__main__':
    # 윈도우 단일 파일 빌드(.exe) 환경에서 멀티프로세싱 지원을 위해 필수
    import multiprocessing
    multiprocessing.freeze_support()

    try:
        from ctypes import windll
        windll.ole32.CoInitialize(None)
    except ImportError:
        pass

    app = QApplication(sys.argv)
    window = WordCompareApp()
    window.show()
    sys.exit(app.exec())