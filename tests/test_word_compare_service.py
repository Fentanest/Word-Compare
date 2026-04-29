import tempfile
import unittest
from pathlib import Path

from app.models import CompareOptions, FilePair
from app.services.word_compare_service import WordCompareService


class _FakeRevisions:
    def __init__(self):
        self.accepted = False

    def AcceptAll(self):
        self.accepted = True


class _FakeDocument:
    def __init__(self, path: str, fail_on_close: bool = False):
        self.path = path
        self.Revisions = _FakeRevisions()
        self.TrackRevisions = True
        self.closed = False
        self.saved_path = None
        self.fail_on_close = fail_on_close

    def Close(self, SaveChanges=False):
        self.closed = True
        if self.fail_on_close:
            raise RuntimeError("COM disconnected during document close")

    def SaveAs(self, path):
        self.saved_path = path


class _FakeDocuments:
    def __init__(self, fail_on_close: bool = False):
        self.opened_docs = []
        self.fail_on_close = fail_on_close

    def Open(self, path):
        document = _FakeDocument(path, fail_on_close=self.fail_on_close)
        self.opened_docs.append(document)
        return document


class _FakeWordApp:
    def __init__(self, fail_on_close: bool = False):
        self.Documents = _FakeDocuments(fail_on_close=fail_on_close)
        self.compare_inputs = None

    def CompareDocuments(self, **kwargs):
        self.compare_inputs = kwargs
        return _FakeDocument("result.docx", fail_on_close=self.Documents.fail_on_close)


class _FakeExcelReportService:
    def __init__(self):
        self.calls = []

    def generate(self, before_doc, after_doc, excel_save_path, log_callback):
        self.calls.append((before_doc, after_doc, excel_save_path))


class WordCompareServiceTests(unittest.TestCase):
    def test_excel_generation_uses_fresh_source_documents(self):
        excel_service = _FakeExcelReportService()
        service = WordCompareService(excel_report_service=excel_service)
        word_app = _FakeWordApp()

        with tempfile.TemporaryDirectory() as temp_dir:
            before_path = str(Path(temp_dir) / "before.docx")
            after_path = str(Path(temp_dir) / "after.docx")
            Path(before_path).write_text("before", encoding="utf-8")
            Path(after_path).write_text("after", encoding="utf-8")

            result = service._compare_single_pair(
                word_app=word_app,
                file_pair=FilePair(before_path=before_path, after_path=after_path),
                options=CompareOptions(
                    save_dir=temp_dir,
                    author_name="Tester",
                    generate_excel=True,
                ),
                log_callback=None,
            )

        self.assertIsNone(result.error)
        self.assertEqual(len(word_app.Documents.opened_docs), 4)
        compare_before, compare_after, report_before, report_after = word_app.Documents.opened_docs

        self.assertTrue(compare_before.Revisions.accepted)
        self.assertTrue(compare_after.Revisions.accepted)
        self.assertFalse(report_before.Revisions.accepted)
        self.assertFalse(report_after.Revisions.accepted)
        self.assertEqual(len(excel_service.calls), 1)
        self.assertIs(excel_service.calls[0][0], report_before)
        self.assertIs(excel_service.calls[0][1], report_after)
        self.assertTrue(compare_before.closed)
        self.assertTrue(compare_after.closed)
        self.assertTrue(report_before.closed)
        self.assertTrue(report_after.closed)

    def test_close_disconnect_does_not_turn_success_into_error(self):
        excel_service = _FakeExcelReportService()
        service = WordCompareService(excel_report_service=excel_service)
        word_app = _FakeWordApp(fail_on_close=True)

        with tempfile.TemporaryDirectory() as temp_dir:
            before_path = str(Path(temp_dir) / "before.docx")
            after_path = str(Path(temp_dir) / "after.docx")
            Path(before_path).write_text("before", encoding="utf-8")
            Path(after_path).write_text("after", encoding="utf-8")

            result = service._compare_single_pair(
                word_app=word_app,
                file_pair=FilePair(before_path=before_path, after_path=after_path),
                options=CompareOptions(
                    save_dir=temp_dir,
                    author_name="Tester",
                    generate_excel=True,
                ),
                log_callback=None,
            )

        self.assertIsNone(result.error)
        self.assertTrue(result.result_docx_path)


if __name__ == "__main__":
    unittest.main()
