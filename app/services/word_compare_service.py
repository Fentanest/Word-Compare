import os

from app.models import CompareOptions, CompareResult, FilePair
from app.reports.excel_report_service import ExcelReportService
from app.services.word_session import WordSession


class WordCompareService:
    def __init__(self, excel_report_service: ExcelReportService | None = None):
        self.excel_report_service = excel_report_service or ExcelReportService()

    def compare_pairs(
        self,
        file_pairs: list[FilePair],
        options: CompareOptions,
        log_callback,
    ) -> list[CompareResult]:
        results: list[CompareResult] = []
        self._log(log_callback, "비교 작업을 시작합니다...")

        try:
            with WordSession() as word_app:
                for file_pair in file_pairs:
                    results.append(
                        self._compare_single_pair(
                            word_app,
                            file_pair,
                            options,
                            log_callback,
                        )
                    )
        except Exception as error:
            self._log(log_callback, f"오류: Microsoft Word 처리 중 문제가 발생했습니다. ({error})")

        self._log(log_callback, "모든 비교 작업을 완료했습니다.")
        return results

    def _compare_single_pair(
        self,
        word_app,
        file_pair: FilePair,
        options: CompareOptions,
        log_callback,
    ) -> CompareResult:
        before_path = os.path.abspath(file_pair.before_path)
        after_path = os.path.abspath(file_pair.after_path)
        original_filename = os.path.basename(after_path)

        doc_before = None
        doc_after = None
        result_doc = None

        try:
            self._log(log_callback, f"'{original_filename}' 파일 처리 중...")
            WordSession.ensure_hidden(word_app)

            doc_before = word_app.Documents.Open(before_path)
            doc_after = word_app.Documents.Open(after_path)

            doc_before.Revisions.AcceptAll()
            doc_before.TrackRevisions = False
            doc_after.Revisions.AcceptAll()
            doc_after.TrackRevisions = False

            self._log(log_callback, f"'{original_filename}' 비교 중...")
            result_doc = word_app.CompareDocuments(
                OriginalDocument=doc_before,
                RevisedDocument=doc_after,
                Destination=2,
                Granularity=1,
                CompareMoves=True,
                RevisedAuthor=options.effective_author_name,
                IgnoreAllComparisonWarnings=True,
            )

            result_filename = f"비교_결과_{original_filename}"
            result_docx_path = os.path.join(options.save_dir, result_filename)
            result_doc.SaveAs(os.path.abspath(result_docx_path))
            self._log(log_callback, f"-> '비교 결과 문서' 저장: {result_docx_path}")

            result_excel_path = None
            if options.generate_excel:
                result_excel_path = os.path.join(
                    options.save_dir,
                    f"변경내용_{os.path.splitext(original_filename)[0]}.xlsx",
                )
                try:
                    self.excel_report_service.generate(
                        doc_before,
                        doc_after,
                        result_excel_path,
                        log_callback,
                    )
                except Exception as error:
                    self._log(log_callback, f"-> Excel 보고서 생성 중 오류 발생: {error}")
                    result_excel_path = None

            return CompareResult(
                source_name=original_filename,
                result_docx_path=result_docx_path,
                result_excel_path=result_excel_path,
            )
        except Exception as error:
            self._log(log_callback, f"'{original_filename}' 처리 중 오류 발생: {error}")
            return CompareResult(
                source_name=original_filename,
                error=str(error),
            )
        finally:
            if doc_before:
                doc_before.Close(SaveChanges=False)
            if doc_after:
                doc_after.Close(SaveChanges=False)
            if result_doc:
                result_doc.Close(SaveChanges=False)

    @staticmethod
    def _log(log_callback, message: str) -> None:
        if log_callback:
            log_callback(message)
