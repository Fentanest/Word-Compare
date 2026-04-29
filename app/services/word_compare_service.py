import os
from time import perf_counter

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
        total_started_at = perf_counter()
        self._log(log_callback, "비교 작업을 시작합니다...")
        self._log(log_callback, f"비교 대상 {len(file_pairs)}건")

        try:
            word_session_started_at = perf_counter()
            with WordSession() as word_app:
                self._log_perf(log_callback, "Word 세션 준비", word_session_started_at)
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

        self._log_perf(log_callback, "전체 비교", total_started_at)
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
        report_doc_before = None
        report_doc_after = None
        result_doc = None

        try:
            file_started_at = perf_counter()
            self._log(log_callback, f"'{original_filename}' 파일 처리 중...")
            WordSession.ensure_hidden(word_app)

            open_started_at = perf_counter()
            doc_before = word_app.Documents.Open(before_path)
            doc_after = word_app.Documents.Open(after_path)
            self._log_perf(log_callback, f"{original_filename} 문서 열기", open_started_at)

            normalize_started_at = perf_counter()
            doc_before.Revisions.AcceptAll()
            doc_before.TrackRevisions = False
            doc_after.Revisions.AcceptAll()
            doc_after.TrackRevisions = False
            self._log_perf(log_callback, f"{original_filename} 비교 전 정리", normalize_started_at)

            self._log(log_callback, f"'{original_filename}' 비교 중...")
            compare_started_at = perf_counter()
            result_doc = word_app.CompareDocuments(
                OriginalDocument=doc_before,
                RevisedDocument=doc_after,
                Destination=2,
                Granularity=1,
                CompareMoves=True,
                RevisedAuthor=options.effective_author_name,
                IgnoreAllComparisonWarnings=True,
            )
            self._log_perf(log_callback, f"{original_filename} Word 비교", compare_started_at)

            result_filename = f"비교_결과_{original_filename}"
            result_docx_path = os.path.join(options.save_dir, result_filename)
            save_started_at = perf_counter()
            result_doc.SaveAs(os.path.abspath(result_docx_path))
            self._log_perf(log_callback, f"{original_filename} 결과 저장", save_started_at)
            self._log(log_callback, f"-> '비교 결과 문서' 저장: {result_docx_path}")

            result_excel_path = None
            if options.generate_excel:
                result_excel_path = os.path.join(
                    options.save_dir,
                    f"변경내용_{os.path.splitext(original_filename)[0]}.xlsx",
                )
                try:
                    excel_open_started_at = perf_counter()
                    report_doc_before = word_app.Documents.Open(before_path)
                    report_doc_after = word_app.Documents.Open(after_path)
                    self._log_perf(log_callback, f"{original_filename} Excel용 원본 재열기", excel_open_started_at)
                    excel_started_at = perf_counter()
                    self.excel_report_service.generate(
                        report_doc_before,
                        report_doc_after,
                        result_excel_path,
                        log_callback,
                    )
                    self._log_perf(log_callback, f"{original_filename} Excel 보고서 생성", excel_started_at)
                except Exception as error:
                    self._log(log_callback, f"-> Excel 보고서 생성 중 오류 발생: {error}")
                    result_excel_path = None

            self._log_perf(log_callback, f"{original_filename} 전체 처리", file_started_at)
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
            self._safe_close_document(doc_before)
            self._safe_close_document(doc_after)
            self._safe_close_document(report_doc_before)
            self._safe_close_document(report_doc_after)
            self._safe_close_document(result_doc)

    @staticmethod
    def _log(log_callback, message: str) -> None:
        if log_callback:
            log_callback(message)

    @staticmethod
    def _log_perf(log_callback, label: str, started_at: float) -> None:
        WordCompareService._log(log_callback, f"{label}: {perf_counter() - started_at:.3f}초")

    @staticmethod
    def _safe_close_document(document) -> None:
        if not document:
            return
        try:
            document.Close(SaveChanges=False)
        except Exception:
            # Word may tear down COM-backed document proxies during shutdown or
            # after compare/save operations. Cleanup noise should not be treated
            # as a user-facing compare failure.
            pass
