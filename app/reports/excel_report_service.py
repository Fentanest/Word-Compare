from time import perf_counter

from app.reports.diff_engine import ExcelDiffEngine
from app.reports.excel_writer import ExcelReportWriter
from app.reports.models import ExcelReportInput

_EXTRACTOR_UNSET = object()


class ExcelReportService:
    def __init__(
        self,
        extractor=_EXTRACTOR_UNSET,
        diff_engine: ExcelDiffEngine | None = None,
        writer: ExcelReportWriter | None = None,
    ):
        if extractor is _EXTRACTOR_UNSET:
            from app.services.docx_extractor import DocxExtractor

            extractor = DocxExtractor()

        self.extractor = extractor
        self.diff_engine = diff_engine or ExcelDiffEngine()
        self.writer = writer or ExcelReportWriter()

    def generate(self, before_doc, after_doc, excel_save_path: str, log_callback, compare_formatting: bool = False) -> None:
        if self.extractor is None:
            raise RuntimeError("문서 추출기가 없어 generate()를 사용할 수 없습니다.")

        total_started_at = perf_counter()
        self._log(log_callback, f"Excel 보고서 시작: {excel_save_path}")

        before_data = self.extractor.extract_data_hybrid(
            before_doc,
            log_callback,
            "수정 전 문서",
        )

        after_data = self.extractor.extract_data_hybrid(
            after_doc,
            log_callback,
            "수정 후 문서",
        )

        def get_loc_info(idx, is_before):
            locations = before_data.paragraph_locations if is_before else after_data.paragraph_locations
            if locations and 0 <= idx < len(locations):
                return locations[idx]
            return f"{idx + 1}행"

        self.generate_from_extracted_data(
            excel_save_path=excel_save_path,
            log_callback=log_callback,
            compare_formatting=compare_formatting,
            paras_before=before_data.paragraphs,
            paras_after=after_data.paragraphs,
            flags_b=before_data.table_flags,
            flags_a=after_data.table_flags,
            tables_before=before_data.tables,
            tables_after=after_data.tables,
            get_loc_cb=get_loc_info,
        )
        self._log_perf(log_callback, "Excel 보고서 전체", total_started_at)

    def generate_from_extracted_data(
        self,
        excel_save_path: str,
        log_callback,
        paras_before,
        paras_after,
        flags_b,
        flags_a,
        tables_before,
        tables_after,
        get_loc_cb=None,
        paragraph_locations_before=None,
        paragraph_locations_after=None,
        compare_formatting: bool = False,
    ) -> None:
        if get_loc_cb is None and (paragraph_locations_before or paragraph_locations_after):
            def get_loc_cb(idx, is_before):
                locations = paragraph_locations_before if is_before else paragraph_locations_after
                if locations and 0 <= idx < len(locations):
                    return locations[idx]
                return f"{idx + 1}행"

        report_input = ExcelReportInput(
            excel_save_path=excel_save_path,
            log_callback=log_callback,
            compare_formatting=compare_formatting,
            paras_before=paras_before,
            paras_after=paras_after,
            get_loc_cb=get_loc_cb,
            flags_b=flags_b,
            flags_a=flags_a,
            tables_before=tables_before,
            tables_after=tables_after,
        )
        self.generate_from_input(report_input)

    def generate_from_input(self, report_input: ExcelReportInput) -> None:
        diff_started_at = perf_counter()
        diff_plan = self.diff_engine.build_diff_plan(report_input)
        self._log_perf(report_input.log_callback, "Excel diff 계획 계산", diff_started_at)

        write_started_at = perf_counter()
        self.writer.write(report_input, diff_plan)
        self._log_perf(report_input.log_callback, "Excel 파일 쓰기", write_started_at)

    @staticmethod
    def _log(log_callback, message: str) -> None:
        if log_callback:
            log_callback(message)

    @staticmethod
    def _log_perf(log_callback, label: str, started_at: float) -> None:
        ExcelReportService._log(log_callback, f"{label}: {perf_counter() - started_at:.3f}초")
