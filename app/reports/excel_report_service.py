from app.reports.diff_engine import ExcelDiffEngine
from app.reports.excel_writer import ExcelReportWriter
from app.reports.models import ExcelReportInput
from app.services.docx_extractor import DocxExtractor


class ExcelReportService:
    def __init__(
        self,
        extractor: DocxExtractor | None = None,
        diff_engine: ExcelDiffEngine | None = None,
        writer: ExcelReportWriter | None = None,
    ):
        self.extractor = extractor or DocxExtractor()
        self.diff_engine = diff_engine or ExcelDiffEngine()
        self.writer = writer or ExcelReportWriter()

    def generate(self, before_doc, after_doc, excel_save_path: str, log_callback) -> None:
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
            return f"{idx + 1}행"

        self.generate_from_extracted_data(
            excel_save_path=excel_save_path,
            log_callback=log_callback,
            paras_before=before_data.paragraphs,
            paras_after=after_data.paragraphs,
            flags_b=before_data.table_flags,
            flags_a=after_data.table_flags,
            tables_before=before_data.tables,
            tables_after=after_data.tables,
            get_loc_cb=get_loc_info,
        )

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
    ) -> None:
        report_input = ExcelReportInput(
            excel_save_path=excel_save_path,
            log_callback=log_callback,
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
        diff_plan = self.diff_engine.build_diff_plan(report_input)
        self.writer.write(report_input, diff_plan)
