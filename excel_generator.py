from app.reports.excel_report_service import ExcelReportService

def create_excel_report(before_file_path, after_file_path, excel_save_path, log_callback, 
                        paras_before=None, paras_after=None, get_loc_cb=None,
                        flags_b=None, flags_a=None, tables_before=None, tables_after=None):
    service = ExcelReportService(extractor=None)
    service.generate_from_extracted_data(
        excel_save_path=excel_save_path,
        log_callback=log_callback,
        compare_formatting=False,
        paras_before=paras_before,
        paras_after=paras_after,
        flags_b=flags_b,
        flags_a=flags_a,
        tables_before=tables_before,
        tables_after=tables_after,
        get_loc_cb=get_loc_cb,
    )
