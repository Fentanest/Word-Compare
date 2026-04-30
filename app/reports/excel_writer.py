import re
from difflib import SequenceMatcher

import xlsxwriter

from app.models import ParagraphData, TableCellData
from app.reports.models import ExcelDiffPlan, ExcelReportInput


class ExcelReportWriter:
    def write(self, report_input: ExcelReportInput, diff_plan: ExcelDiffPlan) -> None:
        self._log(report_input, "-> Excel 보고서(양방향 정밀 서식) 생성 중...")

        workbook = xlsxwriter.Workbook(report_input.excel_save_path)
        formats = self._build_formats(workbook)

        self._write_main_sheet(workbook, formats, report_input, diff_plan)
        self._write_table_sheets(workbook, formats, diff_plan)

        workbook.close()
        self._log(report_input, f"-> 양방향 정밀 보고서 저장 완료: {report_input.excel_save_path}")

    def _write_main_sheet(self, workbook, formats, report_input: ExcelReportInput, diff_plan: ExcelDiffPlan) -> None:
        worksheet = workbook.add_worksheet("변경 내용(일반)")
        headers = ["위치", "수정 전", "수정 후"]
        if report_input.compare_formatting:
            headers.append("서식 변경")
        worksheet.write_row("A1", headers, formats["header"])
        worksheet.set_column("A:A", 25, formats["loc"])
        worksheet.set_column("B:C", 60, formats["default"])
        if report_input.compare_formatting:
            worksheet.set_column("D:D", 12, formats["marker"])
        worksheet.freeze_panes(1, 0)

        excel_row = 1
        for tag, i1, i2, j1, j2 in diff_plan.main_opcodes:
            if tag == "equal":
                continue
            content_before = self._paragraph_block_text(diff_plan.filtered_paras_before[i1:i2])
            content_after = self._paragraph_block_text(diff_plan.filtered_paras_after[j1:j2])
            if not content_before and not content_after:
                continue

            style_changed = False
            if content_before == content_after:
                marker = self._change_marker(
                    diff_plan.filtered_paras_before[i1:i2],
                    diff_plan.filtered_paras_after[j1:j2],
                )
                if marker == "서식 변경" and report_input.compare_formatting:
                    style_changed = True
                elif marker == "서식 변경":
                    continue
                else:
                    content_before = self._mark_change(content_before, marker)
                    content_after = self._mark_change(content_after, marker)

            rich_before, rich_after = self._get_rich_diff(content_before, content_after, formats)
            worksheet.write(
                excel_row,
                0,
                self._resolve_location(report_input, diff_plan, i1, j1),
                formats["loc"],
            )

            for column, rich_data, plain_text in (
                (1, rich_before, content_before),
                (2, rich_after, content_after),
            ):
                self._write_rich_or_plain(
                    worksheet,
                    excel_row,
                    column,
                    rich_data,
                    plain_text,
                    formats,
                )
            if report_input.compare_formatting:
                worksheet.write(excel_row, 3, "O" if style_changed else "", formats["marker"])
            excel_row += 1

    def _write_table_sheets(self, workbook, formats, diff_plan: ExcelDiffPlan) -> None:
        for table_plan in diff_plan.tables:
            sheet_name = self._table_sheet_name(table_plan)
            worksheet = workbook.add_worksheet(sheet_name[:31])

            max_cols_before = max((len(row) for row in table_plan.before_table), default=0)
            before_width = max(1, max_cols_before)
            after_start_col = before_width + 1

            if table_plan.before_index is None:
                before_header = "수정 전 없음"
                after_header = "수정 후만 있음"
            elif table_plan.after_index is None:
                before_header = "수정 전만 있음"
                after_header = "수정 후 없음"
            else:
                before_header = "수정 전"
                after_header = "수정 후"

            worksheet.write(0, 0, before_header, formats["header"])
            worksheet.write(0, after_start_col, after_header, formats["header"])

            if table_plan.before_index is None:
                worksheet.write(1, after_start_col, "이 표는 수정 후 문서에만 있습니다.", formats["note"])
            elif table_plan.after_index is None:
                worksheet.write(1, 0, "이 표는 수정 전 문서에만 있습니다.", formats["note"])

            row_map_after_to_before, row_map_before_to_after = self._build_bidirectional_map(
                table_plan.row_opcodes
            )
            col_map_after_to_before, col_map_before_to_after = self._build_bidirectional_map(
                table_plan.col_opcodes
            )

            for row_index, row in enumerate(table_plan.before_table):
                for col_index, cell_before in enumerate(row):
                    target_row = row_map_before_to_after.get(row_index)
                    target_col = col_map_before_to_after.get(col_index)

                    value_before = self._cell_text(cell_before)
                    is_changed = True
                    if target_row is not None and target_col is not None:
                        try:
                            cell_after = table_plan.after_table[target_row][target_col]
                            value_after = self._cell_text(cell_after)
                            if self._cell_signature(cell_before, report_input.compare_formatting) == self._cell_signature(
                                cell_after,
                                report_input.compare_formatting,
                            ):
                                is_changed = False
                            else:
                                rich_before, _ = self._get_rich_diff(value_before, value_after, formats)
                                if len(rich_before) >= 3:
                                    worksheet.write_rich_string(
                                        row_index + 2,
                                        col_index,
                                        *rich_before,
                                        formats["table_cell"],
                                    )
                                    continue
                        except Exception:
                            pass

                    worksheet.write(
                        row_index + 2,
                        col_index,
                        value_before,
                        formats["table_del"] if is_changed else formats["table_cell"],
                    )

            for row_index, row in enumerate(table_plan.after_table):
                for col_index, cell_after in enumerate(row):
                    original_row = row_map_after_to_before.get(row_index)
                    original_col = col_map_after_to_before.get(col_index)

                    value_after = self._cell_text(cell_after)
                    is_changed = True
                    if original_row is not None and original_col is not None:
                        try:
                            cell_before = table_plan.before_table[original_row][original_col]
                            value_before = self._cell_text(cell_before)
                            if self._cell_signature(cell_before, report_input.compare_formatting) == self._cell_signature(
                                cell_after,
                                report_input.compare_formatting,
                            ):
                                is_changed = False
                            else:
                                _, rich_after = self._get_rich_diff(value_before, value_after, formats)
                                if len(rich_after) >= 3:
                                    worksheet.write_rich_string(
                                        row_index + 2,
                                        col_index + after_start_col,
                                        *rich_after,
                                        formats["table_cell"],
                                    )
                                    continue
                        except Exception:
                            pass

                    worksheet.write(
                        row_index + 2,
                        col_index + after_start_col,
                        value_after,
                        formats["table_ins"] if is_changed else formats["table_cell"],
                    )

    @staticmethod
    def _build_formats(workbook):
        return {
            "header": workbook.add_format(
                {"bold": True, "align": "center", "valign": "vcenter", "border": 1, "bg_color": "#D3D3D3"}
            ),
            "del": workbook.add_format(
                {"font_color": "blue", "font_strikeout": True, "valign": "vcenter", "text_wrap": True}
            ),
            "ins": workbook.add_format(
                {"font_color": "red", "bold": True, "valign": "vcenter", "text_wrap": True}
            ),
            "default": workbook.add_format({"valign": "vcenter", "text_wrap": True}),
            "loc": workbook.add_format({"align": "center", "valign": "vcenter", "text_wrap": True}),
            "marker": workbook.add_format({"align": "center", "valign": "vcenter"}),
            "note": workbook.add_format({"italic": True, "font_color": "#666666"}),
            "table_cell": workbook.add_format({"valign": "vcenter", "border": 1, "text_wrap": True}),
            "table_ins": workbook.add_format(
                {"font_color": "red", "bold": True, "valign": "vcenter", "border": 1, "text_wrap": True}
            ),
            "table_del": workbook.add_format(
                {"font_color": "blue", "font_strikeout": True, "valign": "vcenter", "border": 1, "text_wrap": True}
            ),
        }

    @staticmethod
    def _get_rich_diff(text_before: str, text_after: str, formats):
        if not text_before:
            return [], [formats["ins"], text_after]
        if not text_after:
            return [formats["del"], text_before], []

        words_before = [word for word in re.split(r"(\s+)", text_before) if word]
        words_after = [word for word in re.split(r"(\s+)", text_after) if word]
        matcher = SequenceMatcher(None, words_before, words_after, autojunk=False)

        rich_before = []
        rich_after = []
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            fragment_before = "".join(words_before[i1:i2])
            fragment_after = "".join(words_after[j1:j2])
            if tag == "equal":
                if fragment_before:
                    rich_before.extend([formats["default"], fragment_before])
                    rich_after.extend([formats["default"], fragment_after])
            elif tag == "delete":
                rich_before.extend([formats["del"], fragment_before])
            elif tag == "insert":
                rich_after.extend([formats["ins"], fragment_after])
            elif tag == "replace":
                rich_before.extend([formats["del"], fragment_before])
                rich_after.extend([formats["ins"], fragment_after])

        return rich_before, rich_after

    @staticmethod
    def _write_rich_or_plain(worksheet, row, col, rich_data, plain_text, formats) -> None:
        has_text_change = ExcelReportWriter._has_text_change_markup(rich_data, formats)

        if has_text_change and len(rich_data) >= 3 and len(rich_data) <= 500:
            try:
                worksheet.write_rich_string(row, col, *rich_data, formats["default"])
                return
            except Exception:
                pass

        if col == 2 and has_text_change:
            fallback_format = formats["ins"]
        elif col == 1 and has_text_change:
            fallback_format = formats["del"]
        else:
            fallback_format = formats["default"]

        worksheet.write(row, col, plain_text, fallback_format)

    @staticmethod
    def _has_text_change_markup(rich_data, formats) -> bool:
        return any(
            item is formats["del"] or item is formats["ins"]
            for item in rich_data
            if not isinstance(item, str)
        )

    @staticmethod
    def _build_bidirectional_map(opcodes):
        map_after_to_before = {}
        map_before_to_after = {}
        for tag, i1, i2, j1, j2 in opcodes:
            if tag in ("equal", "replace"):
                for before_index, after_index in zip(range(i1, i2), range(j1, j2)):
                    map_after_to_before[after_index] = before_index
                    map_before_to_after[before_index] = after_index
        return map_after_to_before, map_before_to_after

    @staticmethod
    def _resolve_location(report_input: ExcelReportInput, diff_plan: ExcelDiffPlan, i1: int, j1: int) -> str:
        if not report_input.get_loc_cb:
            return "문단"

        if i1 < len(diff_plan.original_indices_before):
            return report_input.get_loc_cb(diff_plan.original_indices_before[i1], True)
        if j1 < len(diff_plan.original_indices_after):
            return report_input.get_loc_cb(diff_plan.original_indices_after[j1], False)
        return "문단"

    @staticmethod
    def _log(report_input: ExcelReportInput, message: str) -> None:
        if report_input.log_callback:
            report_input.log_callback(message)

    @staticmethod
    def _paragraph_block_text(paragraphs) -> str:
        return "\n".join(ExcelReportWriter._paragraph_text(paragraph) for paragraph in paragraphs).strip()

    @staticmethod
    def _paragraph_text(value) -> str:
        if isinstance(value, ParagraphData):
            return value.text
        return str(value)

    @staticmethod
    def _mark_change(text: str, marker: str) -> str:
        stripped = text.strip()
        if stripped:
            return f"{stripped} [{marker}]"
        return f"[{marker}]"

    @staticmethod
    def _change_marker(before_paragraphs, after_paragraphs) -> str:
        paragraph_items = list(before_paragraphs) + list(after_paragraphs)
        if not paragraph_items:
            return "서식/구조 변경"

        if all(
            isinstance(item, ParagraphData) and item.source_kind == "body"
            for item in paragraph_items
        ):
            return "서식 변경"
        return "구조/메타데이터 변경"

    @staticmethod
    def _cell_text(value) -> str:
        if isinstance(value, TableCellData):
            return value.text
        return str(value)

    @staticmethod
    def _cell_signature(value, compare_formatting: bool = True):
        if isinstance(value, TableCellData):
            if not compare_formatting:
                return (value.text,)
            return value.signature
        if not compare_formatting:
            return (str(value),)
        return (str(value), 1, "", 0, 0, 0, "", "", "", "", 0)

    @staticmethod
    def _table_sheet_name(table_plan) -> str:
        suffix = ""
        if table_plan.before_index is None:
            suffix = " (수정 후만 있음)"
        elif table_plan.after_index is None:
            suffix = " (수정 전만 있음)"
        return f"표 {table_plan.index + 1}{suffix}"
