import concurrent.futures
from difflib import SequenceMatcher

from app.models import ParagraphData, TableCellData, normalize_alignment_text
from app.reports.models import ExcelDiffPlan, ExcelReportInput, TableDiffPlan


def _run_comparison_task(args):
    data_b, data_a = args
    matcher = SequenceMatcher(None, data_b, data_a, autojunk=False)
    return matcher.get_opcodes()


class ExcelDiffEngine:
    def build_diff_plan(self, report_input: ExcelReportInput) -> ExcelDiffPlan:
        filtered_paras_before, original_indices_before = self._filter_paragraphs(
            report_input.paras_before,
            report_input.flags_b,
        )
        filtered_paras_after, original_indices_after = self._filter_paragraphs(
            report_input.paras_after,
            report_input.flags_a,
        )

        tables_before = report_input.tables_before or []
        tables_after = report_input.tables_after or []
        max_tables = max(len(tables_before), len(tables_after))

        tasks = [
            (
                self._build_paragraph_signatures(filtered_paras_before),
                self._build_paragraph_signatures(filtered_paras_after),
            )
        ]
        for table_index in range(max_tables):
            before_table = tables_before[table_index] if table_index < len(tables_before) else []
            after_table = tables_after[table_index] if table_index < len(tables_after) else []
            tasks.append(
                (
                    self._build_row_signatures(before_table),
                    self._build_row_signatures(after_table),
                )
            )
            tasks.append(
                (
                    self._build_column_signatures(before_table),
                    self._build_column_signatures(after_table),
                )
            )

        all_results = self._run_tasks(tasks)
        table_plans: list[TableDiffPlan] = []

        for table_index in range(max_tables):
            before_table = tables_before[table_index] if table_index < len(tables_before) else []
            after_table = tables_after[table_index] if table_index < len(tables_after) else []
            table_plans.append(
                TableDiffPlan(
                    index=table_index,
                    before_table=before_table,
                    after_table=after_table,
                    row_opcodes=all_results[1 + table_index * 2],
                    col_opcodes=all_results[2 + table_index * 2],
                )
            )

        return ExcelDiffPlan(
            filtered_paras_before=filtered_paras_before,
            filtered_paras_after=filtered_paras_after,
            original_indices_before=original_indices_before,
            original_indices_after=original_indices_after,
            main_opcodes=all_results[0],
            tables=table_plans,
        )

    @staticmethod
    def _run_tasks(tasks):
        try:
            with concurrent.futures.ProcessPoolExecutor() as executor:
                return list(executor.map(_run_comparison_task, tasks))
        except Exception:
            return [_run_comparison_task(task) for task in tasks]

    @staticmethod
    def _filter_paragraphs(paragraphs, flags):
        paragraphs = paragraphs or []
        if not flags:
            indices = list(range(len(paragraphs)))
            return paragraphs, indices

        filtered = []
        indices = []
        for index, paragraph in enumerate(paragraphs):
            if not flags[index]:
                filtered.append(paragraph)
                indices.append(index)
        return filtered, indices

    @staticmethod
    def _build_row_signatures(table):
        signatures = []
        for row in table:
            signatures.append(tuple(ExcelDiffEngine._cell_alignment_signature(cell) for cell in row))
        return signatures

    @staticmethod
    def _build_paragraph_signatures(paragraphs):
        return [ExcelDiffEngine._paragraph_signature(paragraph) for paragraph in paragraphs]

    @staticmethod
    def _build_column_signatures(table):
        max_cols = max((len(row) for row in table), default=0)
        signatures = []
        for col_index in range(max_cols):
            column_signature = []
            for row in table:
                if col_index < len(row):
                    column_signature.append(ExcelDiffEngine._cell_alignment_signature(row[col_index]))
                else:
                    column_signature.append(
                        ("__RHWP_MISSING_CELL__", 0, "__RHWP_MISSING_CELL__", 0, 0, 0, "", "", "", "", 0)
                    )
            signatures.append(tuple(column_signature))
        return signatures

    @staticmethod
    def _normalize_text(value):
        return str(value).replace("\r", "\n").strip()

    @staticmethod
    def _normalize_alignment_text(value):
        return normalize_alignment_text(str(value))

    @staticmethod
    def _paragraph_signature(value):
        if isinstance(value, ParagraphData):
            return value.signature
        return (ExcelDiffEngine._normalize_text(value),)

    @staticmethod
    def _cell_signature(value):
        if isinstance(value, TableCellData):
            return (
                ExcelDiffEngine._normalize_text(value.text),
                value.grid_span,
                value.v_merge,
                value.cell_width,
                value.row_height,
                value.grid_col_width,
                value.border_signature,
                value.shading_fill,
                value.vertical_align,
                value.text_direction,
                value.nested_table_count,
            )
        return (ExcelDiffEngine._normalize_text(value), 1, "", 0, 0, 0, "", "", "", "", 0)

    @staticmethod
    def _cell_alignment_signature(value):
        if isinstance(value, TableCellData):
            return value.alignment_signature
        return (ExcelDiffEngine._normalize_alignment_text(value), 1, "", 0, 0, 0, "", "", "", "", 0)
