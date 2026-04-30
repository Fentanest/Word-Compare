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
        compare_formatting = report_input.compare_formatting
        excluded_source_kinds = set(report_input.excluded_source_kinds or ())
        filtered_paras_before, original_indices_before = self._filter_paragraphs(
            report_input.paras_before,
            report_input.flags_b,
            excluded_source_kinds,
        )
        filtered_paras_after, original_indices_after = self._filter_paragraphs(
            report_input.paras_after,
            report_input.flags_a,
            excluded_source_kinds,
        )

        tables_before = report_input.tables_before or []
        tables_after = report_input.tables_after or []
        aligned_tables = self._align_tables(tables_before, tables_after, compare_formatting)

        tasks = [
            (
                self._build_paragraph_signatures(filtered_paras_before, compare_formatting),
                self._build_paragraph_signatures(filtered_paras_after, compare_formatting),
            )
        ]
        for _, _, _, before_table, after_table in aligned_tables:
            tasks.append(
                (
                    self._build_row_signatures(before_table, compare_formatting),
                    self._build_row_signatures(after_table, compare_formatting),
                )
            )
            tasks.append(
                (
                    self._build_column_signatures(before_table, compare_formatting),
                    self._build_column_signatures(after_table, compare_formatting),
                )
            )

        all_results = self._run_tasks(tasks)
        table_plans: list[TableDiffPlan] = []

        for plan_index, (display_index, before_index, after_index, before_table, after_table) in enumerate(aligned_tables):
            table_plans.append(
                TableDiffPlan(
                    index=display_index,
                    before_index=before_index,
                    after_index=after_index,
                    before_table=before_table,
                    after_table=after_table,
                    row_opcodes=all_results[1 + plan_index * 2],
                    col_opcodes=all_results[2 + plan_index * 2],
                    before_metadata=self._resolve_table_metadata(report_input.table_metadata_before, before_index),
                    after_metadata=self._resolve_table_metadata(report_input.table_metadata_after, after_index),
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
    def _filter_paragraphs(paragraphs, flags, excluded_source_kinds):
        paragraphs = paragraphs or []
        if not flags:
            filtered = []
            indices = []
            for index, paragraph in enumerate(paragraphs):
                if isinstance(paragraph, ParagraphData) and paragraph.source_kind in excluded_source_kinds:
                    continue
                filtered.append(paragraph)
                indices.append(index)
            return filtered, indices

        filtered = []
        indices = []
        for index, paragraph in enumerate(paragraphs):
            if flags[index]:
                continue
            if isinstance(paragraph, ParagraphData) and paragraph.source_kind in excluded_source_kinds:
                continue
            if not flags[index]:
                filtered.append(paragraph)
                indices.append(index)
        return filtered, indices

    @staticmethod
    def _resolve_table_metadata(table_metadata, table_index):
        if table_index is None or not table_metadata:
            return None
        return table_metadata.get(table_index)

    @staticmethod
    def _build_row_signatures(table, compare_formatting: bool):
        signatures = []
        for row in table:
            signatures.append(tuple(ExcelDiffEngine._cell_alignment_signature(cell, compare_formatting) for cell in row))
        return signatures

    @staticmethod
    def _build_paragraph_signatures(paragraphs, compare_formatting: bool):
        return [ExcelDiffEngine._paragraph_signature(paragraph, compare_formatting) for paragraph in paragraphs]

    @staticmethod
    def _build_column_signatures(table, compare_formatting: bool):
        max_cols = max((len(row) for row in table), default=0)
        signatures = []
        missing_signature = ExcelDiffEngine._missing_cell_alignment_signature(compare_formatting)
        for col_index in range(max_cols):
            column_signature = []
            for row in table:
                if col_index < len(row):
                    column_signature.append(ExcelDiffEngine._cell_alignment_signature(row[col_index], compare_formatting))
                else:
                    column_signature.append(missing_signature)
            signatures.append(tuple(column_signature))
        return signatures

    @staticmethod
    def _normalize_text(value):
        return str(value).replace("\r", "\n").strip()

    @staticmethod
    def _normalize_alignment_text(value):
        return normalize_alignment_text(str(value))

    def _align_tables(self, tables_before, tables_after, compare_formatting: bool):
        before_signatures = [self._table_signature(table, compare_formatting) for table in tables_before]
        after_signatures = [self._table_signature(table, compare_formatting) for table in tables_after]
        table_opcodes = _run_comparison_task((before_signatures, after_signatures))

        aligned_tables = []
        display_index = 0
        for tag, i1, i2, j1, j2 in table_opcodes:
            if tag in ("equal", "replace"):
                pair_count = min(i2 - i1, j2 - j1)
                for offset in range(pair_count):
                    before_index = i1 + offset
                    after_index = j1 + offset
                    aligned_tables.append(
                        (
                            display_index,
                            before_index,
                            after_index,
                            tables_before[before_index],
                            tables_after[after_index],
                        )
                    )
                    display_index += 1

                for before_index in range(i1 + pair_count, i2):
                    aligned_tables.append((display_index, before_index, None, tables_before[before_index], []))
                    display_index += 1

                for after_index in range(j1 + pair_count, j2):
                    aligned_tables.append((display_index, None, after_index, [], tables_after[after_index]))
                    display_index += 1
            elif tag == "delete":
                for before_index in range(i1, i2):
                    aligned_tables.append((display_index, before_index, None, tables_before[before_index], []))
                    display_index += 1
            elif tag == "insert":
                for after_index in range(j1, j2):
                    aligned_tables.append((display_index, None, after_index, [], tables_after[after_index]))
                    display_index += 1

        return aligned_tables

    @staticmethod
    def _table_signature(table, compare_formatting: bool):
        max_cols = max((len(row) for row in table), default=0)
        first_row_signature = ()
        if table:
            first_row_signature = tuple(
                ExcelDiffEngine._table_identity_cell(cell, compare_formatting) for cell in table[0][: min(len(table[0]), 6)]
            )
        return (
            max_cols,
            first_row_signature,
        )

    @staticmethod
    def _paragraph_signature(value, compare_formatting: bool):
        if isinstance(value, ParagraphData):
            if not compare_formatting and value.source_kind == "body":
                return (ExcelDiffEngine._normalize_text(value.text),)
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
    def _cell_alignment_signature(value, compare_formatting: bool):
        if isinstance(value, TableCellData):
            if not compare_formatting:
                return (ExcelDiffEngine._normalize_alignment_text(value.text),)
            return value.alignment_signature
        return (ExcelDiffEngine._normalize_alignment_text(value),) if not compare_formatting else (
            ExcelDiffEngine._normalize_alignment_text(value), 1, "", 0, 0, 0, "", "", "", "", 0
        )

    @staticmethod
    def _table_identity_cell(value, compare_formatting: bool):
        if isinstance(value, TableCellData):
            if not compare_formatting:
                return ExcelDiffEngine._normalize_alignment_text(value.text)
            return (
                ExcelDiffEngine._normalize_alignment_text(value.text),
                value.grid_span,
                value.v_merge,
            )
        if not compare_formatting:
            return ExcelDiffEngine._normalize_alignment_text(value)
        return (ExcelDiffEngine._normalize_alignment_text(value), 1, "")

    @staticmethod
    def _missing_cell_alignment_signature(compare_formatting: bool):
        if not compare_formatting:
            return ("__RHWP_MISSING_CELL__",)
        return ("__RHWP_MISSING_CELL__", 0, "__RHWP_MISSING_CELL__", 0, 0, 0, "", "", "", "", 0)
