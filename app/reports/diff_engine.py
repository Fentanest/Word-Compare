import concurrent.futures
from difflib import SequenceMatcher

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

        tasks = [(filtered_paras_before, filtered_paras_after)]
        for table_index in range(max_tables):
            before_table = tables_before[table_index] if table_index < len(tables_before) else []
            after_table = tables_after[table_index] if table_index < len(tables_after) else []
            tasks.append(
                (
                    [str(row[0]).strip() if row else "" for row in before_table],
                    [str(row[0]).strip() if row else "" for row in after_table],
                )
            )
            tasks.append(
                (
                    [str(cell).strip() for cell in before_table[0]] if before_table and before_table[0] else [],
                    [str(cell).strip() for cell in after_table[0]] if after_table and after_table[0] else [],
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
