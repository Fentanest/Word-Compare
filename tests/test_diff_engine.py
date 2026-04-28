import unittest

from app.reports.diff_engine import ExcelDiffEngine
from app.reports.models import ExcelReportInput


class ExcelDiffEngineTests(unittest.TestCase):
    def test_build_diff_plan_filters_table_markers_and_keeps_original_indices(self):
        engine = ExcelDiffEngine()
        report_input = ExcelReportInput(
            excel_save_path="unused.xlsx",
            paras_before=["keep", "[TABLE_MARKER]", "old changed"],
            paras_after=["keep", "[TABLE_MARKER]", "new changed"],
            flags_b=[False, True, False],
            flags_a=[False, True, False],
            tables_before=[[["A", "B"], ["C", "D"]]],
            tables_after=[[["A", "B2"], ["C", "D"]]],
        )

        diff_plan = engine.build_diff_plan(report_input)

        self.assertEqual(diff_plan.filtered_paras_before, ["keep", "old changed"])
        self.assertEqual(diff_plan.filtered_paras_after, ["keep", "new changed"])
        self.assertEqual(diff_plan.original_indices_before, [0, 2])
        self.assertEqual(diff_plan.original_indices_after, [0, 2])
        self.assertEqual(len(diff_plan.tables), 1)
        self.assertTrue(
            any(tag == "replace" for tag, *_ in diff_plan.main_opcodes),
            "Expected a replace opcode for the changed body paragraph.",
        )
        self.assertTrue(
            any(tag == "replace" for tag, *_ in diff_plan.tables[0].col_opcodes),
            "Expected a replace opcode for the changed table column heading.",
        )


if __name__ == "__main__":
    unittest.main()
