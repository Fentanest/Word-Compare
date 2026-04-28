import unittest

from app.models import TableCellData
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

    def test_row_shift_alignment_uses_full_row_signature_instead_of_first_cell_only(self):
        engine = ExcelDiffEngine()
        report_input = ExcelReportInput(
            excel_save_path="unused.xlsx",
            tables_before=[
                [
                    ["", "ID", "NAME"],
                    ["", "1", "ALPHA"],
                    ["", "2", "BETA"],
                ]
            ],
            tables_after=[
                [
                    ["", "ID", "NAME"],
                    ["", "0", "INTRO"],
                    ["", "1", "ALPHA"],
                    ["", "2", "BETA"],
                ]
            ],
        )

        diff_plan = engine.build_diff_plan(report_input)
        row_opcodes = diff_plan.tables[0].row_opcodes

        self.assertIn(("insert", 1, 1, 1, 2), row_opcodes)
        self.assertIn(("equal", 1, 3, 2, 4), row_opcodes)

    def test_column_shift_alignment_uses_full_column_signature_instead_of_header_only(self):
        engine = ExcelDiffEngine()
        report_input = ExcelReportInput(
            excel_save_path="unused.xlsx",
            tables_before=[
                [
                    ["", "", ""],
                    ["ROW1", "A1", "B1"],
                    ["ROW2", "A2", "B2"],
                ]
            ],
            tables_after=[
                [
                    ["", "", "", ""],
                    ["ROW1", "X", "A1", "B1"],
                    ["ROW2", "Y", "A2", "B2"],
                ]
            ],
        )

        diff_plan = engine.build_diff_plan(report_input)
        col_opcodes = diff_plan.tables[0].col_opcodes

        self.assertIn(("insert", 1, 1, 1, 2), col_opcodes)
        self.assertIn(("equal", 1, 3, 2, 4), col_opcodes)

    def test_merge_metadata_changes_affect_table_alignment_signatures(self):
        engine = ExcelDiffEngine()
        report_input = ExcelReportInput(
            excel_save_path="unused.xlsx",
            tables_before=[
                [
                    [TableCellData(text="HEADER", grid_span=2), TableCellData(text="HEADER", grid_span=2)],
                    [TableCellData(text="A"), TableCellData(text="B")],
                ]
            ],
            tables_after=[
                [
                    [TableCellData(text="HEADER", grid_span=1), TableCellData(text="HEADER", grid_span=1)],
                    [TableCellData(text="A"), TableCellData(text="B")],
                ]
            ],
        )

        diff_plan = engine.build_diff_plan(report_input)
        row_opcodes = diff_plan.tables[0].row_opcodes

        self.assertIn(("replace", 0, 1, 0, 1), row_opcodes)


if __name__ == "__main__":
    unittest.main()
