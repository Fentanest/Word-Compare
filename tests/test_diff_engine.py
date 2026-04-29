import unittest

from app.models import ParagraphData, RunData, TableCellData
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

    def test_row_shift_alignment_tolerates_numeric_deltas_when_text_identity_matches(self):
        engine = ExcelDiffEngine()
        report_input = ExcelReportInput(
            excel_save_path="unused.xlsx",
            tables_before=[
                [
                    ["주주명", "주식수", "지분율"],
                    ["㈜아이니즈", "10,173", "0.49"],
                    ["기타 개인투자자", "382,499", "18.26"],
                    ["합  계", "392,672", "18.75"],
                ]
            ],
            tables_after=[
                [
                    ["주주명", "주식수", "지분율"],
                    ["패스웨이인사이트투자조합22호", "42,940", "2.01"],
                    ["㈜아이니즈", "10,173", "0.48"],
                    ["기타 개인투자자", "382,499", "17.88"],
                    ["합  계", "435,612", "20.37"],
                ]
            ],
        )

        diff_plan = engine.build_diff_plan(report_input)
        row_opcodes = diff_plan.tables[0].row_opcodes

        self.assertIn(("insert", 1, 1, 1, 2), row_opcodes)
        self.assertIn(("equal", 1, 4, 2, 5), row_opcodes)

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

    def test_paragraph_run_format_changes_affect_main_diff_signatures(self):
        engine = ExcelDiffEngine()
        report_input = ExcelReportInput(
            excel_save_path="unused.xlsx",
            paras_before=[
                ParagraphData(
                    text="같은 본문",
                    runs=(RunData(text="같은 본문", bold=False),),
                )
            ],
            paras_after=[
                ParagraphData(
                    text="같은 본문",
                    runs=(RunData(text="같은 본문", bold=True),),
                )
            ],
            flags_b=[False],
            flags_a=[False],
        )

        diff_plan = engine.build_diff_plan(report_input)

        self.assertEqual(diff_plan.main_opcodes, [("replace", 0, 1, 0, 1)])

    def test_paragraph_style_changes_affect_main_diff_signatures(self):
        engine = ExcelDiffEngine()
        report_input = ExcelReportInput(
            excel_save_path="unused.xlsx",
            paras_before=[ParagraphData(text="제목", style_name="Normal")],
            paras_after=[ParagraphData(text="제목", style_name="Heading 1")],
            flags_b=[False],
            flags_a=[False],
        )

        diff_plan = engine.build_diff_plan(report_input)

        self.assertEqual(diff_plan.main_opcodes, [("replace", 0, 1, 0, 1)])

    def test_metadata_block_changes_affect_main_diff_signatures(self):
        engine = ExcelDiffEngine()
        report_input = ExcelReportInput(
            excel_save_path="unused.xlsx",
            paras_before=[
                ParagraphData(
                    text="구역 1 설정",
                    source_kind="section",
                    source_identifier="1",
                    extra_meta=("pgSz:orient=portrait",),
                )
            ],
            paras_after=[
                ParagraphData(
                    text="구역 1 설정",
                    source_kind="section",
                    source_identifier="1",
                    extra_meta=("pgSz:orient=landscape",),
                )
            ],
            flags_b=[False],
            flags_a=[False],
        )

        diff_plan = engine.build_diff_plan(report_input)

        self.assertEqual(diff_plan.main_opcodes, [("replace", 0, 1, 0, 1)])

    def test_cell_width_changes_affect_table_alignment_signatures(self):
        engine = ExcelDiffEngine()
        report_input = ExcelReportInput(
            excel_save_path="unused.xlsx",
            tables_before=[
                [
                    [TableCellData(text="HEADER", cell_width=2400), TableCellData(text="VALUE", cell_width=1800)],
                    [TableCellData(text="A", cell_width=2400), TableCellData(text="B", cell_width=1800)],
                ]
            ],
            tables_after=[
                [
                    [TableCellData(text="HEADER", cell_width=3200), TableCellData(text="VALUE", cell_width=1800)],
                    [TableCellData(text="A", cell_width=3200), TableCellData(text="B", cell_width=1800)],
                ]
            ],
        )

        diff_plan = engine.build_diff_plan(report_input)
        col_opcodes = diff_plan.tables[0].col_opcodes

        self.assertIn(("replace", 0, 1, 0, 1), col_opcodes)

    def test_row_height_changes_affect_table_alignment_signatures(self):
        engine = ExcelDiffEngine()
        report_input = ExcelReportInput(
            excel_save_path="unused.xlsx",
            tables_before=[
                [
                    [TableCellData(text="HEADER", row_height=320)],
                    [TableCellData(text="BODY", row_height=480)],
                ]
            ],
            tables_after=[
                [
                    [TableCellData(text="HEADER", row_height=320)],
                    [TableCellData(text="BODY", row_height=720)],
                ]
            ],
        )

        diff_plan = engine.build_diff_plan(report_input)
        row_opcodes = diff_plan.tables[0].row_opcodes

        self.assertIn(("replace", 1, 2, 1, 2), row_opcodes)

    def test_table_grid_width_changes_affect_table_alignment_signatures(self):
        engine = ExcelDiffEngine()
        report_input = ExcelReportInput(
            excel_save_path="unused.xlsx",
            tables_before=[
                [
                    [
                        TableCellData(text="ROW1-A", grid_col_width=1400),
                        TableCellData(text="ROW1-B", grid_col_width=2200),
                    ],
                    [
                        TableCellData(text="ROW2-A", grid_col_width=1400),
                        TableCellData(text="ROW2-B", grid_col_width=2200),
                    ],
                ]
            ],
            tables_after=[
                [
                    [
                        TableCellData(text="ROW1-A", grid_col_width=1800),
                        TableCellData(text="ROW1-B", grid_col_width=2200),
                    ],
                    [
                        TableCellData(text="ROW2-A", grid_col_width=1800),
                        TableCellData(text="ROW2-B", grid_col_width=2200),
                    ],
                ]
            ],
        )

        diff_plan = engine.build_diff_plan(report_input)
        col_opcodes = diff_plan.tables[0].col_opcodes

        self.assertIn(("replace", 0, 1, 0, 1), col_opcodes)

    def test_table_cell_style_changes_affect_table_alignment_signatures(self):
        engine = ExcelDiffEngine()
        report_input = ExcelReportInput(
            excel_save_path="unused.xlsx",
            tables_before=[
                [
                    [
                        TableCellData(
                            text="A",
                            border_signature="top:val=single,sz=8,color=auto",
                            shading_fill="fill=FFFFFF",
                            vertical_align="center",
                            text_direction="lrTb",
                            nested_table_count=0,
                        )
                    ]
                ]
            ],
            tables_after=[
                [
                    [
                        TableCellData(
                            text="A",
                            border_signature="top:val=double,sz=8,color=auto",
                            shading_fill="fill=FFFF00",
                            vertical_align="bottom",
                            text_direction="tbRl",
                            nested_table_count=1,
                        )
                    ]
                ]
            ],
        )

        diff_plan = engine.build_diff_plan(report_input)
        row_opcodes = diff_plan.tables[0].row_opcodes

        self.assertIn(("replace", 0, 1, 0, 1), row_opcodes)


if __name__ == "__main__":
    unittest.main()
