import os
import sys
from pathlib import Path

from docx import Document as DocxReader
from docx.table import Table as _Table
from docx.text.paragraph import Paragraph as _Paragraph

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from excel_generator import create_excel_report


def simulate_hybrid_extraction(file_path):
    reader = DocxReader(file_path)
    paragraphs = []
    table_flags = []
    tables = []

    for child in reader.element.body:
        if child.tag.endswith("p"):
            paragraph = _Paragraph(child, reader)
            paragraphs.append(paragraph.text.strip())
            table_flags.append(False)
        elif child.tag.endswith("tbl"):
            table = _Table(child, reader)
            paragraphs.append("[TABLE_MARKER]")
            table_flags.append(True)

            table_grid = []
            for row in table.rows:
                table_grid.append([cell.text.replace("\r", "\n").strip() for cell in row.cells])
            tables.append(table_grid)

    return paragraphs, table_flags, tables


def analyze_generated_data():
    print("--- 데이터 추출 및 시뮬레이션 시작 ---")
    if not os.path.exists("before.docx") or not os.path.exists("after.docx"):
        print("Error: before.docx or after.docx not found.")
        return

    paragraphs_before, flags_before, tables_before = simulate_hybrid_extraction("before.docx")
    paragraphs_after, flags_after, tables_after = simulate_hybrid_extraction("after.docx")

    print(f"수정 전 표 개수: {len(tables_before)}")
    print(f"수정 후 표 개수: {len(tables_after)}")

    if len(tables_after) >= 4:
        table_after = tables_after[3]
        print("\n--- [수정 후 표 4] 데이터 추출 결과 (상단 10행) ---")
        for row_index, row in enumerate(table_after[:10]):
            print(f"Row {row_index}: {row}")

    def dummy_log(message):
        print(message)

    create_excel_report(
        None,
        None,
        "simulated_report.xlsx",
        dummy_log,
        paragraphs_before,
        paragraphs_after,
        None,
        flags_before,
        flags_after,
        tables_before,
        tables_after,
    )

    try:
        import openpyxl

        workbook = openpyxl.load_workbook("simulated_report.xlsx")
        if "표 4" in workbook.sheetnames:
            worksheet = workbook["표 4"]
            print("\n--- [엑셀 결과물 분석] 표 4 시트 ---")
            print(f"H4 셀 값: '{worksheet['H4'].value}'")
            print(f"I4 셀 값: '{worksheet['I4'].value}'")
            print(f"H5 셀 값: '{worksheet['H5'].value}'")
            print(f"I5 셀 값: '{worksheet['I5'].value}'")
            print(f"J5 셀 값: '{worksheet['J5'].value}'")
    except Exception as error:
        print(f"엑셀 분석 중 오류: {error}")


if __name__ == "__main__":
    analyze_generated_data()
