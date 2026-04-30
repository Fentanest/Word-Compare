import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from app.models import ParagraphData, RunData
from app.reports.excel_report_service import ExcelReportService
from excel_generator import create_excel_report


NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
}


def extract_sheet_names(xlsx_path: Path) -> list[str]:
    with zipfile.ZipFile(xlsx_path) as archive:
        workbook_xml = archive.read("xl/workbook.xml")
    root = ET.fromstring(workbook_xml)
    return [sheet.attrib["name"] for sheet in root.findall(".//main:sheets/main:sheet", NS)]


def extract_shared_strings(xlsx_path: Path) -> list[str]:
    with zipfile.ZipFile(xlsx_path) as archive:
        if "xl/sharedStrings.xml" not in archive.namelist():
            return []
        xml_bytes = archive.read("xl/sharedStrings.xml")
    root = ET.fromstring(xml_bytes)
    values = []
    for string_item in root.findall(".//main:si", NS):
        text = "".join(node.text or "" for node in string_item.findall(".//main:t", NS))
        values.append(text)
    return values


def extract_sheet_values(xlsx_path: Path, sheet_name: str) -> list[str]:
    with zipfile.ZipFile(xlsx_path) as archive:
        workbook_xml = archive.read("xl/workbook.xml")
        workbook_root = ET.fromstring(workbook_xml)
        sheets = workbook_root.findall(".//main:sheets/main:sheet", NS)
        target_sheet_id = None
        for sheet in sheets:
            if sheet.attrib["name"] == sheet_name:
                target_sheet_id = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
                break

        if not target_sheet_id:
            return []

        rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        target_path = None
        for rel in rels_root:
            if rel.attrib.get("Id") == target_sheet_id:
                target_path = f"xl/{rel.attrib['Target']}"
                break

        if not target_path:
            return []

        shared_strings = extract_shared_strings(xlsx_path)
        sheet_root = ET.fromstring(archive.read(target_path))

    values = []
    for cell in sheet_root.findall(".//main:c", NS):
        cell_type = cell.attrib.get("t")
        if cell_type == "s":
            value_index = cell.find("main:v", NS)
            if value_index is not None:
                values.append(shared_strings[int(value_index.text)])
        elif cell_type == "inlineStr":
            values.append("".join(node.text or "" for node in cell.findall(".//main:t", NS)))
        else:
            value = cell.find("main:v", NS)
            if value is not None:
                values.append(value.text or "")
    return values


def extract_sheet_cells(xlsx_path: Path, sheet_name: str) -> dict[str, dict[str, str]]:
    with zipfile.ZipFile(xlsx_path) as archive:
        workbook_xml = archive.read("xl/workbook.xml")
        workbook_root = ET.fromstring(workbook_xml)
        sheets = workbook_root.findall(".//main:sheets/main:sheet", NS)
        target_sheet_id = None
        for sheet in sheets:
            if sheet.attrib["name"] == sheet_name:
                target_sheet_id = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
                break

        if not target_sheet_id:
            return {}

        rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        target_path = None
        for rel in rels_root:
            if rel.attrib.get("Id") == target_sheet_id:
                target_path = f"xl/{rel.attrib['Target']}"
                break

        if not target_path:
            return {}

        sheet_root = ET.fromstring(archive.read(target_path))

    return {cell.attrib["r"]: dict(cell.attrib) for cell in sheet_root.findall(".//main:c", NS)}


class ExcelReportPipelineTests(unittest.TestCase):
    def test_excel_report_service_writes_main_sheet_and_changed_text(self):
        service = ExcelReportService(extractor=None)

        with tempfile.TemporaryDirectory() as temp_dir:
            xlsx_path = Path(temp_dir) / "report.xlsx"
            service.generate_from_extracted_data(
                excel_save_path=str(xlsx_path),
                log_callback=None,
                paras_before=["same line", "before text"],
                paras_after=["same line", "after text"],
                flags_b=[False, False],
                flags_a=[False, False],
                tables_before=[],
                tables_after=[],
                get_loc_cb=lambda idx, is_before: f"{idx + 1}행",
            )

            self.assertTrue(xlsx_path.exists())
            self.assertIn("변경 내용(일반)", extract_sheet_names(xlsx_path))

            strings = extract_shared_strings(xlsx_path)
            self.assertIn("위치", strings)
            self.assertIn("수정 전", strings)
            self.assertIn("수정 후", strings)
            self.assertIn("서식 변경", strings)
            self.assertIn("2행", strings)
            self.assertIn("before text", strings)
            self.assertIn("after text", strings)

    def test_legacy_excel_generator_wrapper_keeps_table_sheet_output(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            xlsx_path = Path(temp_dir) / "table-report.xlsx"
            create_excel_report(
                None,
                None,
                str(xlsx_path),
                None,
                paras_before=[],
                paras_after=[],
                get_loc_cb=None,
                flags_b=[],
                flags_a=[],
                tables_before=[[["A", "B"], ["C", "D"]]],
                tables_after=[[["A", "B2"], ["C", "D"]]],
            )

            self.assertTrue(xlsx_path.exists())
            self.assertIn("표 1", extract_sheet_names(xlsx_path))

            strings = extract_shared_strings(xlsx_path)
            self.assertIn("수정 전", strings)
            self.assertIn("수정 후", strings)
            self.assertIn("B", strings)
            self.assertIn("B2", strings)

    def test_format_only_paragraph_changes_are_marked_in_main_sheet(self):
        service = ExcelReportService(extractor=None)

        with tempfile.TemporaryDirectory() as temp_dir:
            xlsx_path = Path(temp_dir) / "format-report.xlsx"
            service.generate_from_extracted_data(
                excel_save_path=str(xlsx_path),
                log_callback=None,
                paras_before=[
                    ParagraphData(
                        text="같은 문장",
                        runs=(RunData(text="같은 문장", bold=False),),
                    )
                ],
                paras_after=[
                    ParagraphData(
                        text="같은 문장",
                        runs=(RunData(text="같은 문장", bold=True),),
                    )
                ],
                flags_b=[False],
                flags_a=[False],
                tables_before=[],
                tables_after=[],
                get_loc_cb=lambda idx, is_before: f"{idx + 1}행",
            )

            strings = extract_shared_strings(xlsx_path)
            self.assertIn("같은 문장", strings)
            self.assertNotIn("같은 문장 [서식 변경]", strings)

            main_values = extract_sheet_values(xlsx_path, "변경 내용(일반)")
            self.assertIn("서식 변경", main_values)
            self.assertIn("O", main_values)

            cells = extract_sheet_cells(xlsx_path, "변경 내용(일반)")
            self.assertEqual(cells["B2"].get("s"), cells["C2"].get("s"))

    def test_metadata_only_changes_are_marked_in_main_sheet(self):
        service = ExcelReportService(extractor=None)

        with tempfile.TemporaryDirectory() as temp_dir:
            xlsx_path = Path(temp_dir) / "metadata-report.xlsx"
            service.generate_from_extracted_data(
                excel_save_path=str(xlsx_path),
                log_callback=None,
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
                tables_before=[],
                tables_after=[],
                get_loc_cb=lambda idx, is_before: "구역 1",
            )

            strings = extract_shared_strings(xlsx_path)
            self.assertIn("구역 1 설정 [구조/메타데이터 변경]", strings)


if __name__ == "__main__":
    unittest.main()
