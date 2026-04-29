import tempfile
import unittest
from pathlib import Path

from app.services.native_docx_extractor import NativeDocxExtractor


class _CompletedProcess:
    def __init__(self, stdout: str, stderr: str = ""):
        self.stdout = stdout
        self.stderr = stderr


class NativeDocxExtractorTests(unittest.TestCase):
    def test_payload_to_document_rehydrates_models(self):
        payload = {
            "paragraphs": [
                {
                    "type": "paragraph",
                    "text": "본문",
                    "style_name": "Heading1",
                    "alignment": "center",
                    "source_kind": "body",
                    "source_identifier": "",
                    "left_indent": 120,
                    "right_indent": 0,
                    "first_line_indent": 60,
                    "space_before": 100,
                    "space_after": 80,
                    "line_spacing": "240",
                    "keep_together": True,
                    "keep_with_next": False,
                    "page_break_before": False,
                    "widow_control": True,
                    "runs": [
                        {
                            "text": "본문",
                            "bold": True,
                            "italic": False,
                            "underline": "single",
                            "font_name": "Malgun Gothic",
                            "font_size": 22,
                            "color": "FF0000",
                            "highlight": "yellow",
                            "strike": False,
                            "style_name": "Strong",
                        }
                    ],
                    "extra_meta": [],
                },
                {"type": "marker", "text": "[TABLE_MARKER]"},
            ],
            "table_flags": [False, True],
            "tables": [
                [
                    [
                        {
                            "text": "A",
                            "grid_span": 2,
                            "v_merge": "",
                            "cell_width": 1200,
                            "row_height": 320,
                            "grid_col_width": 1100,
                            "border_signature": "top:val=single",
                            "shading_fill": "fill=DDDDDD",
                            "vertical_align": "center",
                            "text_direction": "lrTb",
                            "nested_table_count": 0,
                        }
                    ]
                ]
            ],
            "paragraph_locations": ["1행", "2행"],
        }

        extracted = NativeDocxExtractor._payload_to_document(payload)

        self.assertEqual(len(extracted.paragraphs), 2)
        self.assertEqual(extracted.paragraphs[0].text, "본문")
        self.assertEqual(extracted.paragraphs[0].runs[0].font_name, "Malgun Gothic")
        self.assertEqual(extracted.paragraphs[1], "[TABLE_MARKER]")
        self.assertEqual(extracted.tables[0][0][0].grid_span, 2)
        self.assertEqual(extracted.paragraph_locations, ["1행", "2행"])

    def test_extract_uses_configured_binary_and_parses_stdout(self):
        payload = {
            "paragraphs": [{"type": "marker", "text": "[TABLE_MARKER]"}],
            "table_flags": [True],
            "tables": [],
            "paragraph_locations": ["1행"],
        }

        calls = []

        def fake_run(command, capture_output, check, text, encoding):
            calls.append(command)
            return _CompletedProcess(stdout=__import__("json").dumps(payload))

        with tempfile.TemporaryDirectory() as temp_dir:
            binary_path = Path(temp_dir) / "word_compare_native_extractor.exe"
            binary_path.write_text("", encoding="utf-8")
            extractor = NativeDocxExtractor(binary_path=str(binary_path), run_command=fake_run)
            result = extractor.extract("sample.docx")

        self.assertEqual(calls[0][0], str(binary_path))
        self.assertEqual(result.paragraphs, ["[TABLE_MARKER]"])

    def test_extract_returns_none_when_binary_is_missing(self):
        extractor = NativeDocxExtractor(binary_path="missing-native-extractor")
        self.assertIsNone(extractor.extract("sample.docx"))


if __name__ == "__main__":
    unittest.main()
