import subprocess
import tempfile
import unittest
import zipfile
from pathlib import Path

from app.services.native_docx_extractor import NativeDocxExtractor


DOCUMENT_XML = """\
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:jc w:val="center"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:color w:val="FF0000"/>
        </w:rPr>
        <w:t>본문 제목</w:t>
      </w:r>
    </w:p>
    <w:tbl>
      <w:tblGrid>
        <w:gridCol w:w="2000"/>
      </w:tblGrid>
      <w:tr>
        <w:trPr>
          <w:trHeight w:val="320"/>
        </w:trPr>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="2200"/>
            <w:shd w:fill="DDDDDD"/>
          </w:tcPr>
          <w:p><w:r><w:t>셀값</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>
"""


class NativeExtractorBinaryTests(unittest.TestCase):
    def test_built_native_binary_parses_sample_docx(self):
        binary_path = NativeDocxExtractor().resolve_binary_path()
        if not binary_path:
            self.skipTest("native extractor binary is not available")

        with tempfile.TemporaryDirectory() as temp_dir:
            docx_path = Path(temp_dir) / "sample.docx"
            with zipfile.ZipFile(docx_path, "w") as archive:
                archive.writestr("word/document.xml", DOCUMENT_XML)

            completed = subprocess.run(
                [binary_path, str(docx_path)],
                check=True,
                capture_output=True,
                text=True,
                encoding="utf-8",
            )

            payload = __import__("json").loads(completed.stdout)

        self.assertEqual(payload["paragraphs"][0]["text"], "본문 제목")
        self.assertEqual(payload["paragraphs"][0]["style_name"], "Heading1")
        self.assertEqual(payload["paragraphs"][1]["text"], "[TABLE_MARKER]")
        self.assertEqual(payload["tables"][0][0][0]["text"], "셀값")
        self.assertEqual(payload["tables"][0][0][0]["row_height"], 320)
        self.assertEqual(payload["tables"][0][0][0]["grid_col_width"], 2000)


if __name__ == "__main__":
    unittest.main()
