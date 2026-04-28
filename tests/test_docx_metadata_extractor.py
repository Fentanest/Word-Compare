import tempfile
import unittest
import zipfile
from pathlib import Path
import sys
import types

try:
    from app.services.docx_extractor import DocxExtractor
except ModuleNotFoundError:
    docx_module = types.ModuleType("docx")
    docx_module.Document = object
    docx_table_module = types.ModuleType("docx.table")
    docx_table_module.Table = object
    docx_text_module = types.ModuleType("docx.text")
    docx_text_paragraph_module = types.ModuleType("docx.text.paragraph")
    docx_text_paragraph_module.Paragraph = object

    sys.modules.setdefault("docx", docx_module)
    sys.modules.setdefault("docx.table", docx_table_module)
    sys.modules.setdefault("docx.text", docx_text_module)
    sys.modules.setdefault("docx.text.paragraph", docx_text_paragraph_module)

    from app.services.docx_extractor import DocxExtractor


DOCUMENT_XML = """\
<w:document
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:v="urn:schemas-microsoft-com:vml">
  <w:body>
    <w:p><w:r><w:t>본문</w:t></w:r></w:p>
    <w:ins w:id="1" w:author="Alice" w:date="2026-04-29T09:00:00Z">
      <w:r><w:t>추가 텍스트</w:t></w:r>
    </w:ins>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <wp:docPr id="1" name="Flowchart" descr="diagram" title="Main diagram"/>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:pict>
          <v:shape id="TextBox1">
            <w:txbxContent>
              <w:p><w:r><w:t>텍스트박스 내용</w:t></w:r></w:p>
            </w:txbxContent>
          </v:shape>
        </w:pict>
      </w:r>
    </w:p>
    <w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="TableGrid"/>
        <w:tblW w:w="5000" w:type="dxa"/>
        <w:jc w:val="center"/>
        <w:tblLayout w:type="fixed"/>
        <w:tblLook w:val="04A0" w:firstRow="1" w:noVBand="1"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="8" w:color="auto"/>
        </w:tblBorders>
        <w:tblCellMar>
          <w:left w:w="100" w:type="dxa"/>
        </w:tblCellMar>
        <w:shd w:val="clear" w:fill="DDDDDD"/>
      </w:tblPr>
    </w:tbl>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838" w:orient="landscape"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      <w:cols w:num="2" w:space="720"/>
      <w:titlePg/>
    </w:sectPr>
  </w:body>
</w:document>
"""

HEADER_XML = """\
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>머리말 내용</w:t></w:r></w:p>
</w:hdr>
"""

FOOTER_XML = """\
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>꼬리말 내용</w:t></w:r></w:p>
</w:ftr>
"""

FOOTNOTES_XML = """\
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="1">
    <w:p><w:r><w:t>각주 내용</w:t></w:r></w:p>
  </w:footnote>
</w:footnotes>
"""

ENDNOTES_XML = """\
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:id="2">
    <w:p><w:r><w:t>미주 내용</w:t></w:r></w:p>
  </w:endnote>
</w:endnotes>
"""

COMMENTS_XML = """\
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Bob" w:initials="BB" w:date="2026-04-29T10:00:00Z">
    <w:p><w:r><w:t>주석 내용</w:t></w:r></w:p>
  </w:comment>
</w:comments>
"""


class DocxMetadataExtractorTests(unittest.TestCase):
    def test_extract_metadata_blocks_covers_remaining_document_structures(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            archive_path = Path(temp_dir) / "metadata.docx"
            with zipfile.ZipFile(archive_path, "w") as archive:
                archive.writestr("word/document.xml", DOCUMENT_XML)
                archive.writestr("word/header1.xml", HEADER_XML)
                archive.writestr("word/footer1.xml", FOOTER_XML)
                archive.writestr("word/footnotes.xml", FOOTNOTES_XML)
                archive.writestr("word/endnotes.xml", ENDNOTES_XML)
                archive.writestr("word/comments.xml", COMMENTS_XML)

            with zipfile.ZipFile(archive_path) as archive:
                blocks, locations = DocxExtractor._extract_metadata_blocks(archive)

        location_set = set(locations)
        self.assertTrue({"머리말 1", "꼬리말 1", "각주 1", "미주 2", "주석 0", "구역 1", "텍스트 상자 1", "도형 1", "표 1 서식"}.issubset(location_set))
        self.assertTrue(any(location.startswith("변경 추적 삽입") for location in locations))

        block_by_kind = {}
        for block in blocks:
            block_by_kind.setdefault(block.source_kind, []).append(block)

        self.assertIn("header", block_by_kind)
        self.assertIn("footer", block_by_kind)
        self.assertIn("footnote", block_by_kind)
        self.assertIn("endnote", block_by_kind)
        self.assertIn("comment", block_by_kind)
        self.assertIn("revision", block_by_kind)
        self.assertIn("section", block_by_kind)
        self.assertIn("textbox", block_by_kind)
        self.assertIn("shape", block_by_kind)
        self.assertIn("table-meta", block_by_kind)

        section_block = block_by_kind["section"][0]
        self.assertIn("pgSz:orient=landscape", section_block.extra_meta)
        self.assertIn("cols:num=2", section_block.extra_meta)
        self.assertIn("titlePg=true", section_block.extra_meta)

        table_block = block_by_kind["table-meta"][0]
        self.assertIn("tblStyle:val=TableGrid", table_block.extra_meta)
        self.assertIn("tblBorders:top:val=single", table_block.extra_meta)
        self.assertIn("tblShd:fill=DDDDDD", table_block.extra_meta)

        comment_block = block_by_kind["comment"][0]
        self.assertEqual(comment_block.text, "주석 내용")
        self.assertIn("author=Bob", comment_block.extra_meta)


if __name__ == "__main__":
    unittest.main()
