import os
import tempfile

from docx import Document as DocxReader
from docx.table import Table as _Table
from docx.text.paragraph import Paragraph as _Paragraph

from app.models import ExtractedDocument, TableCellData


class DocxExtractor:
    def extract_data_hybrid(self, doc, log_callback=None, doc_name: str = "") -> ExtractedDocument:
        try:
            self._log(log_callback, f"-> '{doc_name}' 데이터 분석 및 고속 추출 준비 중...")

            # Word가 자동 번호를 실제 텍스트로 확정하도록 한 번 정리한다.
            doc.Content.ListFormat.ConvertNumbersToText()

            fd, temp_path = tempfile.mkstemp(suffix=".docx", prefix="extract_")
            os.close(fd)
            doc.SaveAs(os.path.abspath(temp_path), FileFormat=12)

            reader = DocxReader(temp_path)
            paragraphs: list[str] = []
            table_flags: list[bool] = []
            tables: list[list[list[str | TableCellData]]] = []

            for child in reader.element.body:
                if child.tag.endswith("p"):
                    paragraph = _Paragraph(child, reader)
                    paragraphs.append(paragraph.text.strip())
                    table_flags.append(False)
                elif child.tag.endswith("tbl"):
                    table = _Table(child, reader)
                    paragraphs.append("[TABLE_MARKER]")
                    table_flags.append(True)

                    table_grid: list[list[str | TableCellData]] = []
                    try:
                        for row in table.rows:
                            row_data = [self._build_cell_data(cell) for cell in row.cells]
                            table_grid.append(row_data)
                        tables.append(table_grid)
                    except Exception as table_error:
                        self._log(log_callback, f"-> 표 추출 중 오류: {table_error}")
                        tables.append([["[데이터 추출 실패]"]])

            try:
                os.remove(temp_path)
            except OSError:
                pass

            self._log(log_callback, f"-> '{doc_name}' 데이터 추출 완료 (표 {len(tables)}개 발견)")
            return ExtractedDocument(
                paragraphs=paragraphs,
                table_flags=table_flags,
                tables=tables,
            )
        except Exception as error:
            self._log(log_callback, f"-> 하이브리드 추출 오류: {error}")
            return ExtractedDocument(
                paragraphs=[paragraph.Range.Text for paragraph in doc.Paragraphs],
                table_flags=[False] * doc.Paragraphs.Count,
                tables=[],
            )

    @staticmethod
    def _log(log_callback, message: str) -> None:
        if log_callback:
            log_callback(message)

    @staticmethod
    def _build_cell_data(cell) -> TableCellData:
        text = cell.text.replace("\r", "\n").strip()
        grid_span = 1
        v_merge = ""

        tc_pr = getattr(cell._tc, "tcPr", None)
        if tc_pr is not None:
            grid_span_element = getattr(tc_pr, "gridSpan", None)
            if grid_span_element is not None:
                try:
                    grid_span = int(grid_span_element.val)
                except (TypeError, ValueError):
                    grid_span = 1

            v_merge_element = getattr(tc_pr, "vMerge", None)
            if v_merge_element is not None:
                v_merge_value = getattr(v_merge_element, "val", None)
                if v_merge_value is None:
                    v_merge = "continue"
                else:
                    v_merge = str(v_merge_value)

        return TableCellData(
            text=text,
            grid_span=grid_span,
            v_merge=v_merge,
        )
