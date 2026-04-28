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
                        grid_widths = self._extract_table_grid_widths(table)
                        for row in table.rows:
                            row_height = self._extract_row_height(row)
                            row_data = [
                                self._build_cell_data(
                                    cell,
                                    row_height=row_height,
                                    grid_col_width=grid_widths[cell_index] if cell_index < len(grid_widths) else 0,
                                )
                                for cell_index, cell in enumerate(row.cells)
                            ]
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
    def _build_cell_data(cell, row_height: int = 0, grid_col_width: int = 0) -> TableCellData:
        text = cell.text.replace("\r", "\n").strip()
        grid_span = 1
        v_merge = ""
        cell_width = 0

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

            tc_width_element = getattr(tc_pr, "tcW", None)
            cell_width = DocxExtractor._extract_xml_int_value(tc_width_element, "w")

        return TableCellData(
            text=text,
            grid_span=grid_span,
            v_merge=v_merge,
            cell_width=cell_width,
            row_height=row_height,
            grid_col_width=grid_col_width,
        )

    @staticmethod
    def _extract_table_grid_widths(table) -> list[int]:
        tbl_grid = getattr(table._tbl, "tblGrid", None)
        if tbl_grid is None:
            return []

        grid_cols = getattr(tbl_grid, "gridCol_lst", None) or []
        return [DocxExtractor._extract_xml_int_value(grid_col, "w") for grid_col in grid_cols]

    @staticmethod
    def _extract_row_height(row) -> int:
        tr_pr = getattr(row._tr, "trPr", None)
        if tr_pr is None:
            return 0

        tr_height = getattr(tr_pr, "trHeight", None)
        if isinstance(tr_height, list):
            tr_height = tr_height[0] if tr_height else None
        return DocxExtractor._extract_xml_int_value(tr_height, "val")

    @staticmethod
    def _extract_xml_int_value(element, attribute_name: str) -> int:
        if element is None:
            return 0

        try:
            value = getattr(element, attribute_name)
        except Exception:
            value = None

        if value is None:
            return 0

        try:
            return int(value)
        except (TypeError, ValueError):
            return 0
