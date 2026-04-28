import os
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from time import perf_counter

from docx import Document as DocxReader
from docx.table import Table as _Table
from docx.text.paragraph import Paragraph as _Paragraph

from app.models import ExtractedDocument, ParagraphData, RunData, TableCellData


class DocxExtractor:
    WORD_NS = {
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
        "v": "urn:schemas-microsoft-com:vml",
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    }
    WORD_TAG = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

    def extract_data_hybrid(self, doc, log_callback=None, doc_name: str = "") -> ExtractedDocument:
        try:
            total_started_at = perf_counter()
            self._log(log_callback, f"-> '{doc_name}' 데이터 분석 및 고속 추출 준비 중...")

            # Word가 자동 번호를 실제 텍스트로 확정하도록 한 번 정리한다.
            convert_started_at = perf_counter()
            doc.Content.ListFormat.ConvertNumbersToText()
            self._log_perf(log_callback, f"{doc_name} 번호 텍스트화", convert_started_at)

            save_started_at = perf_counter()
            fd, temp_path = tempfile.mkstemp(suffix=".docx", prefix="extract_")
            os.close(fd)
            doc.SaveAs(os.path.abspath(temp_path), FileFormat=12)
            self._log_perf(log_callback, f"{doc_name} 임시 DOCX 저장", save_started_at)

            reader_started_at = perf_counter()
            reader = DocxReader(temp_path)
            self._log_perf(log_callback, f"{doc_name} python-docx 로드", reader_started_at)
            paragraphs: list[str | ParagraphData] = []
            table_flags: list[bool] = []
            tables: list[list[list[str | TableCellData]]] = []
            paragraph_locations: list[str] = []

            body_parse_started_at = perf_counter()
            for child in reader.element.body:
                if child.tag.endswith("p"):
                    paragraph = _Paragraph(child, reader)
                    paragraphs.append(self._build_paragraph_data(paragraph))
                    table_flags.append(False)
                    paragraph_locations.append(f"{len(paragraph_locations) + 1}행")
                elif child.tag.endswith("tbl"):
                    table = _Table(child, reader)
                    paragraphs.append("[TABLE_MARKER]")
                    table_flags.append(True)
                    paragraph_locations.append(f"{len(paragraph_locations) + 1}행")

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
            self._log_perf(log_callback, f"{doc_name} 본문/표 파싱", body_parse_started_at)

            metadata_started_at = perf_counter()
            with zipfile.ZipFile(temp_path) as archive:
                metadata_blocks, metadata_locations = self._extract_metadata_blocks(archive)
                paragraphs.extend(metadata_blocks)
                table_flags.extend([False] * len(metadata_blocks))
                paragraph_locations.extend(metadata_locations)
            self._log_perf(log_callback, f"{doc_name} XML 메타데이터 파싱", metadata_started_at)

            try:
                os.remove(temp_path)
            except OSError:
                pass

            self._log(
                log_callback,
                f"-> '{doc_name}' 데이터 추출 완료 (표 {len(tables)}개, 메타데이터 {len(metadata_blocks)}개 발견)",
            )
            self._log_perf(log_callback, f"{doc_name} 추출 전체", total_started_at)
            return ExtractedDocument(
                paragraphs=paragraphs,
                table_flags=table_flags,
                tables=tables,
                paragraph_locations=paragraph_locations,
            )
        except Exception as error:
            self._log(log_callback, f"-> 하이브리드 추출 오류: {error}")
            return ExtractedDocument(
                paragraphs=[paragraph.Range.Text for paragraph in doc.Paragraphs],
                table_flags=[False] * doc.Paragraphs.Count,
                tables=[],
                paragraph_locations=[f"{index + 1}행" for index in range(doc.Paragraphs.Count)],
            )

    @staticmethod
    def _log(log_callback, message: str) -> None:
        if log_callback:
            log_callback(message)

    @staticmethod
    def _log_perf(log_callback, label: str, started_at: float) -> None:
        DocxExtractor._log(log_callback, f"[성능] {label}: {perf_counter() - started_at:.3f}초")

    @staticmethod
    def _build_cell_data(cell, row_height: int = 0, grid_col_width: int = 0) -> TableCellData:
        text = cell.text.replace("\r", "\n").strip()
        grid_span = 1
        v_merge = ""
        cell_width = 0
        border_signature = ""
        shading_fill = ""
        vertical_align = ""
        text_direction = ""
        nested_table_count = len(getattr(cell, "tables", []) or [])

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
            border_signature = DocxExtractor._extract_border_signature(getattr(tc_pr, "tcBorders", None))
            shading_fill = DocxExtractor._extract_shading_fill(getattr(tc_pr, "shd", None))
            vertical_align = DocxExtractor._extract_string_value(getattr(tc_pr, "vAlign", None), "val")
            text_direction = DocxExtractor._extract_string_value(getattr(tc_pr, "textDirection", None), "val")

        return TableCellData(
            text=text,
            grid_span=grid_span,
            v_merge=v_merge,
            cell_width=cell_width,
            row_height=row_height,
            grid_col_width=grid_col_width,
            border_signature=border_signature,
            shading_fill=shading_fill,
            vertical_align=vertical_align,
            text_direction=text_direction,
            nested_table_count=nested_table_count,
        )

    @staticmethod
    def _build_paragraph_data(paragraph) -> ParagraphData:
        paragraph_format = paragraph.paragraph_format
        runs = tuple(DocxExtractor._build_run_data(run) for run in paragraph.runs if run.text)

        return ParagraphData(
            text=paragraph.text.strip(),
            style_name=DocxExtractor._safe_name(getattr(paragraph, "style", None)),
            alignment=DocxExtractor._normalize_enum(paragraph.alignment),
            left_indent=DocxExtractor._normalize_length(paragraph_format.left_indent),
            right_indent=DocxExtractor._normalize_length(paragraph_format.right_indent),
            first_line_indent=DocxExtractor._normalize_length(paragraph_format.first_line_indent),
            space_before=DocxExtractor._normalize_length(paragraph_format.space_before),
            space_after=DocxExtractor._normalize_length(paragraph_format.space_after),
            line_spacing=DocxExtractor._normalize_scalar(paragraph_format.line_spacing),
            keep_together=bool(paragraph_format.keep_together),
            keep_with_next=bool(paragraph_format.keep_with_next),
            page_break_before=bool(paragraph_format.page_break_before),
            widow_control=bool(paragraph_format.widow_control),
            runs=runs,
        )

    @staticmethod
    def _build_run_data(run) -> RunData:
        font = run.font
        color = ""
        try:
            color = str(font.color.rgb) if getattr(font.color, "rgb", None) else ""
        except Exception:
            color = ""

        return RunData(
            text=run.text.replace("\r", "\n"),
            bold=bool(font.bold),
            italic=bool(font.italic),
            underline=DocxExtractor._normalize_scalar(font.underline),
            font_name=font.name or "",
            font_size=DocxExtractor._normalize_length(font.size),
            color=color,
            highlight=DocxExtractor._normalize_scalar(font.highlight_color),
            strike=bool(font.strike),
            style_name=DocxExtractor._safe_name(getattr(run, "style", None)),
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

    @staticmethod
    def _extract_string_value(element, attribute_name: str) -> str:
        if element is None:
            return ""
        try:
            value = getattr(element, attribute_name)
        except Exception:
            value = None
        return "" if value is None else str(value)

    @staticmethod
    def _extract_shading_fill(element) -> str:
        if element is None:
            return ""

        parts = []
        for attribute_name in ("val", "color", "fill"):
            value = DocxExtractor._extract_string_value(element, attribute_name)
            if value:
                parts.append(f"{attribute_name}={value}")
        return "|".join(parts)

    @staticmethod
    def _extract_border_signature(borders) -> str:
        if borders is None:
            return ""

        parts = []
        for edge_name in ("top", "left", "bottom", "right", "insideH", "insideV", "tl2br", "tr2bl"):
            edge = getattr(borders, edge_name, None)
            if edge is None:
                continue

            edge_bits = []
            for attribute_name in ("val", "sz", "space", "color"):
                value = DocxExtractor._extract_string_value(edge, attribute_name)
                if value:
                    edge_bits.append(f"{attribute_name}={value}")
            if edge_bits:
                parts.append(f"{edge_name}:" + ",".join(edge_bits))
        return "|".join(parts)

    @classmethod
    def _extract_metadata_blocks(cls, archive) -> tuple[list[ParagraphData], list[str]]:
        blocks: list[ParagraphData] = []
        locations: list[str] = []

        cls._append_header_footer_blocks(archive, blocks, locations, "header", "머리말")
        cls._append_header_footer_blocks(archive, blocks, locations, "footer", "꼬리말")
        cls._append_note_blocks(archive, blocks, locations, "word/footnotes.xml", "footnote", "각주")
        cls._append_note_blocks(archive, blocks, locations, "word/endnotes.xml", "endnote", "미주")
        cls._append_comment_blocks(archive, blocks, locations)
        cls._append_revision_blocks(archive, blocks, locations)
        cls._append_section_blocks(archive, blocks, locations)
        cls._append_shape_blocks(archive, blocks, locations)
        cls._append_table_metadata_blocks(archive, blocks, locations)
        return blocks, locations

    @classmethod
    def _append_header_footer_blocks(cls, archive, blocks, locations, part_prefix: str, label_prefix: str) -> None:
        part_names = sorted(
            name
            for name in archive.namelist()
            if name.startswith(f"word/{part_prefix}") and name.endswith(".xml")
        )
        for index, part_name in enumerate(part_names, start=1):
            root = cls._load_xml_root(archive, part_name)
            if root is None:
                continue

            text = "\n".join(filter(None, cls._extract_paragraph_texts(root))).strip()
            cls._append_metadata_block(
                blocks,
                locations,
                label=f"{label_prefix} {index}",
                text=text or f"{label_prefix} {index}",
                kind=part_prefix,
                identifier=part_name,
            )

    @classmethod
    def _append_note_blocks(cls, archive, blocks, locations, part_name: str, tag_name: str, label_prefix: str) -> None:
        root = cls._load_xml_root(archive, part_name)
        if root is None:
            return

        for note in root.findall(f".//w:{tag_name}", cls.WORD_NS):
            note_type = cls._get_w_attr(note, "type")
            note_id = cls._get_w_attr(note, "id") or "0"
            try:
                if int(note_id) < 0:
                    continue
            except ValueError:
                pass
            if note_type:
                continue

            text = "\n".join(filter(None, cls._extract_paragraph_texts(note))).strip()
            cls._append_metadata_block(
                blocks,
                locations,
                label=f"{label_prefix} {note_id}",
                text=text or f"{label_prefix} {note_id}",
                kind=tag_name,
                identifier=note_id,
                extra_meta=(f"part={part_name}",),
            )

    @classmethod
    def _append_comment_blocks(cls, archive, blocks, locations) -> None:
        root = cls._load_xml_root(archive, "word/comments.xml")
        if root is None:
            return

        for comment in root.findall(".//w:comment", cls.WORD_NS):
            comment_id = cls._get_w_attr(comment, "id") or "0"
            text = "\n".join(filter(None, cls._extract_paragraph_texts(comment))).strip()
            cls._append_metadata_block(
                blocks,
                locations,
                label=f"주석 {comment_id}",
                text=text or f"주석 {comment_id}",
                kind="comment",
                identifier=comment_id,
                extra_meta=(
                    f"author={cls._get_w_attr(comment, 'author')}",
                    f"initials={cls._get_w_attr(comment, 'initials')}",
                    f"date={cls._get_w_attr(comment, 'date')}",
                ),
            )

    @classmethod
    def _append_revision_blocks(cls, archive, blocks, locations) -> None:
        tag_labels = {
            "ins": "삽입",
            "del": "삭제",
            "moveFrom": "이동-출발",
            "moveTo": "이동-도착",
        }
        part_names = [
            name
            for name in archive.namelist()
            if name == "word/document.xml"
            or name.startswith("word/header")
            or name.startswith("word/footer")
            or name.endswith("/footnotes.xml")
            or name.endswith("/endnotes.xml")
        ]
        counters = {tag_name: 0 for tag_name in tag_labels}

        for part_name in sorted(part_names):
            root = cls._load_xml_root(archive, part_name)
            if root is None:
                continue

            for tag_name, tag_label in tag_labels.items():
                for revision in root.findall(f".//w:{tag_name}", cls.WORD_NS):
                    counters[tag_name] += 1
                    text = cls._extract_text_from_element(revision)
                    cls._append_metadata_block(
                        blocks,
                        locations,
                        label=f"변경 추적 {tag_label} {counters[tag_name]}",
                        text=text or f"변경 추적 {tag_label}",
                        kind="revision",
                        identifier=f"{tag_name}:{counters[tag_name]}",
                        extra_meta=(
                            f"part={part_name}",
                            f"id={cls._get_w_attr(revision, 'id')}",
                            f"author={cls._get_w_attr(revision, 'author')}",
                            f"date={cls._get_w_attr(revision, 'date')}",
                            f"type={tag_name}",
                        ),
                    )

    @classmethod
    def _append_section_blocks(cls, archive, blocks, locations) -> None:
        root = cls._load_xml_root(archive, "word/document.xml")
        if root is None:
            return

        for index, sect_pr in enumerate(root.findall(".//w:sectPr", cls.WORD_NS), start=1):
            tokens = cls._extract_section_tokens(sect_pr)
            cls._append_metadata_block(
                blocks,
                locations,
                label=f"구역 {index}",
                text=f"구역 {index} 설정",
                kind="section",
                identifier=str(index),
                extra_meta=tokens,
            )

    @classmethod
    def _append_shape_blocks(cls, archive, blocks, locations) -> None:
        part_names = [
            name
            for name in archive.namelist()
            if name == "word/document.xml"
            or name.startswith("word/header")
            or name.startswith("word/footer")
        ]
        textbox_counter = 0
        shape_counter = 0

        for part_name in sorted(part_names):
            root = cls._load_xml_root(archive, part_name)
            if root is None:
                continue

            for textbox in root.findall(".//w:txbxContent", cls.WORD_NS):
                textbox_counter += 1
                text = "\n".join(filter(None, cls._extract_paragraph_texts(textbox))).strip()
                cls._append_metadata_block(
                    blocks,
                    locations,
                    label=f"텍스트 상자 {textbox_counter}",
                    text=text or f"텍스트 상자 {textbox_counter}",
                    kind="textbox",
                    identifier=f"{part_name}:{textbox_counter}",
                    extra_meta=(f"part={part_name}",),
                )

            for doc_pr in root.findall(".//wp:docPr", cls.WORD_NS):
                shape_counter += 1
                name = doc_pr.attrib.get("name", "")
                descr = doc_pr.attrib.get("descr", "")
                title = doc_pr.attrib.get("title", "")
                cls._append_metadata_block(
                    blocks,
                    locations,
                    label=f"도형 {shape_counter}",
                    text=name or descr or title or f"도형 {shape_counter}",
                    kind="shape",
                    identifier=f"{part_name}:{shape_counter}",
                    extra_meta=(
                        f"part={part_name}",
                        f"name={name}",
                        f"descr={descr}",
                        f"title={title}",
                    ),
                )

    @classmethod
    def _append_table_metadata_blocks(cls, archive, blocks, locations) -> None:
        root = cls._load_xml_root(archive, "word/document.xml")
        if root is None:
            return

        for index, table in enumerate(root.findall(".//w:tbl", cls.WORD_NS), start=1):
            tokens = cls._extract_table_tokens(table)
            cls._append_metadata_block(
                blocks,
                locations,
                label=f"표 {index} 서식",
                text=f"표 {index} 서식",
                kind="table-meta",
                identifier=str(index),
                extra_meta=tokens,
            )

    @classmethod
    def _append_metadata_block(cls, blocks, locations, label: str, text: str, kind: str, identifier: str, extra_meta=()) -> None:
        cleaned_text = text.strip()
        cleaned_meta = tuple(token for token in extra_meta if token and token != "=")
        blocks.append(
            ParagraphData(
                text=cleaned_text,
                source_kind=kind,
                source_identifier=identifier,
                extra_meta=cleaned_meta,
            )
        )
        locations.append(label)

    @classmethod
    def _load_xml_root(cls, archive, part_name: str):
        try:
            xml_bytes = archive.read(part_name)
        except KeyError:
            return None
        try:
            return ET.fromstring(xml_bytes)
        except ET.ParseError:
            return None

    @classmethod
    def _extract_paragraph_texts(cls, element) -> list[str]:
        paragraphs = []
        for paragraph in element.findall(".//w:p", cls.WORD_NS):
            text = cls._extract_text_from_element(paragraph)
            if text:
                paragraphs.append(text)

        if paragraphs:
            return paragraphs

        fallback = cls._extract_text_from_element(element)
        return [fallback] if fallback else []

    @classmethod
    def _extract_text_from_element(cls, element) -> str:
        parts = []
        for node in element.iter():
            if node.tag in {
                cls.WORD_TAG + "t",
                cls.WORD_TAG + "delText",
            }:
                text = node.text or ""
                if text:
                    parts.append(text)
            elif node.tag == cls.WORD_TAG + "tab":
                parts.append("\t")
            elif node.tag in {cls.WORD_TAG + "br", cls.WORD_TAG + "cr"}:
                parts.append("\n")
        return "".join(parts).strip()

    @classmethod
    def _get_w_attr(cls, element, attribute_name: str) -> str:
        return element.attrib.get(cls.WORD_TAG + attribute_name, "")

    @classmethod
    def _extract_section_tokens(cls, sect_pr) -> tuple[str, ...]:
        tokens: list[str] = []
        pg_sz = sect_pr.find("w:pgSz", cls.WORD_NS)
        if pg_sz is not None:
            tokens.extend(
                filter(
                    None,
                    (
                        f"pgSz:w={cls._get_w_attr(pg_sz, 'w')}",
                        f"pgSz:h={cls._get_w_attr(pg_sz, 'h')}",
                        f"pgSz:orient={cls._get_w_attr(pg_sz, 'orient')}",
                    ),
                )
            )

        pg_mar = sect_pr.find("w:pgMar", cls.WORD_NS)
        if pg_mar is not None:
            for edge_name in ("top", "right", "bottom", "left", "header", "footer", "gutter"):
                edge_value = cls._get_w_attr(pg_mar, edge_name)
                if edge_value:
                    tokens.append(f"pgMar:{edge_name}={edge_value}")

        cols = sect_pr.find("w:cols", cls.WORD_NS)
        if cols is not None:
            for attribute_name in ("num", "space", "sep", "equalWidth"):
                attribute_value = cls._get_w_attr(cols, attribute_name)
                if attribute_value:
                    tokens.append(f"cols:{attribute_name}={attribute_value}")

        pg_num_type = sect_pr.find("w:pgNumType", cls.WORD_NS)
        if pg_num_type is not None:
            for attribute_name in ("start", "fmt", "chapStyle"):
                attribute_value = cls._get_w_attr(pg_num_type, attribute_name)
                if attribute_value:
                    tokens.append(f"pgNumType:{attribute_name}={attribute_value}")

        doc_grid = sect_pr.find("w:docGrid", cls.WORD_NS)
        if doc_grid is not None:
            for attribute_name in ("type", "linePitch", "charSpace"):
                attribute_value = cls._get_w_attr(doc_grid, attribute_name)
                if attribute_value:
                    tokens.append(f"docGrid:{attribute_name}={attribute_value}")

        if sect_pr.find("w:titlePg", cls.WORD_NS) is not None:
            tokens.append("titlePg=true")
        return tuple(tokens)

    @classmethod
    def _extract_table_tokens(cls, table) -> tuple[str, ...]:
        tokens: list[str] = []
        table_pr = table.find("w:tblPr", cls.WORD_NS)
        if table_pr is None:
            return ()

        for tag_name in ("tblStyle", "tblW", "jc", "tblLayout"):
            element = table_pr.find(f"w:{tag_name}", cls.WORD_NS)
            if element is None:
                continue
            for attribute_name in ("val", "w", "type"):
                attribute_value = cls._get_w_attr(element, attribute_name)
                if attribute_value:
                    tokens.append(f"{tag_name}:{attribute_name}={attribute_value}")

        tbl_look = table_pr.find("w:tblLook", cls.WORD_NS)
        if tbl_look is not None:
            for attribute_name in ("val", "firstRow", "lastRow", "firstColumn", "lastColumn", "noHBand", "noVBand"):
                attribute_value = cls._get_w_attr(tbl_look, attribute_name)
                if attribute_value:
                    tokens.append(f"tblLook:{attribute_name}={attribute_value}")

        tokens.extend(cls._extract_border_tokens(table_pr.find("w:tblBorders", cls.WORD_NS), "tblBorders"))
        tokens.extend(cls._extract_margin_tokens(table_pr.find("w:tblCellMar", cls.WORD_NS), "tblCellMar"))

        shading = table_pr.find("w:shd", cls.WORD_NS)
        if shading is not None:
            for attribute_name in ("val", "color", "fill"):
                attribute_value = cls._get_w_attr(shading, attribute_name)
                if attribute_value:
                    tokens.append(f"tblShd:{attribute_name}={attribute_value}")

        return tuple(tokens)

    @classmethod
    def _extract_border_tokens(cls, borders, prefix: str) -> list[str]:
        if borders is None:
            return []

        tokens = []
        for edge_name in ("top", "left", "bottom", "right", "insideH", "insideV"):
            edge = borders.find(f"w:{edge_name}", cls.WORD_NS)
            if edge is None:
                continue
            for attribute_name in ("val", "sz", "space", "color"):
                attribute_value = cls._get_w_attr(edge, attribute_name)
                if attribute_value:
                    tokens.append(f"{prefix}:{edge_name}:{attribute_name}={attribute_value}")
        return tokens

    @classmethod
    def _extract_margin_tokens(cls, margins, prefix: str) -> list[str]:
        if margins is None:
            return []

        tokens = []
        for edge_name in ("top", "left", "bottom", "right"):
            edge = margins.find(f"w:{edge_name}", cls.WORD_NS)
            if edge is None:
                continue
            for attribute_name in ("w", "type"):
                attribute_value = cls._get_w_attr(edge, attribute_name)
                if attribute_value:
                    tokens.append(f"{prefix}:{edge_name}:{attribute_name}={attribute_value}")
        return tokens

    @staticmethod
    def _normalize_length(value) -> int:
        if value is None:
            return 0
        try:
            return int(value)
        except (TypeError, ValueError):
            return 0

    @staticmethod
    def _normalize_scalar(value) -> str:
        if value is None:
            return ""
        return str(value)

    @staticmethod
    def _normalize_enum(value) -> str:
        if value is None:
            return ""
        try:
            return str(int(value))
        except (TypeError, ValueError):
            return str(value)

    @staticmethod
    def _safe_name(value) -> str:
        return getattr(value, "name", "") or ""
