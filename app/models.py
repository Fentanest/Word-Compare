import os
from dataclasses import dataclass


def normalize_alignment_text(value: str) -> str:
    text = str(value).replace("\r", "\n").strip()
    if not text:
        return ""

    compact = (
        text.replace(",", "")
        .replace(".", "")
        .replace("%", "")
        .replace("+", "")
        .replace("-", "")
        .replace("(", "")
        .replace(")", "")
        .replace(" ", "")
    )
    if compact.isdigit():
        return "__NUMBER__"
    return text


@dataclass(frozen=True)
class AppSettings:
    save_path: str
    author: str
    excel_checked: bool


@dataclass(frozen=True)
class FilePair:
    before_path: str
    after_path: str

    @property
    def source_name(self) -> str:
        return os.path.basename(self.after_path)


@dataclass(frozen=True)
class CompareOptions:
    save_dir: str
    author_name: str
    generate_excel: bool

    @property
    def effective_author_name(self) -> str:
        return self.author_name.strip() or "Administrator"


@dataclass(frozen=True)
class TableCellData:
    text: str
    grid_span: int = 1
    v_merge: str = ""
    cell_width: int = 0
    row_height: int = 0
    grid_col_width: int = 0
    border_signature: str = ""
    shading_fill: str = ""
    vertical_align: str = ""
    text_direction: str = ""
    nested_table_count: int = 0

    @property
    def signature(self) -> tuple[str, int, str, int, int, int, str, str, str, str, int]:
        return (
            self.text,
            self.grid_span,
            self.v_merge,
            self.cell_width,
            self.row_height,
            self.grid_col_width,
            self.border_signature,
            self.shading_fill,
            self.vertical_align,
            self.text_direction,
            self.nested_table_count,
        )

    @property
    def alignment_signature(self) -> tuple[str, int, str, int, int, int, str, str, str, str, int]:
        return (
            normalize_alignment_text(self.text),
            self.grid_span,
            self.v_merge,
            self.cell_width,
            self.row_height,
            self.grid_col_width,
            self.border_signature,
            self.shading_fill,
            self.vertical_align,
            self.text_direction,
            self.nested_table_count,
        )


@dataclass(frozen=True)
class RunData:
    text: str
    bold: bool = False
    italic: bool = False
    underline: str = ""
    font_name: str = ""
    font_size: int = 0
    color: str = ""
    highlight: str = ""
    strike: bool = False
    style_name: str = ""

    @property
    def signature(self) -> tuple[str, bool, bool, str, str, int, str, str, bool, str]:
        return (
            self.text,
            self.bold,
            self.italic,
            self.underline,
            self.font_name,
            self.font_size,
            self.color,
            self.highlight,
            self.strike,
            self.style_name,
        )


@dataclass(frozen=True)
class ParagraphData:
    text: str
    style_name: str = ""
    alignment: str = ""
    source_kind: str = "body"
    source_identifier: str = ""
    left_indent: int = 0
    right_indent: int = 0
    first_line_indent: int = 0
    space_before: int = 0
    space_after: int = 0
    line_spacing: str = ""
    keep_together: bool = False
    keep_with_next: bool = False
    page_break_before: bool = False
    widow_control: bool = False
    runs: tuple[RunData, ...] = ()
    extra_meta: tuple[str, ...] = ()

    @property
    def signature(self) -> tuple[str, str, str, str, str, int, int, int, int, int, str, bool, bool, bool, bool, tuple, tuple[str, ...]]:
        return (
            self.text,
            self.style_name,
            self.alignment,
            self.source_kind,
            self.source_identifier,
            self.left_indent,
            self.right_indent,
            self.first_line_indent,
            self.space_before,
            self.space_after,
            self.line_spacing,
            self.keep_together,
            self.keep_with_next,
            self.page_break_before,
            self.widow_control,
            tuple(run.signature for run in self.runs),
            self.extra_meta,
        )


@dataclass(frozen=True)
class ExtractedDocument:
    paragraphs: list[str | ParagraphData]
    table_flags: list[bool]
    tables: list[list[list[str | TableCellData]]]
    paragraph_locations: list[str] | None = None


@dataclass(frozen=True)
class CompareResult:
    source_name: str
    result_docx_path: str | None = None
    result_excel_path: str | None = None
    error: str | None = None
