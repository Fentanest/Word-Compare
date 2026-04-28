import os
from dataclasses import dataclass


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

    @property
    def signature(self) -> tuple[str, int, str, int, int, int]:
        return (
            self.text,
            self.grid_span,
            self.v_merge,
            self.cell_width,
            self.row_height,
            self.grid_col_width,
        )


@dataclass(frozen=True)
class ExtractedDocument:
    paragraphs: list[str]
    table_flags: list[bool]
    tables: list[list[list[str | TableCellData]]]


@dataclass(frozen=True)
class CompareResult:
    source_name: str
    result_docx_path: str | None = None
    result_excel_path: str | None = None
    error: str | None = None
