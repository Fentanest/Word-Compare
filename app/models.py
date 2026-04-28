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
class ExtractedDocument:
    paragraphs: list[str]
    table_flags: list[bool]
    tables: list[list[list[str]]]


@dataclass(frozen=True)
class CompareResult:
    source_name: str
    result_docx_path: str | None = None
    result_excel_path: str | None = None
    error: str | None = None
