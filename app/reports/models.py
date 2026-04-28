from dataclasses import dataclass
from typing import Callable

from app.models import ParagraphData, TableCellData


type Opcode = tuple[str, int, int, int, int]
type ParagraphValue = str | ParagraphData
type TableCellValue = str | TableCellData


@dataclass(frozen=True)
class ExcelReportInput:
    excel_save_path: str
    log_callback: Callable[[str], None] | None = None
    paras_before: list[ParagraphValue] | None = None
    paras_after: list[ParagraphValue] | None = None
    get_loc_cb: Callable[[int, bool], str] | None = None
    flags_b: list[bool] | None = None
    flags_a: list[bool] | None = None
    tables_before: list[list[list[TableCellValue]]] | None = None
    tables_after: list[list[list[TableCellValue]]] | None = None


@dataclass(frozen=True)
class TableDiffPlan:
    index: int
    before_table: list[list[TableCellValue]]
    after_table: list[list[TableCellValue]]
    row_opcodes: list[Opcode]
    col_opcodes: list[Opcode]


@dataclass(frozen=True)
class ExcelDiffPlan:
    filtered_paras_before: list[ParagraphValue]
    filtered_paras_after: list[ParagraphValue]
    original_indices_before: list[int]
    original_indices_after: list[int]
    main_opcodes: list[Opcode]
    tables: list[TableDiffPlan]
