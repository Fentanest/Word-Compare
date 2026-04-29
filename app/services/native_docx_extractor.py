import json
import os
import subprocess
import sys
from pathlib import Path
from time import perf_counter

from app.models import ExtractedDocument, ParagraphData, RunData, TableCellData
from app.resources import resource_path


class NativeDocxExtractor:
    ENV_VAR = "WORD_COMPARE_NATIVE_EXTRACTOR"
    WINDOWS_BINARY = "word_compare_native_extractor.exe"
    POSIX_BINARY = "word_compare_native_extractor"

    def __init__(self, binary_path: str | None = None, run_command=None):
        self.binary_path = binary_path
        self.run_command = run_command or subprocess.run

    def extract(self, docx_path: str, log_callback=None) -> ExtractedDocument | None:
        binary_path = self.resolve_binary_path()
        if not binary_path:
            return None

        started_at = perf_counter()
        try:
            completed = self.run_command(
                [binary_path, docx_path],
                capture_output=True,
                check=True,
                text=True,
                encoding="utf-8",
            )
            if completed.stderr.strip():
                self._log(log_callback, f"-> Rust 추출기 메시지: {completed.stderr.strip()}")

            payload = json.loads(completed.stdout)
            extracted = self._payload_to_document(payload)
            self._log_perf(log_callback, "Rust 추출기 실행", started_at)
            return extracted
        except Exception as error:
            self._log(log_callback, f"-> Rust 추출기 사용 실패, Python 추출기로 대체합니다. ({error})")
            return None

    def resolve_binary_path(self) -> str | None:
        for candidate in self._candidate_paths():
            if candidate.is_file():
                return str(candidate)
        return None

    def _candidate_paths(self) -> list[Path]:
        candidates: list[Path] = []
        binary_name = self._binary_name()
        repo_root = Path(__file__).resolve().parents[2]

        if self.binary_path:
            candidates.append(Path(self.binary_path))

        env_binary_path = os.getenv(self.ENV_VAR)
        if env_binary_path:
            candidates.append(Path(env_binary_path))

        candidates.extend(
            [
                Path(resource_path(binary_name)),
                repo_root / "build" / "native" / binary_name,
                repo_root / "native" / "docx-structure-extractor" / "target" / "release" / binary_name,
            ]
        )
        return candidates

    @staticmethod
    def _payload_to_document(payload: dict) -> ExtractedDocument:
        paragraphs = []
        for paragraph_payload in payload.get("paragraphs", []):
            if paragraph_payload.get("type") == "marker":
                paragraphs.append(paragraph_payload.get("text", "[TABLE_MARKER]"))
            else:
                paragraphs.append(NativeDocxExtractor._paragraph_from_payload(paragraph_payload))

        tables = [
            [
                [NativeDocxExtractor._cell_from_payload(cell_payload) for cell_payload in row_payload]
                for row_payload in table_payload
            ]
            for table_payload in payload.get("tables", [])
        ]

        return ExtractedDocument(
            paragraphs=paragraphs,
            table_flags=list(payload.get("table_flags", [])),
            tables=tables,
            paragraph_locations=list(payload.get("paragraph_locations", [])),
        )

    @staticmethod
    def _paragraph_from_payload(payload: dict) -> ParagraphData:
        runs = tuple(NativeDocxExtractor._run_from_payload(run_payload) for run_payload in payload.get("runs", []))
        extra_meta = tuple(payload.get("extra_meta", []))
        return ParagraphData(
            text=payload.get("text", ""),
            style_name=payload.get("style_name", ""),
            alignment=payload.get("alignment", ""),
            source_kind=payload.get("source_kind", "body"),
            source_identifier=payload.get("source_identifier", ""),
            left_indent=int(payload.get("left_indent", 0)),
            right_indent=int(payload.get("right_indent", 0)),
            first_line_indent=int(payload.get("first_line_indent", 0)),
            space_before=int(payload.get("space_before", 0)),
            space_after=int(payload.get("space_after", 0)),
            line_spacing=payload.get("line_spacing", ""),
            keep_together=bool(payload.get("keep_together", False)),
            keep_with_next=bool(payload.get("keep_with_next", False)),
            page_break_before=bool(payload.get("page_break_before", False)),
            widow_control=bool(payload.get("widow_control", False)),
            runs=runs,
            extra_meta=extra_meta,
        )

    @staticmethod
    def _run_from_payload(payload: dict) -> RunData:
        return RunData(
            text=payload.get("text", ""),
            bold=bool(payload.get("bold", False)),
            italic=bool(payload.get("italic", False)),
            underline=payload.get("underline", ""),
            font_name=payload.get("font_name", ""),
            font_size=int(payload.get("font_size", 0)),
            color=payload.get("color", ""),
            highlight=payload.get("highlight", ""),
            strike=bool(payload.get("strike", False)),
            style_name=payload.get("style_name", ""),
        )

    @staticmethod
    def _cell_from_payload(payload: dict) -> TableCellData:
        return TableCellData(
            text=payload.get("text", ""),
            grid_span=int(payload.get("grid_span", 1)),
            v_merge=payload.get("v_merge", ""),
            cell_width=int(payload.get("cell_width", 0)),
            row_height=int(payload.get("row_height", 0)),
            grid_col_width=int(payload.get("grid_col_width", 0)),
            border_signature=payload.get("border_signature", ""),
            shading_fill=payload.get("shading_fill", ""),
            vertical_align=payload.get("vertical_align", ""),
            text_direction=payload.get("text_direction", ""),
            nested_table_count=int(payload.get("nested_table_count", 0)),
        )

    @staticmethod
    def _binary_name() -> str:
        if sys.platform.startswith("win"):
            return NativeDocxExtractor.WINDOWS_BINARY
        return NativeDocxExtractor.POSIX_BINARY

    @staticmethod
    def _log(log_callback, message: str) -> None:
        if log_callback:
            log_callback(message)

    @staticmethod
    def _log_perf(log_callback, label: str, started_at: float) -> None:
        NativeDocxExtractor._log(log_callback, f"{label}: {perf_counter() - started_at:.3f}초")
