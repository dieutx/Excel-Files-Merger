from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook, load_workbook


SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}
INVALID_SHEET_TITLE_CHARS = set("[]:*?/\\")
MAX_SHEET_TITLE_LENGTH = 31
DEDUPLICATED_TITLE_BASE_LENGTH = 20
MERGE_MODE_COMBINE = "combine"
MERGE_MODE_SEPARATE = "separate"
COMBINED_SHEET_NAME = "Combined"
SOURCE_FILE_COLUMN = "_source_file"
SOURCE_SHEET_COLUMN = "_source_sheet"


@dataclass(frozen=True)
class MergeResult:
    files_processed: int
    sheets_processed: int
    rows_written: int
    output_path: Path


def discover_excel_files(folder: str | Path, exclude_path: str | Path | None = None) -> list[Path]:
    folder_path = Path(folder)
    excluded = Path(exclude_path).resolve() if exclude_path else None
    files: list[Path] = []

    for path in folder_path.iterdir():
        if not path.is_file():
            continue
        if path.name.startswith("~$"):
            continue
        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue
        if excluded and path.resolve() == excluded:
            continue
        files.append(path)

    return sorted(files, key=lambda item: item.name.lower())


def sanitize_sheet_title(raw_title: str, used_titles: set[str]) -> str:
    cleaned = "".join("_" if char in INVALID_SHEET_TITLE_CHARS else char for char in raw_title).strip()
    cleaned = cleaned.strip("'").strip()
    if not cleaned or set(cleaned) == {"_"}:
        cleaned = "Sheet"

    base = cleaned[:MAX_SHEET_TITLE_LENGTH]
    candidate = base
    counter = 2
    used_title_keys = {title.casefold() for title in used_titles}

    while candidate.casefold() in used_title_keys:
        suffix = f" ({counter})"
        base_length = min(DEDUPLICATED_TITLE_BASE_LENGTH, MAX_SHEET_TITLE_LENGTH - len(suffix))
        candidate = f"{base[:base_length]}{suffix}"
        counter += 1

    used_titles.add(candidate)
    return candidate


def merge_excel_files(folder: str | Path, output_path: str | Path, mode: str) -> MergeResult:
    output = Path(output_path)
    files = discover_excel_files(folder, exclude_path=output)
    if not files:
        raise ValueError("No Excel files were found in the selected folder.")

    if mode == MERGE_MODE_COMBINE:
        return _merge_combined(files, output)
    if mode == MERGE_MODE_SEPARATE:
        raise NotImplementedError("Separate sheet mode is implemented in a later task.")
    raise ValueError(f"Unsupported merge mode: {mode}")


def _merge_combined(files: Iterable[Path], output: Path) -> MergeResult:
    records: list[dict[str, object]] = []
    data_columns: list[str] = []
    files_processed = 0
    sheets_processed = 0

    for file_path in files:
        files_processed += 1
        workbook = load_workbook(file_path, read_only=True, data_only=True)
        try:
            for sheet in workbook.worksheets:
                rows = _non_empty_rows(sheet.iter_rows(values_only=True))
                if not rows:
                    continue

                sheets_processed += 1
                headers = _normalize_headers(rows[0])
                for header in headers:
                    if header not in data_columns:
                        data_columns.append(header)

                for row in rows[1:]:
                    record = {
                        SOURCE_FILE_COLUMN: file_path.name,
                        SOURCE_SHEET_COLUMN: sheet.title,
                    }
                    for header, value in zip(headers, row):
                        record[header] = value
                    records.append(record)
        finally:
            workbook.close()

    if not records:
        raise ValueError("The selected Excel files do not contain any data rows.")

    output.parent.mkdir(parents=True, exist_ok=True)
    output_workbook = Workbook()
    output_sheet = output_workbook.active
    output_sheet.title = COMBINED_SHEET_NAME

    columns = [SOURCE_FILE_COLUMN, SOURCE_SHEET_COLUMN, *data_columns]
    output_sheet.append(columns)
    for record in records:
        output_sheet.append([record.get(column) for column in columns])

    output_workbook.save(output)
    output_workbook.close()
    return MergeResult(files_processed, sheets_processed, len(records), output)


def _non_empty_rows(rows: Iterable[tuple[object, ...]]) -> list[tuple[object, ...]]:
    return [tuple(row) for row in rows if any(value is not None for value in row)]


def _normalize_headers(row: tuple[object, ...]) -> list[str]:
    base_headers: list[tuple[str, bool]] = []
    real_header_bases: set[str] = set()

    for index, value in enumerate(row, start=1):
        text = str(value).strip() if value is not None else ""
        is_blank = not text
        base = text if text else f"Column {index}"
        base_headers.append((base, is_blank))
        if not is_blank:
            real_header_bases.add(base)

    reserved_bases = {base for base, _ in base_headers}
    headers: list[str] = []
    used: set[str] = set()

    for base, is_blank in base_headers:
        if base not in used and not (is_blank and base in real_header_bases):
            header = base
        else:
            counter = 2
            header = f"{base} {counter}"
            while header in used or header in reserved_bases:
                counter += 1
                header = f"{base} {counter}"
        used.add(header)
        headers.append(header)

    return headers
