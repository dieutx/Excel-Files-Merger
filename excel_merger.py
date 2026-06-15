from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}
INVALID_SHEET_TITLE_CHARS = set("[]:*?/\\")
MAX_SHEET_TITLE_LENGTH = 31
DEDUPLICATED_TITLE_BASE_LENGTH = 20
MERGE_MODE_COMBINE = "combine"
MERGE_MODE_SEPARATE = "separate"


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
    cleaned = cleaned.strip("'")
    if not cleaned or set(cleaned) == {"_"}:
        cleaned = "Sheet"

    base = cleaned[:MAX_SHEET_TITLE_LENGTH]
    candidate = base
    counter = 2

    while candidate in used_titles:
        suffix = f" ({counter})"
        base_length = min(DEDUPLICATED_TITLE_BASE_LENGTH, MAX_SHEET_TITLE_LENGTH - len(suffix))
        candidate = f"{base[:base_length]}{suffix}"
        counter += 1

    used_titles.add(candidate)
    return candidate


def merge_excel_files(folder: str | Path, output_path: str | Path, mode: str) -> MergeResult:
    raise NotImplementedError("Merge behavior is implemented in later tasks.")
