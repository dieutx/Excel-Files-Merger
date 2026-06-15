from pathlib import Path

from openpyxl import Workbook, load_workbook

from excel_merger import (
    MERGE_MODE_COMBINE,
    MERGE_MODE_SEPARATE,
    discover_excel_files,
    merge_excel_files,
    sanitize_sheet_title,
)


def make_workbook(path: Path, sheets: dict[str, list[list[object]]]) -> None:
    workbook = Workbook()
    default = workbook.active
    workbook.remove(default)

    for title, rows in sheets.items():
        sheet = workbook.create_sheet(title)
        for row in rows:
            sheet.append(row)

    workbook.save(path)


def read_sheet_rows(path: Path, sheet_name: str) -> list[tuple[object, ...]]:
    workbook = load_workbook(path, data_only=True)
    sheet = workbook[sheet_name]
    return [tuple(row) for row in sheet.iter_rows(values_only=True)]


def test_discovers_supported_excel_files_and_ignores_temp_files(tmp_path: Path):
    (tmp_path / "b.xlsx").write_bytes(b"")
    (tmp_path / "a.xlsm").write_bytes(b"")
    (tmp_path / "~$locked.xlsx").write_bytes(b"")
    (tmp_path / "notes.csv").write_text("not excel", encoding="utf-8")
    (tmp_path / "nested").mkdir()

    files = discover_excel_files(tmp_path)

    assert [path.name for path in files] == ["a.xlsm", "b.xlsx"]


def test_discovery_can_exclude_output_file_inside_source_folder(tmp_path: Path):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "combined.xlsx"
    source.write_bytes(b"")
    output.write_bytes(b"")

    files = discover_excel_files(tmp_path, exclude_path=output)

    assert [path.name for path in files] == ["source.xlsx"]


def test_sanitizes_and_deduplicates_excel_sheet_titles():
    used: set[str] = set()

    first = sanitize_sheet_title("2026/Q1:Report*East[Raw]", used)
    second = sanitize_sheet_title("2026/Q1:Report*East[Raw]", used)
    blank = sanitize_sheet_title("[]:*?/\\", used)

    assert first == "2026_Q1_Report_East_Raw_"
    assert second == "2026_Q1_Report_East_ (2)"
    assert blank == "Sheet"
    assert all(len(title) <= 31 for title in used)


def test_sanitizes_sheet_titles_by_stripping_outer_apostrophes_and_whitespace():
    used: set[str] = set()

    assert sanitize_sheet_title("'  Report  '", used) == "Report"
    assert sanitize_sheet_title("  'Quarterly Report'  ", used) == "Quarterly Report"


def test_blank_and_whitespace_sheet_titles_default_to_sheet():
    assert sanitize_sheet_title("", set()) == "Sheet"
    assert sanitize_sheet_title("   ", set()) == "Sheet"
    assert sanitize_sheet_title("'   '", set()) == "Sheet"
