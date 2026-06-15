# Excel Merger Refresh Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a polished Tkinter Excel merger with two merge modes, tests, documentation, and GitHub Actions artifacts for portable Windows 32-bit and 64-bit executables.

**Architecture:** Move workbook discovery and merge behavior into `excel_merger.py` so it can be tested without launching the GUI. Keep `run.py` as the Tkinter/ttk entrypoint and call the merge API from a background thread. Use `openpyxl` only for Excel I/O to reduce dependency and Windows 32-bit packaging risk.

**Tech Stack:** Python 3.11+, Tkinter/ttk, openpyxl, pytest, PyInstaller, GitHub Actions on `windows-latest`.

---

## File Structure

- Create `excel_merger.py`: workbook discovery, sheet-name sanitization, merge mode enum constants, merge API, and result dataclass.
- Replace `run.py`: Tkinter/ttk dashboard UI that imports `excel_merger.py`.
- Create `tests/test_excel_merger.py`: fast non-GUI tests covering discovery, sheet names, combine mode, and separate sheet mode.
- Create `requirements.txt`: runtime dependency list.
- Create `requirements-dev.txt`: test/build dependency list.
- Create `.github/workflows/build-windows.yml`: test and Windows artifact build workflow.
- Modify `README.md`: accurate usage, development, and artifact download instructions.
- Modify `.gitignore`: keep generated PyInstaller outputs ignored while allowing source files and workflow files.

---

### Task 1: Test File Discovery And Sheet Name Helpers

**Files:**
- Create: `tests/test_excel_merger.py`
- Create: `requirements.txt`
- Create: `requirements-dev.txt`

- [ ] **Step 1: Add dependency files**

Create `requirements.txt`:

```text
openpyxl>=3.1,<4
```

Create `requirements-dev.txt`:

```text
-r requirements.txt
pytest>=8,<9
pyinstaller>=6,<7
```

- [ ] **Step 2: Write failing tests for discovery and sheet title sanitization**

Create `tests/test_excel_merger.py`:

```python
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
```

- [ ] **Step 3: Run tests to verify they fail**

Run:

```bash
python -m pytest tests/test_excel_merger.py -v
```

Expected: FAIL because `excel_merger` does not exist.

- [ ] **Step 4: Leave the failing tests uncommitted until the helper implementation is green**

Expected: `requirements.txt`, `requirements-dev.txt`, and `tests/test_excel_merger.py` remain unstaged until Task 2 passes.

---

### Task 2: Implement Discovery And Sheet Name Helpers

**Files:**
- Create: `excel_merger.py`
- Test: `tests/test_excel_merger.py`

- [ ] **Step 1: Implement minimal helper code**

Create `excel_merger.py`:

```python
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}
INVALID_SHEET_TITLE_CHARS = set("[]:*?/\\")
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

    base = cleaned[:31]
    candidate = base
    counter = 2

    while candidate in used_titles:
        suffix = f" ({counter})"
        candidate = f"{base[:31 - len(suffix)]}{suffix}"
        counter += 1

    used_titles.add(candidate)
    return candidate


def merge_excel_files(folder: str | Path, output_path: str | Path, mode: str) -> MergeResult:
    raise NotImplementedError("Merge behavior is implemented in later tasks.")
```

- [ ] **Step 2: Run helper tests**

Run:

```bash
python -m pytest tests/test_excel_merger.py::test_discovers_supported_excel_files_and_ignores_temp_files tests/test_excel_merger.py::test_discovery_can_exclude_output_file_inside_source_folder tests/test_excel_merger.py::test_sanitizes_and_deduplicates_excel_sheet_titles -v
```

Expected: PASS for the three helper tests.

- [ ] **Step 3: Run full test file**

Run:

```bash
python -m pytest tests/test_excel_merger.py -v
```

Expected: PASS for helper tests; no merge behavior tests exist yet.

- [ ] **Step 4: Commit helper implementation and passing helper tests**

```bash
git add excel_merger.py requirements.txt requirements-dev.txt tests/test_excel_merger.py
git commit -m "feat: add Excel file discovery helpers"
```

---

### Task 3: Test And Implement Combined Worksheet Merge

**Files:**
- Modify: `tests/test_excel_merger.py`
- Modify: `excel_merger.py`

- [ ] **Step 1: Add failing tests for combined merge mode**

Append to `tests/test_excel_merger.py`:

```python
def test_combines_workbooks_into_one_sheet_with_source_metadata(tmp_path: Path):
    make_workbook(
        tmp_path / "a_sales.xlsx",
        {"North": [["Name", "Amount"], ["Ada", 10], ["Ben", 20]]},
    )
    make_workbook(
        tmp_path / "b_returns.xlsx",
        {"South": [["Name", "Reason"], ["Cora", "Damaged"]]},
    )
    output = tmp_path / "combined.xlsx"

    result = merge_excel_files(tmp_path, output, MERGE_MODE_COMBINE)

    rows = read_sheet_rows(output, "Combined")
    assert result.files_processed == 2
    assert result.sheets_processed == 2
    assert result.rows_written == 3
    assert rows == [
        ("_source_file", "_source_sheet", "Name", "Amount", "Reason"),
        ("a_sales.xlsx", "North", "Ada", 10, None),
        ("a_sales.xlsx", "North", "Ben", 20, None),
        ("b_returns.xlsx", "South", "Cora", None, "Damaged"),
    ]


def test_combined_mode_fails_cleanly_when_no_excel_files_exist(tmp_path: Path):
    output = tmp_path / "combined.xlsx"

    try:
        merge_excel_files(tmp_path, output, MERGE_MODE_COMBINE)
    except ValueError as exc:
        assert str(exc) == "No Excel files were found in the selected folder."
    else:
        raise AssertionError("Expected ValueError for an empty source folder")
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```bash
python -m pytest tests/test_excel_merger.py::test_combines_workbooks_into_one_sheet_with_source_metadata tests/test_excel_merger.py::test_combined_mode_fails_cleanly_when_no_excel_files_exist -v
```

Expected: FAIL because `merge_excel_files` raises `NotImplementedError`.

- [ ] **Step 3: Implement combined merge mode**

Replace `excel_merger.py` with:

```python
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook, load_workbook


SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}
INVALID_SHEET_TITLE_CHARS = set("[]:*?/\\")
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
    cleaned = cleaned.strip("'")
    if not cleaned or set(cleaned) == {"_"}:
        cleaned = "Sheet"

    base = cleaned[:31]
    candidate = base
    counter = 2

    while candidate in used_titles:
        suffix = f" ({counter})"
        candidate = f"{base[:31 - len(suffix)]}{suffix}"
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
    return MergeResult(files_processed, sheets_processed, len(records), output)


def _non_empty_rows(rows: Iterable[tuple[object, ...]]) -> list[tuple[object, ...]]:
    return [tuple(row) for row in rows if any(value is not None for value in row)]


def _normalize_headers(row: tuple[object, ...]) -> list[str]:
    headers: list[str] = []
    seen: dict[str, int] = {}

    for index, value in enumerate(row, start=1):
        header = str(value).strip() if value is not None and str(value).strip() else f"Column {index}"
        count = seen.get(header, 0) + 1
        seen[header] = count
        if count > 1:
            header = f"{header} {count}"
        headers.append(header)

    return headers
```

- [ ] **Step 4: Run combined merge tests**

Run:

```bash
python -m pytest tests/test_excel_merger.py::test_combines_workbooks_into_one_sheet_with_source_metadata tests/test_excel_merger.py::test_combined_mode_fails_cleanly_when_no_excel_files_exist -v
```

Expected: PASS.

- [ ] **Step 5: Run full test file**

Run:

```bash
python -m pytest tests/test_excel_merger.py -v
```

Expected: PASS.

- [ ] **Step 6: Commit combined merge mode**

```bash
git add excel_merger.py tests/test_excel_merger.py
git commit -m "feat: combine Excel files into one sheet"
```

---

### Task 4: Test And Implement Separate Sheet Merge

**Files:**
- Modify: `tests/test_excel_merger.py`
- Modify: `excel_merger.py`

- [ ] **Step 1: Add failing tests for separate sheet mode**

Append to `tests/test_excel_merger.py`:

```python
def test_keeps_source_sheets_separate_with_safe_unique_names(tmp_path: Path):
    make_workbook(
        tmp_path / "east.xlsx",
        {"Raw/Data": [["Name"], ["Ada"]]},
    )
    make_workbook(
        tmp_path / "east-copy.xlsx",
        {"Raw/Data": [["Name"], ["Ben"]]},
    )
    output = tmp_path / "separate.xlsx"

    result = merge_excel_files(tmp_path, output, MERGE_MODE_SEPARATE)

    workbook = load_workbook(output, data_only=True)
    assert workbook.sheetnames == ["east - Raw_Data", "east-copy - Raw_Data"]
    assert read_sheet_rows(output, "east - Raw_Data") == [("Name",), ("Ada",)]
    assert read_sheet_rows(output, "east-copy - Raw_Data") == [("Name",), ("Ben",)]
    assert result.files_processed == 2
    assert result.sheets_processed == 2
    assert result.rows_written == 4


def test_rejects_unknown_merge_mode(tmp_path: Path):
    make_workbook(tmp_path / "source.xlsx", {"Sheet1": [["Name"], ["Ada"]]})

    try:
        merge_excel_files(tmp_path, tmp_path / "out.xlsx", "unknown")
    except ValueError as exc:
        assert str(exc) == "Unsupported merge mode: unknown"
    else:
        raise AssertionError("Expected ValueError for unsupported mode")
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```bash
python -m pytest tests/test_excel_merger.py::test_keeps_source_sheets_separate_with_safe_unique_names tests/test_excel_merger.py::test_rejects_unknown_merge_mode -v
```

Expected: separate sheet test FAILS because separate mode is not implemented; unknown mode test PASSES.

- [ ] **Step 3: Implement separate sheet mode**

Modify `excel_merger.py`:

```python
def merge_excel_files(folder: str | Path, output_path: str | Path, mode: str) -> MergeResult:
    output = Path(output_path)
    files = discover_excel_files(folder, exclude_path=output)
    if not files:
        raise ValueError("No Excel files were found in the selected folder.")

    if mode == MERGE_MODE_COMBINE:
        return _merge_combined(files, output)
    if mode == MERGE_MODE_SEPARATE:
        return _merge_separate(files, output)
    raise ValueError(f"Unsupported merge mode: {mode}")


def _merge_separate(files: Iterable[Path], output: Path) -> MergeResult:
    output.parent.mkdir(parents=True, exist_ok=True)
    output_workbook = Workbook()
    default_sheet = output_workbook.active
    output_workbook.remove(default_sheet)

    used_titles: set[str] = set()
    files_processed = 0
    sheets_processed = 0
    rows_written = 0

    for file_path in files:
        files_processed += 1
        workbook = load_workbook(file_path, read_only=True, data_only=True)
        try:
            for sheet in workbook.worksheets:
                rows = _non_empty_rows(sheet.iter_rows(values_only=True))
                if not rows:
                    continue
                title = sanitize_sheet_title(f"{file_path.stem} - {sheet.title}", used_titles)
                output_sheet = output_workbook.create_sheet(title)
                for row in rows:
                    output_sheet.append(row)
                    rows_written += 1
                sheets_processed += 1
        finally:
            workbook.close()

    if not output_workbook.worksheets:
        raise ValueError("The selected Excel files do not contain any usable sheets.")

    output_workbook.save(output)
    return MergeResult(files_processed, sheets_processed, rows_written, output)
```

- [ ] **Step 4: Run separate sheet tests**

Run:

```bash
python -m pytest tests/test_excel_merger.py::test_keeps_source_sheets_separate_with_safe_unique_names tests/test_excel_merger.py::test_rejects_unknown_merge_mode -v
```

Expected: PASS.

- [ ] **Step 5: Run full tests**

Run:

```bash
python -m pytest -v
```

Expected: PASS.

- [ ] **Step 6: Commit separate sheet mode**

```bash
git add excel_merger.py tests/test_excel_merger.py
git commit -m "feat: keep Excel sheets separate"
```

---

### Task 5: Build Tkinter Dashboard UI

**Files:**
- Replace: `run.py`

- [ ] **Step 1: Replace the old UI with a dashboard app**

Replace `run.py` with:

```python
from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from excel_merger import (
    MERGE_MODE_COMBINE,
    MERGE_MODE_SEPARATE,
    discover_excel_files,
    merge_excel_files,
)


class ExcelMergerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Excel Files Merger")
        self.root.minsize(760, 520)

        self.source_folder = tk.StringVar()
        self.output_file = tk.StringVar()
        self.merge_mode = tk.StringVar(value=MERGE_MODE_COMBINE)
        self.summary = tk.StringVar(value="Choose a source folder to begin.")

        self._configure_style()
        self._build_layout()

    def _configure_style(self) -> None:
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("TFrame", background="#f5f7fb")
        style.configure("Panel.TFrame", background="#ffffff", relief="flat")
        style.configure("Title.TLabel", background="#f5f7fb", foreground="#172033", font=("Segoe UI", 18, "bold"))
        style.configure("Subtitle.TLabel", background="#f5f7fb", foreground="#5d6b82", font=("Segoe UI", 10))
        style.configure("TLabel", background="#ffffff", foreground="#172033", font=("Segoe UI", 10))
        style.configure("Muted.TLabel", background="#ffffff", foreground="#5d6b82", font=("Segoe UI", 9))
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 8))
        style.configure("TButton", font=("Segoe UI", 10), padding=(10, 6))
        style.configure("TRadiobutton", background="#ffffff", foreground="#172033", font=("Segoe UI", 10))

    def _build_layout(self) -> None:
        container = ttk.Frame(self.root, padding=24)
        container.pack(fill=tk.BOTH, expand=True)

        ttk.Label(container, text="Excel Files Merger", style="Title.TLabel").pack(anchor=tk.W)
        ttk.Label(
            container,
            text="Merge multiple Excel workbooks into one portable output file.",
            style="Subtitle.TLabel",
        ).pack(anchor=tk.W, pady=(4, 18))

        panel = ttk.Frame(container, style="Panel.TFrame", padding=18)
        panel.pack(fill=tk.X)
        panel.columnconfigure(1, weight=1)

        self._path_row(panel, 0, "Source folder", self.source_folder, self.choose_source_folder)
        self._path_row(panel, 1, "Output file", self.output_file, self.choose_output_file)

        options = ttk.Frame(panel, style="Panel.TFrame")
        options.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(16, 4))
        ttk.Label(options, text="Merge mode").pack(anchor=tk.W)
        ttk.Radiobutton(
            options,
            text="Combine all files into one worksheet",
            variable=self.merge_mode,
            value=MERGE_MODE_COMBINE,
        ).pack(anchor=tk.W, pady=(8, 2))
        ttk.Radiobutton(
            options,
            text="Keep source sheets separate",
            variable=self.merge_mode,
            value=MERGE_MODE_SEPARATE,
        ).pack(anchor=tk.W)

        status = ttk.Frame(container, style="Panel.TFrame", padding=18)
        status.pack(fill=tk.BOTH, expand=True, pady=(16, 0))
        status.columnconfigure(0, weight=1)

        ttk.Label(status, textvariable=self.summary, style="Muted.TLabel").grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(status, mode="indeterminate")
        self.progress.grid(row=1, column=0, sticky="ew", pady=(12, 12))

        self.log = scrolledtext.ScrolledText(status, height=8, wrap=tk.WORD, font=("Consolas", 9))
        self.log.grid(row=2, column=0, sticky="nsew")

        actions = ttk.Frame(container, style="TFrame")
        actions.pack(fill=tk.X, pady=(16, 0))
        self.merge_button = ttk.Button(actions, text="Merge Excel Files", style="Primary.TButton", command=self.merge)
        self.merge_button.pack(side=tk.RIGHT)

    def _path_row(self, parent: ttk.Frame, row: int, label: str, variable: tk.StringVar, command) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=(0, 12))
        entry = ttk.Entry(parent, textvariable=variable)
        entry.grid(row=row, column=1, sticky="ew", padx=12, pady=(0, 12))
        ttk.Button(parent, text="Browse", command=command).grid(row=row, column=2, sticky="e", pady=(0, 12))

    def choose_source_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select folder containing Excel files")
        if not folder:
            return
        self.source_folder.set(folder)
        self.refresh_summary()

    def choose_output_file(self) -> None:
        file_path = filedialog.asksaveasfilename(
            title="Save merged workbook as",
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if file_path:
            self.output_file.set(file_path)

    def refresh_summary(self) -> None:
        folder = self.source_folder.get()
        if not folder:
            self.summary.set("Choose a source folder to begin.")
            return
        try:
            files = discover_excel_files(folder, exclude_path=self.output_file.get() or None)
        except OSError as exc:
            self.summary.set(f"Cannot read folder: {exc}")
            return
        self.summary.set(f"{len(files)} Excel file(s) ready to merge.")

    def merge(self) -> None:
        folder = self.source_folder.get()
        output = self.output_file.get()
        if not folder or not output:
            messagebox.showerror("Missing selection", "Choose both a source folder and an output file.")
            return

        self.merge_button.state(["disabled"])
        self.progress.start(10)
        self.log.delete("1.0", tk.END)
        self.log.insert(tk.END, "Merging files...\n")

        worker = threading.Thread(target=self._merge_worker, args=(folder, output, self.merge_mode.get()), daemon=True)
        worker.start()

    def _merge_worker(self, folder: str, output: str, mode: str) -> None:
        try:
            result = merge_excel_files(folder, output, mode)
        except Exception as exc:
            self.root.after(0, self._merge_failed, exc)
            return
        self.root.after(0, self._merge_finished, result)

    def _merge_finished(self, result) -> None:
        self.progress.stop()
        self.merge_button.state(["!disabled"])
        self.summary.set(f"Created {result.output_path.name} from {result.files_processed} file(s).")
        self.log.insert(tk.END, f"Done.\nFiles: {result.files_processed}\nSheets: {result.sheets_processed}\nRows: {result.rows_written}\nOutput: {result.output_path}\n")
        messagebox.showinfo("Merge complete", f"Created:\n{result.output_path}")

    def _merge_failed(self, exc: Exception) -> None:
        self.progress.stop()
        self.merge_button.state(["!disabled"])
        self.summary.set("Merge failed. Check the message below.")
        self.log.insert(tk.END, f"Error: {exc}\n")
        messagebox.showerror("Merge failed", str(exc))


def main() -> None:
    root = tk.Tk()
    app = ExcelMergerApp(root)
    app.refresh_summary()
    root.mainloop()


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Validate syntax**

Run:

```bash
python -m py_compile run.py excel_merger.py
```

Expected: exit code 0.

- [ ] **Step 3: Run tests**

Run:

```bash
python -m pytest -v
```

Expected: PASS.

- [ ] **Step 4: Commit UI refresh**

```bash
git add run.py
git commit -m "feat: refresh Tkinter dashboard UI"
```

---

### Task 6: Add Build Workflow And Documentation

**Files:**
- Create: `.github/workflows/build-windows.yml`
- Modify: `README.md`
- Modify: `.gitignore`

- [ ] **Step 1: Add GitHub Actions workflow**

Create `.github/workflows/build-windows.yml`:

```yaml
name: Build Windows EXE

on:
  push:
    branches: [main]
  pull_request:
    branches: [main]
  workflow_dispatch:

jobs:
  test:
    name: Test on Linux
    runs-on: ubuntu-latest
    steps:
      - name: Checkout
        uses: actions/checkout@v4

      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.11"

      - name: Install dependencies
        run: |
          python -m pip install --upgrade pip
          python -m pip install -r requirements-dev.txt

      - name: Run tests
        run: python -m pytest -v

  build-windows:
    name: Build Windows ${{ matrix.label }}
    runs-on: windows-latest
    needs: test
    strategy:
      fail-fast: false
      matrix:
        include:
          - label: win32
            architecture: x86
            artifact: Excel-Files-Merger-win32
          - label: win64
            architecture: x64
            artifact: Excel-Files-Merger-win64

    steps:
      - name: Checkout
        uses: actions/checkout@v4

      - name: Set up Python ${{ matrix.architecture }}
        uses: actions/setup-python@v5
        with:
          python-version: "3.11"
          architecture: ${{ matrix.architecture }}

      - name: Install dependencies
        run: |
          python -m pip install --upgrade pip
          python -m pip install -r requirements-dev.txt

      - name: Run tests
        run: python -m pytest -v

      - name: Build executable
        run: >
          python -m PyInstaller
          --noconfirm
          --clean
          --onefile
          --windowed
          --name Excel-Files-Merger
          run.py

      - name: Upload executable artifact
        uses: actions/upload-artifact@v4
        with:
          name: ${{ matrix.artifact }}
          path: dist/Excel-Files-Merger.exe
          if-no-files-found: error
```

- [ ] **Step 2: Update README**

Replace `README.md` with:

```markdown
# Excel Files Merger

A small Windows desktop utility for merging Excel workbooks into one output file.

## Features

- Select a source folder and output workbook from a simple desktop UI.
- Merge `.xlsx` and `.xlsm` files.
- Ignore temporary Excel lock files such as `~$report.xlsx`.
- Choose either:
  - combine all source data into one worksheet, or
  - keep source sheets separate in the output workbook.
- Download portable Windows `.exe` artifacts from GitHub Actions.

## Download Portable Windows Builds

1. Open the repository on GitHub.
2. Go to **Actions**.
3. Open the latest **Build Windows EXE** workflow run.
4. Download either:
   - `Excel-Files-Merger-win64` for most Windows machines.
   - `Excel-Files-Merger-win32` for older 32-bit Windows machines.
5. Extract the downloaded artifact and run `Excel-Files-Merger.exe`.

No Python installation is required for the downloaded `.exe`.

## Run From Source

```sh
python -m pip install -r requirements.txt
python run.py
```

## Build Locally On Windows

```sh
python -m pip install -r requirements-dev.txt
python -m PyInstaller --noconfirm --clean --onefile --windowed --name Excel-Files-Merger run.py
```

The executable will be created at `dist/Excel-Files-Merger.exe`.

## Development

```sh
python -m pip install -r requirements-dev.txt
python -m pytest -v
python -m py_compile run.py excel_merger.py
```

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
```

- [ ] **Step 3: Confirm `.gitignore` still ignores only generated outputs**

Ensure `.gitignore` contains:

```text
.superpowers/
__pycache__/
.pytest_cache/
build/
dist/
*.spec
*.exe
```

- [ ] **Step 4: Run verification**

Run:

```bash
python -m pytest -v
python -m py_compile run.py excel_merger.py
```

Expected: both commands exit 0.

- [ ] **Step 5: Commit docs and workflow**

```bash
git add .github/workflows/build-windows.yml README.md .gitignore requirements.txt requirements-dev.txt
git commit -m "ci: build portable Windows executables"
```

---

### Task 7: Final Verification And Push

**Files:**
- Verify all changed files.

- [ ] **Step 1: Inspect status and recent commits**

Run:

```bash
git status -sb
git log --oneline -8
```

Expected: working tree clean except ignored `.superpowers/`, and commits show the design plus implementation commits.

- [ ] **Step 2: Run final verification**

Run:

```bash
python -m pytest -v
python -m py_compile run.py excel_merger.py
```

Expected: both commands exit 0.

- [ ] **Step 3: Push to GitHub**

If the user wants direct push to `main`, run:

```bash
git push origin main
```

Expected: push succeeds and GitHub Actions starts the `Build Windows EXE` workflow.

- [ ] **Step 4: Report artifact path**

Tell the user that the portable executables will appear in the GitHub Actions run artifacts:

```text
Actions -> Build Windows EXE -> latest run -> Artifacts
```

Mention both artifact names:

```text
Excel-Files-Merger-win32
Excel-Files-Merger-win64
```

---

## Self-Review

- Spec coverage: dashboard UI is Task 5; merge modes are Tasks 3 and 4; friendly errors are Tasks 3, 4, and 5; dependency/test files are Task 1; GitHub Actions Windows 32/64 artifacts and docs are Task 6; final push is Task 7.
- Placeholder scan: no task uses forbidden placeholder markers.
- Type consistency: constants, function names, and result fields are introduced in Task 1/2 and reused consistently in later tasks.
