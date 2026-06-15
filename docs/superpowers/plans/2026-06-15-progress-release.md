# Progress Logging and Release Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add real merge progress, timestamped user-visible logs, Ubuntu portable builds, and versioned GitHub Releases for Windows and Ubuntu downloads.

**Architecture:** Keep merge behavior in `excel_merger.py` independent from Tkinter by adding a structured progress callback. `run.py` will translate progress events into determinate progress bar updates, percentage text, status text, and timestamped log entries. GitHub Actions will build three portable assets and publish them to a release only when a `v*` tag is pushed.

**Tech Stack:** Python 3.10, Tkinter/ttk, openpyxl, pytest, PyInstaller, GitHub Actions.

---

## File Structure

- Modify `excel_merger.py`: add `MergeProgressEvent`, `ProgressCallback`, event string handling, and callback emission inside both merge modes.
- Modify `tests/test_excel_merger.py`: add tests for progress event behavior in combined and separate modes.
- Modify `run.py`: display determinate progress, percent text, timestamped logs, and worker-thread progress event routing.
- Rename `.github/workflows/build-windows.yml` to `.github/workflows/build-release.yml`: build Windows x86, Windows x64, and Ubuntu x64 assets; upload artifacts; publish releases on `v*` tags.
- Modify `README.md`: document progress/log behavior, GitHub Releases, and local Ubuntu build commands.

---

### Task 1: Core Progress Event Contract

**Files:**
- Modify: `tests/test_excel_merger.py`
- Modify: `excel_merger.py`

- [ ] **Step 1: Write the failing combined-mode progress test**

Add this test to `tests/test_excel_merger.py` after `test_combines_workbooks_into_one_sheet_with_source_metadata`:

```python
def test_combined_mode_emits_progress_events(tmp_path: Path):
    make_workbook(
        tmp_path / "a_sales.xlsx",
        {"North": [["Name"], ["Ada"]]},
    )
    make_workbook(
        tmp_path / "b_sales.xlsx",
        {"South": [["Name"], ["Ben"]]},
    )
    output = tmp_path / "combined.xlsx"
    events = []

    result = merge_excel_files(
        tmp_path,
        output,
        MERGE_MODE_COMBINE,
        on_progress=events.append,
    )

    event_types = [event.event_type for event in events]
    assert event_types == [
        "scan_complete",
        "file_start",
        "sheet_processed",
        "file_complete",
        "file_start",
        "sheet_processed",
        "file_complete",
        "writing_output",
        "complete",
    ]
    assert [event.current_file_index for event in events if event.event_type == "file_start"] == [1, 2]
    assert all(event.total_files == 2 for event in events)
    assert events[-1].files_processed == result.files_processed == 2
    assert events[-1].sheets_processed == result.sheets_processed == 2
    assert events[-1].rows_written == result.rows_written == 2
    assert events[-1].output_path == output
```

- [ ] **Step 2: Run the combined-mode progress test and verify it fails**

Run:

```bash
python -m pytest tests/test_excel_merger.py::test_combined_mode_emits_progress_events -v
```

Expected: FAIL with `TypeError: merge_excel_files() got an unexpected keyword argument 'on_progress'`.

- [ ] **Step 3: Add progress datatypes and callback plumbing**

In `excel_merger.py`, update imports:

```python
from collections.abc import Callable, Iterable
```

Add this dataclass after `MergeResult`:

```python
@dataclass(frozen=True)
class MergeProgressEvent:
    event_type: str
    total_files: int
    current_file_index: int = 0
    file_path: Path | None = None
    sheet_title: str | None = None
    files_processed: int = 0
    sheets_processed: int = 0
    rows_written: int = 0
    output_path: Path | None = None


ProgressCallback = Callable[[MergeProgressEvent], None]
```

Change the public function signature:

```python
def merge_excel_files(
    folder: str | Path,
    output_path: str | Path,
    mode: str,
    on_progress: ProgressCallback | None = None,
) -> MergeResult:
```

Inside `merge_excel_files()`, emit scan progress and pass `on_progress` into mode-specific helpers:

```python
    if on_progress:
        on_progress(MergeProgressEvent("scan_complete", total_files=len(files)))

    if mode == MERGE_MODE_COMBINE:
        return _merge_combined(files, output, on_progress)
    if mode == MERGE_MODE_SEPARATE:
        return _merge_separate(files, output, on_progress)
```

Update helper signatures:

```python
def _merge_combined(
    files: Iterable[Path],
    output: Path,
    on_progress: ProgressCallback | None = None,
) -> MergeResult:
```

```python
def _merge_separate(
    files: Iterable[Path],
    output: Path,
    on_progress: ProgressCallback | None = None,
) -> MergeResult:
```

At the start of each helper, materialize files once:

```python
    file_list = list(files)
    total_files = len(file_list)
```

Loop with an index:

```python
    for file_index, file_path in enumerate(file_list, start=1):
```

Emit `file_start` before loading each workbook in combined mode:

```python
        if on_progress:
            on_progress(
                MergeProgressEvent(
                    "file_start",
                    total_files=total_files,
                    current_file_index=file_index,
                    file_path=file_path,
                    files_processed=files_processed,
                    sheets_processed=sheets_processed,
                    rows_written=len(records),
                )
            )
```

For combined mode, after each non-empty sheet has been read and records were appended, emit `sheet_processed`:

```python
                    if on_progress:
                        on_progress(
                            MergeProgressEvent(
                                "sheet_processed",
                                total_files=total_files,
                                current_file_index=file_index,
                                file_path=file_path,
                                sheet_title=sheet.title,
                                files_processed=files_processed,
                                sheets_processed=sheets_processed,
                                rows_written=len(records),
                            )
                        )
```

After closing each workbook in combined mode, emit `file_complete`:

```python
        if on_progress:
            on_progress(
                MergeProgressEvent(
                    "file_complete",
                    total_files=total_files,
                    current_file_index=file_index,
                    file_path=file_path,
                    files_processed=files_processed,
                    sheets_processed=sheets_processed,
                    rows_written=len(records),
                )
            )
```

Before saving in combined mode, emit `writing_output`:

```python
    if on_progress:
        on_progress(
            MergeProgressEvent(
                "writing_output",
                total_files=total_files,
                current_file_index=total_files,
                files_processed=files_processed,
                sheets_processed=sheets_processed,
                rows_written=len(records),
                output_path=output,
            )
        )
```

After creating the combined-mode `MergeResult`, emit `complete`:

```python
    result = MergeResult(files_processed, sheets_processed, len(records), output)
    if on_progress:
        on_progress(
            MergeProgressEvent(
                "complete",
                total_files=total_files,
                current_file_index=total_files,
                files_processed=result.files_processed,
                sheets_processed=result.sheets_processed,
                rows_written=result.rows_written,
                output_path=result.output_path,
            )
        )
    return result
```

- [ ] **Step 4: Run the combined-mode progress test and verify it passes**

Run:

```bash
python -m pytest tests/test_excel_merger.py::test_combined_mode_emits_progress_events -v
```

Expected: PASS.

- [ ] **Step 5: Write the failing separate-mode progress test**

Add this test after `test_keeps_source_sheets_separate_with_safe_unique_names`:

```python
def test_separate_mode_emits_progress_events(tmp_path: Path):
    make_workbook(
        tmp_path / "a_east.xlsx",
        {"Raw": [["Name"], ["Ada"]]},
    )
    make_workbook(
        tmp_path / "b_west.xlsx",
        {"Raw": [["Name"], ["Ben"]]},
    )
    output = tmp_path / "separate.xlsx"
    events = []

    result = merge_excel_files(
        tmp_path,
        output,
        MERGE_MODE_SEPARATE,
        on_progress=events.append,
    )

    event_types = [event.event_type for event in events]
    assert event_types == [
        "scan_complete",
        "file_start",
        "sheet_processed",
        "file_complete",
        "file_start",
        "sheet_processed",
        "file_complete",
        "writing_output",
        "complete",
    ]
    assert [event.current_file_index for event in events if event.event_type == "file_complete"] == [1, 2]
    assert all(event.total_files == 2 for event in events)
    assert events[-1].files_processed == result.files_processed == 2
    assert events[-1].sheets_processed == result.sheets_processed == 2
    assert events[-1].rows_written == result.rows_written == 4
    assert events[-1].output_path == output
```

- [ ] **Step 6: Run the separate-mode progress test and verify it fails for the separate helper**

Run:

```bash
python -m pytest tests/test_excel_merger.py::test_separate_mode_emits_progress_events -v
```

Expected: FAIL until `_merge_separate()` emits the same event contract.

- [ ] **Step 7: Implement separate-mode progress emissions**

Mirror the combined-mode callback structure in `_merge_separate()`, but use `rows_written` rather than `len(records)`.

Emit `file_start` before loading each workbook:

```python
        if on_progress:
            on_progress(
                MergeProgressEvent(
                    "file_start",
                    total_files=total_files,
                    current_file_index=file_index,
                    file_path=file_path,
                    files_processed=files_processed,
                    sheets_processed=sheets_processed,
                    rows_written=rows_written,
                )
            )
```

Emit `sheet_processed` after rows are appended to the output sheet:

```python
                    if on_progress:
                        on_progress(
                            MergeProgressEvent(
                                "sheet_processed",
                                total_files=total_files,
                                current_file_index=file_index,
                                file_path=file_path,
                                sheet_title=sheet.title,
                                files_processed=files_processed,
                                sheets_processed=sheets_processed,
                                rows_written=rows_written,
                            )
                        )
```

Emit `file_complete` after closing each workbook:

```python
        if on_progress:
            on_progress(
                MergeProgressEvent(
                    "file_complete",
                    total_files=total_files,
                    current_file_index=file_index,
                    file_path=file_path,
                    files_processed=files_processed,
                    sheets_processed=sheets_processed,
                    rows_written=rows_written,
                )
            )
```

Emit `writing_output` before `output_workbook.save(output)`:

```python
        if on_progress:
            on_progress(
                MergeProgressEvent(
                    "writing_output",
                    total_files=total_files,
                    current_file_index=total_files,
                    files_processed=files_processed,
                    sheets_processed=sheets_processed,
                    rows_written=rows_written,
                    output_path=output,
                )
            )
```

Emit `complete` after constructing the final separate-mode `MergeResult`:

```python
        result = MergeResult(files_processed, sheets_processed, rows_written, output)
        if on_progress:
            on_progress(
                MergeProgressEvent(
                    "complete",
                    total_files=total_files,
                    current_file_index=total_files,
                    files_processed=result.files_processed,
                    sheets_processed=result.sheets_processed,
                    rows_written=result.rows_written,
                    output_path=result.output_path,
                )
            )
        return result
```

- [ ] **Step 8: Run the focused progress tests**

Run:

```bash
python -m pytest tests/test_excel_merger.py::test_combined_mode_emits_progress_events tests/test_excel_merger.py::test_separate_mode_emits_progress_events -v
```

Expected: both tests PASS.

- [ ] **Step 9: Commit core progress contract**

```bash
git add excel_merger.py tests/test_excel_merger.py
git commit -m "feat: emit Excel merge progress events"
```

---

### Task 2: Tkinter Progress and Timestamped Logs

**Files:**
- Modify: `run.py`

- [ ] **Step 1: Add UI state for determinate progress**

In `run.py`, add this import:

```python
from datetime import datetime
```

Add `MergeProgressEvent` to the `excel_merger` import list.

In `__init__`, add:

```python
        self.progress_percent = tk.StringVar(value="0%")
```

In `_build_layout()`, change the action row progress widgets to:

```python
        progress_frame = ttk.Frame(actions)
        progress_frame.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        progress_frame.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(
            progress_frame,
            mode="determinate",
            maximum=100,
            value=0,
        )
        self.progress.grid(row=0, column=0, sticky="ew")

        ttk.Label(progress_frame, textvariable=self.progress_percent, width=5, anchor="e").grid(
            row=0, column=1, sticky="e", padx=(8, 0)
        )
```

- [ ] **Step 2: Route progress events from the worker to the UI thread**

Change `_merge_worker()` to pass a callback:

```python
            result = merge_excel_files(
                source,
                output,
                mode,
                on_progress=lambda event: self.root.after(0, self._handle_progress_event, event),
            )
```

Add this method to `ExcelMergerApp`:

```python
    def _handle_progress_event(self, event: MergeProgressEvent) -> None:
        if event.total_files > 0 and event.event_type in {"file_start", "sheet_processed", "file_complete"}:
            completed = event.current_file_index - 1
            if event.event_type == "file_complete":
                completed = event.current_file_index
            percent = int((completed / event.total_files) * 100)
        elif event.event_type == "complete":
            percent = 100
        else:
            percent = int(self.progress["value"])

        self.progress.configure(value=percent)
        self.progress_percent.set(f"{percent}%")

        status = self._format_progress_status(event)
        if status:
            self.summary.set(status)

        log_message = self._format_progress_log_message(event)
        if log_message:
            self._append_log(log_message)
```

- [ ] **Step 3: Add progress formatting helpers**

Add these methods to `ExcelMergerApp`:

```python
    def _format_progress_status(self, event: MergeProgressEvent) -> str:
        if event.event_type == "scan_complete":
            return f"Found {event.total_files} Excel workbook(s). Preparing merge..."
        if event.event_type == "file_start" and event.file_path:
            return f"Processing {event.current_file_index}/{event.total_files}: {event.file_path.name}"
        if event.event_type == "sheet_processed" and event.file_path and event.sheet_title:
            return (
                f"Processed sheet '{event.sheet_title}' from "
                f"{event.file_path.name} ({event.current_file_index}/{event.total_files})."
            )
        if event.event_type == "writing_output":
            return "Writing merged workbook..."
        if event.event_type == "complete":
            return "Merge complete."
        return self.summary.get()

    def _format_progress_log_message(self, event: MergeProgressEvent) -> str:
        if event.event_type == "scan_complete":
            return f"Found {event.total_files} Excel workbook(s)."
        if event.event_type == "file_start" and event.file_path:
            return f"Processing file {event.current_file_index}/{event.total_files}: {event.file_path.name}"
        if event.event_type == "sheet_processed" and event.file_path and event.sheet_title:
            return (
                f"Processed sheet '{event.sheet_title}' from {event.file_path.name}; "
                f"{event.sheets_processed} sheet(s), {event.rows_written} row(s) so far."
            )
        if event.event_type == "file_complete" and event.file_path:
            return f"Finished file {event.current_file_index}/{event.total_files}: {event.file_path.name}"
        if event.event_type == "writing_output":
            return "Writing output workbook..."
        return ""
```

- [ ] **Step 4: Timestamp log entries and reset progress cleanly**

Change `_append_log()`:

```python
    def _append_log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
```

Change `_set_running()` so start resets progress and stop leaves the last value visible:

```python
        if is_running:
            self.progress.configure(value=0)
            self.progress_percent.set("0%")
            self.summary.set("Preparing merge...")
        else:
            if int(float(self.progress["value"])) < 100:
                self.progress_percent.set(f"{int(float(self.progress['value']))}%")
```

Keep `_merge_failed()` logging the error through `_append_log()`, which will now add the timestamp automatically.

- [ ] **Step 5: Verify UI code compiles**

Run:

```bash
python -m py_compile run.py excel_merger.py
```

Expected: exit code 0.

- [ ] **Step 6: Commit UI progress and logs**

```bash
git add run.py
git commit -m "feat: show merge progress and timestamped logs"
```

---

### Task 3: Cross-Platform Build and Release Workflow

**Files:**
- Rename: `.github/workflows/build-windows.yml` to `.github/workflows/build-release.yml`
- Modify: `.github/workflows/build-release.yml`

- [ ] **Step 1: Rename the workflow file**

Run:

```bash
git mv .github/workflows/build-windows.yml .github/workflows/build-release.yml
```

- [ ] **Step 2: Replace workflow content**

Replace `.github/workflows/build-release.yml` with:

```yaml
name: Build and Release

on:
  push:
    branches:
      - main
    tags:
      - "v*"
  pull_request:
    branches:
      - main
  workflow_dispatch:

permissions:
  contents: write

env:
  PYTHON_VERSION: "3.10"

jobs:
  test:
    name: Test
    runs-on: ubuntu-latest

    steps:
      - name: Check out repository
        uses: actions/checkout@v6

      - name: Set up Python
        uses: actions/setup-python@v6
        with:
          python-version: ${{ env.PYTHON_VERSION }}

      - name: Install dependencies
        run: |
          python -m pip install --upgrade pip
          python -m pip install -r requirements-dev.txt

      - name: Run tests
        run: python -m pytest -v

  build:
    name: Build ${{ matrix.artifact-name }}
    needs: test
    runs-on: ${{ matrix.os }}

    strategy:
      fail-fast: false
      matrix:
        include:
          - os: windows-latest
            architecture: x86
            artifact-name: Excel-Files-Merger-win32
            asset-name: Excel-Files-Merger-win32.exe
            build-output: dist/Excel-Files-Merger.exe
          - os: windows-latest
            architecture: x64
            artifact-name: Excel-Files-Merger-win64
            asset-name: Excel-Files-Merger-win64.exe
            build-output: dist/Excel-Files-Merger.exe
          - os: ubuntu-latest
            architecture: x64
            artifact-name: Excel-Files-Merger-ubuntu
            asset-name: Excel-Files-Merger-ubuntu
            build-output: dist/Excel-Files-Merger

    steps:
      - name: Check out repository
        uses: actions/checkout@v6

      - name: Install Ubuntu GUI build dependencies
        if: runner.os == 'Linux'
        run: |
          sudo apt-get update
          sudo apt-get install -y python3-tk

      - name: Set up Python
        uses: actions/setup-python@v6
        with:
          python-version: ${{ env.PYTHON_VERSION }}
          architecture: ${{ matrix.architecture }}

      - name: Install dependencies
        run: |
          python -m pip install --upgrade pip
          python -m pip install -r requirements-dev.txt

      - name: Run tests
        run: python -m pytest -v

      - name: Build executable
        run: python -m PyInstaller --noconfirm --clean --onefile --windowed --name Excel-Files-Merger run.py

      - name: Stage release asset
        shell: bash
        run: |
          mkdir -p release-assets
          cp "${{ matrix.build-output }}" "release-assets/${{ matrix.asset-name }}"
          if [[ "${{ runner.os }}" == "Linux" ]]; then
            chmod +x "release-assets/${{ matrix.asset-name }}"
          fi

      - name: Upload executable artifact
        uses: actions/upload-artifact@v7
        with:
          name: ${{ matrix.artifact-name }}
          path: release-assets/${{ matrix.asset-name }}
          if-no-files-found: error

      - name: Upload release asset bundle
        if: startsWith(github.ref, 'refs/tags/v')
        uses: actions/upload-artifact@v7
        with:
          name: release-assets-${{ matrix.artifact-name }}
          path: release-assets/${{ matrix.asset-name }}
          if-no-files-found: error

  release:
    name: Publish GitHub Release
    needs: build
    runs-on: ubuntu-latest
    if: startsWith(github.ref, 'refs/tags/v')

    steps:
      - name: Download release assets
        uses: actions/download-artifact@v7
        with:
          pattern: release-assets-*
          path: release-assets
          merge-multiple: true

      - name: Publish release
        uses: softprops/action-gh-release@v2
        with:
          generate_release_notes: true
          files: release-assets/*
```

- [ ] **Step 3: Validate workflow structure locally**

Run:

```bash
python - <<'PY'
from pathlib import Path
path = Path(".github/workflows/build-release.yml")
text = path.read_text(encoding="utf-8")
required = [
    "name: Build and Release",
    "Excel-Files-Merger-win32.exe",
    "Excel-Files-Merger-win64.exe",
    "Excel-Files-Merger-ubuntu",
    "softprops/action-gh-release@v2",
]
missing = [item for item in required if item not in text]
if missing:
    raise SystemExit(f"missing required workflow text: {missing}")
if "\t" in text:
    raise SystemExit("workflow contains tab characters")
print("workflow structure checks passed")
PY
```

Expected: prints `workflow structure checks passed`. GitHub Actions will perform the authoritative workflow syntax validation after push.

- [ ] **Step 4: Commit workflow changes**

```bash
git add .github/workflows/build-release.yml
git add -u .github/workflows/build-windows.yml
git commit -m "ci: build Ubuntu binary and publish releases"
```

---

### Task 4: README Release Documentation

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Update feature and download sections**

Change the intro and feature/download sections to describe Windows and Ubuntu releases:

````markdown
Excel Files Merger is a small desktop utility for merging Excel workbooks on Windows and Ubuntu.

## Features

- Select a source folder and output file from the desktop UI.
- Merge `.xlsx` and `.xlsm` workbooks.
- Ignore Excel temporary files that start with `~$`.
- Choose whether to combine data into one sheet or keep workbook sheets separate.
- See merge progress, current file status, and timestamped log messages.
- Download portable Windows and Ubuntu builds from GitHub Releases.

## Download Portable Builds

1. Open this repository on GitHub.
2. Go to **Releases**.
3. Open the latest release.
4. Download the right asset:
   - `Excel-Files-Merger-win64.exe` for most Windows machines.
   - `Excel-Files-Merger-win32.exe` for older 32-bit Windows machines.
   - `Excel-Files-Merger-ubuntu` for Ubuntu x64 desktops.
5. Run the downloaded file. On Ubuntu, make it executable first if needed:

```sh
chmod +x Excel-Files-Merger-ubuntu
./Excel-Files-Merger-ubuntu
```

No Python installation is required for the downloaded executable.
````

- [ ] **Step 2: Add local Ubuntu build command**

Add this section after `Build Locally On Windows`:

````markdown
## Build Locally On Ubuntu

```sh
sudo apt-get install -y python3-tk
python -m pip install -r requirements-dev.txt
python -m PyInstaller --noconfirm --clean --onefile --windowed --name Excel-Files-Merger run.py
```

The executable is written to `dist/Excel-Files-Merger`.
````

- [ ] **Step 3: Commit README changes**

```bash
git add README.md
git commit -m "docs: document releases and progress logging"
```

---

### Task 5: Verification, Push, and Release Tag

**Files:**
- Read: repository state and workflow run status

- [ ] **Step 1: Run full local verification**

Run:

```bash
python -m pytest -v
python -m py_compile run.py excel_merger.py
git diff --check
```

Expected: pytest reports all tests passed, py_compile exits 0, and `git diff --check` exits 0.

- [ ] **Step 2: Inspect final git state**

Run:

```bash
git status -sb
git log --oneline -5
```

Expected: clean working tree on `main`, with the new progress/release commits visible.

- [ ] **Step 3: Push main**

Run:

```bash
git push origin main
```

Expected: push succeeds.

- [ ] **Step 4: Create and push release tag**

Use the next available minor version tag. If no newer tag exists, use `v1.1.0`:

```bash
git tag v1.1.0
git push origin v1.1.0
```

Expected: tag push succeeds and triggers the `Build and Release` workflow.

- [ ] **Step 5: Verify GitHub Actions and release assets**

Check the tag workflow run. It must finish successfully and the GitHub Release must contain:

- `Excel-Files-Merger-win32.exe`
- `Excel-Files-Merger-win64.exe`
- `Excel-Files-Merger-ubuntu`

If the workflow fails, inspect the failed job logs, fix the root cause, commit, move the tag to the fixed commit only after explaining the retagging action, push again, and verify the new run.

- [ ] **Step 6: Final response**

Report:

- commits pushed
- release tag
- release URL
- direct asset links for Windows 32-bit, Windows 64-bit, and Ubuntu
- local verification commands and results
