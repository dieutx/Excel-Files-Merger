# Excel Merger Refresh Design

## Goal

Refresh the existing Excel Files Merger into a small, polished Windows desktop utility that can be distributed as portable 32-bit and 64-bit executables through GitHub Actions.

## Current State

The repository contains a single Tkinter script, `run.py`. It lets a user choose a folder and output file, then concatenates all `.xlsx` files into one worksheet. The current implementation has several gaps:

- The UI is basic and does not show file counts, progress, or merge status.
- Empty folders and unreadable files produce generic errors.
- Temporary Excel lock files are not filtered.
- Only `.xlsx` files are considered.
- The README references the wrong repository and script name.
- There is no dependency file, test suite, packaging script, or CI workflow.

## Chosen Approach

Use Tkinter with themed `ttk` widgets instead of adding a heavier desktop UI dependency. This keeps the app easy to package with PyInstaller and reduces risk for portable Windows executables.

The UI will use a compact dashboard layout:

- Source folder picker.
- Output file picker.
- File count and status summary.
- Merge mode selector:
  - Combine all files into one worksheet.
  - Keep source sheets separate in the output workbook.
- Merge button with disabled/running state.
- Status/log area with clear success and error messages.

## Merge Behavior

The app will scan the selected folder for `.xlsx` and `.xlsm` files, excluding temporary lock files such as `~$report.xlsx`.

In "combine into one worksheet" mode, each sheet read from each source workbook is appended into one output sheet. The output includes source metadata columns so users can trace rows back to the original file and sheet.

In "keep sheets separate" mode, each source sheet is written to its own sheet in the output workbook. Sheet names are sanitized, deduplicated, and capped to Excel's 31-character sheet name limit.

The app will show friendly errors for missing selections, empty folders, unreadable workbooks, write failures, or invalid output paths.

## Packaging

Add project files for reproducible builds:

- `requirements.txt` for runtime dependencies.
- `requirements-dev.txt` for build/test dependencies.
- PyInstaller build instructions or spec file.
- GitHub Actions workflow that builds portable Windows artifacts for 32-bit and 64-bit Python.

The workflow will upload `.exe` artifacts from each matrix target so the user can download them from GitHub without installing Python or dependencies on Windows.

## Testing

Add focused tests around file discovery, sheet-name sanitization, and merge behavior. UI smoke testing will stay lightweight because Tkinter GUI automation is fragile in headless CI.

Before pushing, run the available Python tests and at least one local packaging or syntax validation command. The Windows `.exe` build itself will be validated by GitHub Actions after push.

## Out Of Scope

- Drag-and-drop support.
- Installer generation.
- Mac or Linux desktop executables.
- Advanced Excel formatting preservation beyond worksheet data.
