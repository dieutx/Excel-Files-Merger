# Progress Logging and Cross-Platform Release Design

## Goal

Improve the desktop app so users can see merge progress in real time, inspect useful log messages when something goes wrong, and download portable builds for Windows and Ubuntu from GitHub Releases.

## Current State

The app already has a Tkinter dashboard, an indeterminate progress bar, and a log text area. The merge logic in `excel_merger.py` returns a final `MergeResult`, but it does not emit progress while work is happening. GitHub Actions currently builds Windows 32-bit and 64-bit artifacts only, uploaded from workflow runs rather than attached to a versioned release.

## User Experience

During a merge, the progress bar will switch to determinate mode and show real completion progress based on files processed. The status line will show the current operation, such as scanning files, processing `3/12`, writing output, or completed.

The log area will receive timestamped messages for key events:

- merge start and selected mode
- discovered file count
- each file start and completion
- each non-empty sheet processed
- output writing
- success summary
- failure details

If a workbook fails to load or a sheet fails while being processed, the UI will show the failure in the summary, keep the details in the log, and display an error dialog. The app will continue to prevent closing while a merge worker is active.

## Merge Logic

`excel_merger.merge_excel_files()` will accept an optional progress callback. The callback receives structured progress events rather than formatted UI text, so the core merge logic stays independent from Tkinter.

The event model will include enough context for the UI and tests:

- event type, such as `scan_complete`, `file_start`, `sheet_processed`, `file_complete`, `writing_output`, `complete`
- current file path when relevant
- current file index and total file count
- sheet title when relevant
- running counts for processed files, processed sheets, and written rows

Progress will be based on file count because total sheet and row counts are not known without pre-reading every workbook. This avoids doubling the work on large folders while still giving users a useful visible estimate.

## UI Changes

`run.py` will route progress events from the worker thread back to Tkinter through `root.after()`. The UI thread will update:

- the determinate progress bar value
- a percentage label
- the summary/status text
- the log text area

The log will remain visible in the main window. It will auto-scroll to the latest message and include timestamps so users can report or screenshot errors more easily.

## Tests

The existing merge behavior tests will remain. New tests will cover progress events for both merge modes:

- callbacks are emitted during successful combined merges
- callbacks are emitted during successful separate-sheet merges
- file progress reaches the expected total
- the completion event contains final counts matching `MergeResult`

UI behavior will stay thin and mostly untested because it is Tkinter state plumbing. The core progress contract will be tested in `tests/test_excel_merger.py`.

## Build and Release

The GitHub Actions workflow will be renamed from Windows-only to a cross-platform build workflow. It will:

- run tests on Ubuntu
- build Windows 32-bit portable `.exe`
- build Windows 64-bit portable `.exe`
- build Ubuntu x64 portable binary
- upload all build artifacts for normal workflow runs
- publish a GitHub Release when a tag matching `v*` is pushed

Release asset names will be stable:

- `Excel-Files-Merger-win32.exe`
- `Excel-Files-Merger-win64.exe`
- `Excel-Files-Merger-ubuntu`

The Ubuntu build will use PyInstaller on `ubuntu-latest`. It will be portable for a compatible Ubuntu/Linux desktop environment, but Linux GUI portability depends on system GUI libraries such as Tk and the display environment already available on the target machine.

## Documentation

`README.md` will describe:

- progress and log visibility
- downloading builds from GitHub Releases
- which asset to choose for Windows 32-bit, Windows 64-bit, and Ubuntu
- local build commands for Windows and Ubuntu

## Out of Scope

This change will not add cancellation, recursive folder scanning, drag-and-drop, automatic updates, installers, code signing, or packaging for other Linux distributions.
