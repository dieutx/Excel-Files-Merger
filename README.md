# Excel Files Merger

Excel Files Merger is a small Windows desktop utility for merging Excel workbooks.

## Features

- Select a source folder and output file from the desktop UI.
- Merge `.xlsx` and `.xlsm` workbooks.
- Ignore Excel temporary files that start with `~$`.
- Choose whether to combine data into one sheet or keep workbook sheets separate.
- Download portable Windows executables built by GitHub Actions.

## Download Portable Windows Builds

1. Open this repository on GitHub.
2. Go to **Actions**.
3. Open the latest **Build Windows EXE** workflow run.
4. Download **Excel-Files-Merger-win64** for most Windows machines, or **Excel-Files-Merger-win32** for older 32-bit Windows machines.
5. Extract the artifact and run `Excel-Files-Merger.exe`.

No Python installation is required for the downloaded executable.

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

The executable is written to `dist/Excel-Files-Merger.exe`.

## Development Commands

```sh
python -m pip install -r requirements-dev.txt
python -m pytest -v
python -m py_compile run.py excel_merger.py
```

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
