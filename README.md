# Excel Files Merger

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

Ubuntu builds require a compatible graphical desktop environment with the needed Tk/display libraries available.

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

## Build Locally On Ubuntu

```sh
sudo apt-get install -y python3-tk
python -m pip install -r requirements-dev.txt
python -m PyInstaller --noconfirm --clean --onefile --windowed --name Excel-Files-Merger run.py
```

The executable is written to `dist/Excel-Files-Merger`.

## Development Commands

```sh
python -m pip install -r requirements-dev.txt
python -m pytest -v
python -m py_compile run.py excel_merger.py
```

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
