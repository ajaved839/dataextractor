# XLXS Data Extractor

Offline desktop app for extracting hospital discharge summary `.xlsx` files into one clean Excel sheet.

The app is designed for templates like the provided discharge summary screenshot. It reads real Excel cell text, not images, and creates one row per uploaded file.

## Features

- Runs locally on Windows and macOS.
- No cloud service, no paid API, and no patient data upload.
- Select multiple `.xlsx` files.
- Export a combined `.xlsx` on the Desktop with columns for MR number, patient name, age, contact, diagnosis, procedure, operative notes, medicines, and follow-up details.
- Shows warnings when expected fields are missing.

## Install For Development

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -e ".[dev]"
```

On Windows PowerShell:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e ".[dev]"
```

## Run The Desktop App

```bash
xlxs-dataextractor
```

Or:

```bash
python -m xlxs_dataextractor.app
```

## Run From Command Line

```bash
python -m xlxs_dataextractor.cli "file1.xlsx" "file2.xlsx"
```

By default, the output is saved as `Desktop/extracted_discharge_data.xlsx`. Use `-o "output.xlsx"` only when you want a different location.

## Test

```bash
pytest
```

## Build Standalone App

Build on the same operating system you want to distribute for. Build macOS on macOS and Windows `.exe` on Windows.

macOS:

```bash
pyinstaller --name "XLXS Data Extractor" --windowed --onefile --paths src packaging/pyinstaller_entry.py
```

Windows:

```powershell
pyinstaller --name "XLXS Data Extractor" --windowed --onefile --paths src packaging\pyinstaller_entry.py
```

The built app will be created inside the `dist` folder.

## Build Windows EXE From macOS

You cannot build the Windows `.exe` directly on macOS with PyInstaller, but this project includes a free GitHub Actions workflow that builds it on a Windows runner.

Steps:

1. Create a GitHub repository.
2. Push this project to GitHub.
3. Open the repository on GitHub.
4. Go to `Actions`.
5. Select `Build Windows EXE`.
6. Click `Run workflow`.
7. After it finishes, open the workflow run and download the artifact named `XLXS-Data-Extractor-Windows`.

The downloaded artifact contains `XLXS Data Extractor.exe`.

## Accuracy Notes

This app extracts text that already exists in Excel cells. If a workbook contains only a screenshot or scanned image, exact extraction is not possible with this zero-cost version. OCR can be added later, but medical OCR should always include manual review.
