# Grayscale Trinarization Desktop App

A Python/Tkinter desktop application for classifying grayscale image pixels into three classes: black, gray, and white.

The application supports Japanese / English UI switching, batch image processing, a two-threshold Otsu method, trinary denoising and morphology, customizable display colors, and CSV / Excel / ZIP export.

## Features

- Japanese / English UI
- Simple Tkinter-based desktop GUI
- Load multiple images and delete multiple selected images
- Drag and drop when `tkinterdnd2` is available
- Paste images from the clipboard
- Grayscale trinarization using two thresholds
- Default values
  - Lower threshold: `80`
  - Upper threshold: `180`
  - Middle gray level: `130`
- One-step adjustment with sliders and numeric inputs
- Two-threshold Otsu method
- Trinary denoising
  - Majority filter 3x3
  - Mode filter 3x3
- Ordered trinary morphology for black < gray < white
  - Dilation
  - Erosion
  - Opening
  - Closing
- Black/white inversion
- Customizable class display colors
  - Black: RGB `(0, 0, 0)`
  - Gray: RGB `(120, 120, 120)`
  - White: RGB `(255, 255, 255)`
- Apply the first image settings to all images and lock them
- PNG / JPEG export
- ZIP export containing all processed images and a summary CSV
- Pixel statistics export to CSV / Excel
- Preserve PNG transparency
- Composite transparent areas onto a white background for JPEG output

## GUI

The left side contains the image list and processing controls. The right side displays three previews:

- Original image
- Grayscale
- Trinarization result

The main workflow is designed to fit into a single window without a complex tab structure.

## Requirements

- Python 3.11+
- NumPy
- Pillow
- openpyxl
- tkinterdnd2

Tkinter is normally included with the Windows distribution of Python.

## Installation

```powershell
git clone https://github.com/Rio4431/trinary-image-app.git
cd trinary-image-app

python -m venv .venv
.\.venv\Scripts\Activate.ps1

python -m pip install --upgrade pip
pip install -r requirements.txt
python main.py
```

To check whether Tkinter is available:

```powershell
python -m tkinter
```

## Build a Windows Executable

Install the build dependencies:

```powershell
pip install -r requirements-dev.txt
.\build_windows.bat
```

If the build succeeds, the executable will be created at:

```text
dist\TrinaryImageApp.exe
```

The intended distribution method is GitHub Releases rather than committing the `.exe` directly to the repository.

## Project Structure

```text
trinary-image-app/
├─ main.py
├─ image_processing.py
├─ requirements.txt
├─ requirements-dev.txt
├─ build_windows.bat
├─ .gitignore
└─ README.md
```

## Main Files

- `main.py` — Tkinter GUI, image list management, saving, and CSV / Excel / ZIP export
- `image_processing.py` — Grayscale conversion, trinarization, Otsu thresholding, denoising, morphology, and image generation
- `requirements.txt` — Runtime dependencies
- `requirements-dev.txt` — Dependencies for building the Windows executable
- `build_windows.bat` — PyInstaller-based Windows executable build script

## License

No open-source license is included.

Source code is publicly available for portfolio and review purposes. No license for reuse, modification, or redistribution is granted.
