# Grayscale Trinarization Desktop App

Python / Tkinterで作成した、画像を黒・グレー・白の3クラスへ分類するデスクトップアプリです。  
日本語 / English のUI切り替えに対応し、複数画像の一括処理、2しきい値版の大津法、ノイズ除去、3値モルフォロジー、CSV / Excel / ZIP出力を備えています。

## Features

- 日本語 / English UI
- TkinterベースのシンプルなデスクトップGUI
- 複数画像の読み込み・複数選択削除
- ドラッグ＆ドロップ（`tkinterdnd2` 利用時）
- クリップボード画像の貼り付け
- 2つのしきい値によるグレースケール3値化
- 初期値
  - Lower threshold: `80`
  - Upper threshold: `180`
  - Middle gray level: `130`
- スライダーと数値入力による1刻みの調整
- 2しきい値版の大津法
- 3値専用ノイズ除去
  - Majority filter 3×3
  - Mode filter 3×3
- 黒 < グレー < 白の順序に対する3値モルフォロジー
  - Dilation
  - Erosion
  - Opening
  - Closing
- 白黒反転
- クラス表示色の変更
  - Black: RGB `(0, 0, 0)`
  - Gray: RGB `(120, 120, 120)`
  - White: RGB `(255, 255, 255)`
- 1枚目の設定を全画像へ適用してロック
- PNG / JPEG 保存
- 全画像 + 集計CSVのZIP保存
- ピクセル集計のCSV / Excel出力
- PNGの透明度維持
- JPEG出力時は透明部分を白背景へ合成

## GUI

左側に画像一覧と処理設定、右側に以下の3つのプレビューを表示します。

- Original image
- Grayscale
- Trinarization result

複雑なタブ構成を使わず、1画面で主要操作を完結できる構成です。

## Requirements

- Python 3.11+
- NumPy
- Pillow
- openpyxl
- tkinterdnd2

Tkinterは通常、Windows版Pythonに含まれています。

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

Tkinterが利用できるか確認する場合:

```powershell
python -m tkinter
```

## Build Windows executable

ビルド用依存関係をインストールします。

```powershell
pip install -r requirements-dev.txt
.\build_windows.bat
```

成功すると以下が生成されます。

```text
dist\TrinaryImageApp.exe
```

`.exe` はリポジトリ本体へ直接コミットせず、GitHub Releasesで配布する構成を想定しています。

## Project structure

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

## Main files

- `main.py` — Tkinter GUI、画像一覧、保存、CSV / Excel / ZIP出力
- `image_processing.py` — グレースケール化、3値化、大津法、ノイズ除去、モルフォロジー、画像生成
- `requirements.txt` — 実行時依存関係
- `requirements-dev.txt` — Windows EXE作成用依存関係
- `build_windows.bat` — PyInstallerによるWindows EXE作成

---

## English

A Python/Tkinter desktop application for classifying grayscale image pixels into three classes: black, gray, and white.

The application supports Japanese / English UI switching, batch image processing, a two-threshold Otsu method, trinary denoising and morphology, customizable display colors, and CSV / Excel / ZIP export.

### Run

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python main.py
```

### Build an EXE

```powershell
pip install -r requirements-dev.txt
.\build_windows.bat
```

The generated executable will be placed in `dist\TrinaryImageApp.exe`.

## License

No open-source license is included.

Source code is publicly available for portfolio and review purposes. No license for reuse, modification, or redistribution is granted.
