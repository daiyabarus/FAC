"""
build.py - Script untuk membuild FAC Report Generator menjadi executable menggunakan PyInstaller
Usage: python build.py
"""

import os
import shutil
import sys
from pathlib import Path

import PyInstaller.__main__

# =============================================================================
# KONFIGURASI
# =============================================================================

APP_NAME = "FAC-GR"
MAIN_SCRIPT = "main.py"

ICON_FILE = "app.ico"


INCLUDE_DIRS = [
    "assets",

]

EXCLUDED_FILES = [
    "datatemplate.xlsx",
    "config/kpi_config.json",
    "config/band_mapping.json",
]

ONEFILE = False
CLEAN_FIRST = True


def clean_build():
    folders = ["build", "dist"]
    files = [f"{APP_NAME}.spec"]

    for folder in folders:
        if Path(folder).exists():
            print(f"Menghapus folder: {folder}")
            shutil.rmtree(folder)

    for file in files:
        if Path(file).exists():
            print(f"Menghapus file: {file}")
            os.remove(file)


def collect_data_args():
    data_args = []

    for dir_path in INCLUDE_DIRS:
        if Path(dir_path).exists():
            data_args.append(f"{dir_path}{os.pathsep}{dir_path}")
        else:
            print(f"Warning: Folder tidak ditemukan untuk include: {dir_path}")

    return data_args


def build_executable():
    if not Path(MAIN_SCRIPT).exists():
        print(f"Error: File utama tidak ditemukan: {MAIN_SCRIPT}")
        sys.exit(1)

    if not Path(ICON_FILE).exists():
        print(
            f"Warning: Icon tidak ditemukan: {ICON_FILE} → menggunakan icon default")

    args = [
        MAIN_SCRIPT,
        "--name=" + APP_NAME,
        "--windowed",
        "--noconfirm",
    ]

    # Tambahkan icon
    if Path(ICON_FILE).exists():
        args.append(f"--icon={ICON_FILE}")

    # Mode build
    args.append("--onefile" if ONEFILE else "--onedir")

    for data in collect_data_args():
        args.append(f"--add-data={data}")

    hidden_imports = [
        "PyQt6.QtCore",
        "PyQt6.QtGui",
        "PyQt6.QtWidgets",
        "pandas",
        "openpyxl",
        "numpy",
    ]
    for hi in hidden_imports:
        args.append(f"--hidden-import={hi}")

    print("\nMemulai PyInstaller build...")
    print("Argumen:", " ".join(args))
    print("=" * 80)

    PyInstaller.__main__.run(args)

    print("=" * 80)
    exe_path = f"dist/{APP_NAME}.exe" if ONEFILE else f"dist/{APP_NAME}/{APP_NAME}.exe"
    print("Build selesai!")
    print(f"Executable: {exe_path}")
    print("\nFile yang TIDAK di-pack (bisa di-update kapan saja):")
    for f in EXCLUDED_FILES:
        print(f"   • {f}")
    print("\nLetakkan file-file di atas di folder yang sama dengan .exe saat distribusi.")


if __name__ == "__main__":
    print(f"Membuild {APP_NAME} executable...\n")

    if CLEAN_FIRST:
        clean_build()

    build_executable()
