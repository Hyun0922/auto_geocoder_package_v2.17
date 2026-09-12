# -*- coding: utf-8 -*-
"""Windows EXE build helper for Auto Geocoder v2.17.

Called by the ASCII-only build_exe.cmd file so Windows CMD never has to
parse Korean text or Unicode paths during the build itself.

If icon.png exists beside this script, it is automatically converted to a
multi-size icon.ico before PyInstaller runs.
"""
from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "auto_geocoder.py"
REQUIREMENTS = ROOT / "requirements.txt"
RELEASE_DIR = ROOT / "release"
DIST_DIR = ROOT / "dist"
ASCII_EXE = DIST_DIR / "AutoGeocoder.exe"
FINAL_EXE = RELEASE_DIR / "자동 지오코더.exe"
ICON_PNG = ROOT / "icon.png"
ICON_ICO = ROOT / "icon.ico"
ICON_SIZES = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]


def run(args: list[str], label: str) -> None:
    print()
    print("=" * 68)
    print(label)
    print("=" * 68)
    subprocess.run(args, cwd=ROOT, check=True)


def convert_png_to_ico() -> Path | None:
    """Convert icon.png to a real multi-resolution Windows icon.ico."""
    if not ICON_PNG.exists():
        if ICON_ICO.exists():
            print(f"Custom icon: using existing {ICON_ICO.name}")
            return ICON_ICO
        print("Custom icon: icon.png/icon.ico not found (default EXE icon will be used)")
        return None

    from PIL import Image

    with Image.open(ICON_PNG) as source:
        image = source.convert("RGBA")
        # Keep aspect ratio and center it on a square transparent canvas.
        side = max(image.size)
        square = Image.new("RGBA", (side, side), (0, 0, 0, 0))
        x = (side - image.width) // 2
        y = (side - image.height) // 2
        square.paste(image, (x, y), image)
        square.save(ICON_ICO, format="ICO", sizes=ICON_SIZES)

    print(f"Custom icon: converted {ICON_PNG.name} -> {ICON_ICO.name}")
    print("Icon sizes: 16, 32, 48, 64, 128, 256 px")
    return ICON_ICO


def main() -> int:
    print("Auto Geocoder v2.17 Windows EXE build")
    print(f"Python: {sys.executable}")
    print(f"Project: {ROOT}")

    if not SOURCE.exists():
        raise FileNotFoundError(f"Missing source file: {SOURCE.name}")
    if not REQUIREMENTS.exists():
        raise FileNotFoundError(f"Missing requirements file: {REQUIREMENTS.name}")

    run(
        [sys.executable, "-m", "pip", "install", "-r", str(REQUIREMENTS)],
        "[1/4] Installing required packages",
    )

    print()
    print("=" * 68)
    print("[2/4] Preparing application icon")
    print("=" * 68)
    icon = convert_png_to_ico()

    pyinstaller_args = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name",
        "AutoGeocoder",
        "--collect-all",
        "pyproj",
    ]

    if icon is not None:
        sep = ";" if sys.platform.startswith("win") else ":"
        pyinstaller_args += [
            "--icon", str(icon),
            "--add-data", f"{icon}{sep}.",
        ]

    pyinstaller_args.append(str(SOURCE))
    run(pyinstaller_args, "[3/4] Building executable")

    if not ASCII_EXE.exists():
        raise FileNotFoundError("PyInstaller finished, but dist/AutoGeocoder.exe was not found.")

    print()
    print("=" * 68)
    print("[4/4] Preparing release folder")
    print("=" * 68)
    RELEASE_DIR.mkdir(exist_ok=True)
    shutil.copy2(ASCII_EXE, FINAL_EXE)

    config = ROOT / "config.json"
    example = ROOT / "config.example.json"
    if config.exists():
        shutil.copy2(config, RELEASE_DIR / "config.json")
    elif example.exists():
        shutil.copy2(example, RELEASE_DIR / "config.json")

    readme = ROOT / "README_자동지오코더.md"
    if readme.exists():
        shutil.copy2(readme, RELEASE_DIR / readme.name)

    print()
    print("BUILD SUCCESS")
    print(f"EXE: {FINAL_EXE}")
    print("End-user files: EXE + config.json (+ README recommended)")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except subprocess.CalledProcessError as exc:
        print()
        print("BUILD FAILED")
        print(f"A command returned exit code {exc.returncode}.")
        raise SystemExit(exc.returncode)
    except Exception as exc:
        print()
        print("BUILD FAILED")
        print(f"{type(exc).__name__}: {exc}")
        raise SystemExit(1)
