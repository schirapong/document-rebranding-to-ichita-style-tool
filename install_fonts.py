#!/usr/bin/env python3
"""
Install Ichita brand fonts (Aeonik + Bai Jamjuree) to the system font directory.

Aeonik: converted from bundled woff2 files in Aeonik-Essentials-Web/
Bai Jamjuree: downloaded from Google Fonts (OFL licensed)

Usage:
    python3 install_fonts.py              # install all fonts
    python3 install_fonts.py --check      # check which fonts are installed
"""

import argparse
import glob
import os
import platform
import sys
import urllib.request
import zipfile
import tempfile


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
AEONIK_DIR = os.path.join(SCRIPT_DIR, "Aeonik-Essentials-Web")

# Google Fonts download URL for Bai Jamjuree
BAI_JAMJUREE_URL = "https://fonts.google.com/download?family=Bai+Jamjuree"


def get_font_dir():
    """Return the user font directory for the current platform."""
    system = platform.system()
    if system == "Darwin":
        return os.path.expanduser("~/Library/Fonts")
    elif system == "Linux":
        return os.path.expanduser("~/.local/share/fonts")
    elif system == "Windows":
        return os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")
    else:
        print(f"Unsupported platform: {system}")
        sys.exit(1)


def check_fonts(font_dir):
    """Check which brand fonts are already installed."""
    if not os.path.isdir(font_dir):
        print(f"Font directory does not exist: {font_dir}")
        return False, False

    files = [f.lower() for f in os.listdir(font_dir)]
    has_aeonik = any("aeonik" in f for f in files)
    has_bai = any("baijamjuree" in f or "bai-jamjuree" in f or "bai_jamjuree" in f for f in files)

    return has_aeonik, has_bai


def install_aeonik(font_dir):
    """Convert Aeonik woff2 → OTF and install to font directory."""
    woff2_files = glob.glob(os.path.join(AEONIK_DIR, "*.woff2"))
    if not woff2_files:
        print(f"  No woff2 files found in {AEONIK_DIR}")
        return False

    try:
        from fontTools.ttLib import TTFont
    except ImportError:
        print("  fonttools not installed. Run: pip3 install fonttools brotli")
        return False

    os.makedirs(font_dir, exist_ok=True)
    installed = 0

    for woff2 in sorted(woff2_files):
        basename = os.path.splitext(os.path.basename(woff2))[0]
        out_path = os.path.join(font_dir, f"{basename}.otf")

        if os.path.exists(out_path):
            print(f"  Already installed: {basename}.otf")
            installed += 1
            continue

        font = TTFont(woff2)
        font.flavor = None
        font.save(out_path)
        print(f"  Installed: {basename}.otf")
        installed += 1

    print(f"  Aeonik: {installed} font files in {font_dir}")
    return True


def install_bai_jamjuree(font_dir):
    """Download Bai Jamjuree from Google Fonts and install."""
    os.makedirs(font_dir, exist_ok=True)

    print("  Downloading Bai Jamjuree from Google Fonts...")
    try:
        with tempfile.NamedTemporaryFile(suffix=".zip", delete=False) as tmp:
            tmp_path = tmp.name
            urllib.request.urlretrieve(BAI_JAMJUREE_URL, tmp_path)

        installed = 0
        with zipfile.ZipFile(tmp_path, 'r') as zf:
            for name in zf.namelist():
                if name.endswith('.ttf'):
                    basename = os.path.basename(name)
                    out_path = os.path.join(font_dir, basename)
                    if os.path.exists(out_path):
                        print(f"  Already installed: {basename}")
                    else:
                        with zf.open(name) as src, open(out_path, 'wb') as dst:
                            dst.write(src.read())
                        print(f"  Installed: {basename}")
                    installed += 1

        os.unlink(tmp_path)
        print(f"  Bai Jamjuree: {installed} font files in {font_dir}")
        return True

    except Exception as e:
        print(f"  Failed to download Bai Jamjuree: {e}")
        print("  You can manually download from: https://fonts.google.com/specimen/Bai+Jamjuree")
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        return False


def main():
    parser = argparse.ArgumentParser(
        description="Install Ichita brand fonts (Aeonik + Bai Jamjuree)")
    parser.add_argument("--check", action="store_true",
                        help="Check which fonts are installed without installing")
    parser.add_argument("--font-dir", type=str, default=None,
                        help="Override font installation directory")
    args = parser.parse_args()

    font_dir = args.font_dir or get_font_dir()
    print(f"Font directory: {font_dir}")

    has_aeonik, has_bai = check_fonts(font_dir)

    if args.check:
        print(f"  Aeonik:       {'installed' if has_aeonik else 'NOT installed'}")
        print(f"  Bai Jamjuree: {'installed' if has_bai else 'NOT installed'}")
        if has_aeonik and has_bai:
            print("\nAll brand fonts are installed.")
        else:
            print(f"\nRun 'python3 install_fonts.py' to install missing fonts.")
        return

    print()
    if not has_aeonik:
        print("Installing Aeonik...")
        install_aeonik(font_dir)
    else:
        print("Aeonik: already installed")

    print()
    if not has_bai:
        print("Installing Bai Jamjuree...")
        install_bai_jamjuree(font_dir)
    else:
        print("Bai Jamjuree: already installed")

    # Refresh font cache on Linux
    if platform.system() == "Linux":
        print("\nRefreshing font cache...")
        os.system("fc-cache -f")

    print("\nDone. Fonts are ready to use.")


if __name__ == "__main__":
    main()
