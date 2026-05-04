#!/usr/bin/env python3
"""
Build a distributable Pebble release zip for Windows.

Usage:
    python build_release.py

Output:
    pebble-release.zip — ready to distribute.
    Contains everything needed: backend, pre-built frontend, installer.

The end user just extracts the zip and double-clicks Install.bat.
"""
import subprocess, sys, os, shutil, zipfile, glob

ROOT = os.path.dirname(os.path.abspath(__file__))
FRONTEND = os.path.join(ROOT, "frontend")
DIST = os.path.join(FRONTEND, "dist")
WHEELS = os.path.join(ROOT, "dist-wheels")
OUTPUT = os.path.join(ROOT, "pebble-release.zip")

# Files/dirs to include in the release
INCLUDE = [
    "backend",
    "frontend/dist",
    "installer/install.ps1",
    "installer/Install.bat",
    "Pebble.bat",
    "requirements.txt",
    "start.py",
    "start.bat",
]

# Files/patterns to exclude
EXCLUDE_PATTERNS = {
    "__pycache__",
    ".pyc",
    ".pyo",
    ".DS_Store",
    "thumbs.db",
    "venv",
    ".venv",
    "node_modules",
    ".git",
    "tests",
    ".pytest_cache",
}


EXCLUDE_DIRS = {"venv", ".venv", "node_modules", ".git", "tests", ".pytest_cache", "__pycache__"}
EXCLUDE_EXTS = {".pyc", ".pyo"}
EXCLUDE_FILES = {".DS_Store", "thumbs.db"}


def should_exclude(path):
    parts = path.replace("\\", "/").split("/")
    for part in parts:
        if part in EXCLUDE_DIRS:
            return True
    basename = parts[-1] if parts else ""
    if basename in EXCLUDE_FILES:
        return True
    _, ext = os.path.splitext(basename)
    if ext in EXCLUDE_EXTS:
        return True
    return False


def ensure_frontend_built():
    if os.path.isdir(DIST) and os.listdir(DIST):
        print("[build] Frontend already built.")
        return
    print("[build] Building frontend...")
    npm = shutil.which("npm")
    if not npm:
        print("[build] ERROR: npm not found. Install Node.js first.")
        sys.exit(1)
    subprocess.check_call([npm, "install"], cwd=FRONTEND)
    subprocess.check_call([npm, "run", "build"], cwd=FRONTEND)
    print("[build] Frontend built.")


def collect_wheels():
    """Find Windows pebble_calc wheel(s) for bundling."""
    if not os.path.isdir(WHEELS):
        return []
    win = sorted(glob.glob(os.path.join(WHEELS, "pebble_calc-*-win_amd64.whl")))
    if not win:
        win = sorted(glob.glob(os.path.join(WHEELS, "pebble_calc-*.whl")))
    return win


def build_zip():
    ensure_frontend_built()

    wheels = collect_wheels()
    if not wheels:
        print("[build] WARNING: no pebble_calc wheel found in dist-wheels/ — Windows install will fail.")
    else:
        for w in wheels:
            print(f"[build] Bundling wheel: {os.path.basename(w)}")

    print(f"[build] Creating {OUTPUT}...")

    with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in INCLUDE:
            full_path = os.path.join(ROOT, item)
            if os.path.isfile(full_path):
                arcname = os.path.join("pebble", item)
                if not should_exclude(arcname):
                    zf.write(full_path, arcname)
                    print(f"  + {arcname}")
            elif os.path.isdir(full_path):
                for dirpath, dirnames, filenames in os.walk(full_path):
                    # Skip excluded directories
                    dirnames[:] = [d for d in dirnames if d not in EXCLUDE_DIRS]
                    for fname in filenames:
                        filepath = os.path.join(dirpath, fname)
                        relpath = os.path.relpath(filepath, ROOT)
                        arcname = os.path.join("pebble", relpath)
                        if not should_exclude(arcname):
                            zf.write(filepath, arcname)

        for w in wheels:
            arcname = os.path.join("pebble", "wheels", os.path.basename(w))
            zf.write(w, arcname)
            print(f"  + {arcname}")

        # Add a top-level Install.bat that points into the pebble dir
        launcher = (
            '@echo off\r\n'
            'chcp 65001 >nul 2>&1\r\n'
            'cd /d "%~dp0pebble\\installer"\r\n'
            'call Install.bat\r\n'
        )
        zf.writestr("Install.bat", launcher)

    size_mb = os.path.getsize(OUTPUT) / (1024 * 1024)
    print(f"[build] Done! {OUTPUT} ({size_mb:.1f} MB)")
    print()
    print("Distribution:")
    print("  Push to main — GitHub Actions will build & publish automatically.")
    print("  Users download from: https://github.com/fvyshkov/pebble/releases/latest")
    print()
    print("  Or one-liner install:")
    print('  powershell -c "irm https://raw.githubusercontent.com/fvyshkov/pebble/main/installer/install.ps1 | iex"')


if __name__ == "__main__":
    build_zip()
