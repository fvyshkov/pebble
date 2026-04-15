#!/usr/bin/env python3
"""
Pebble — one-click launcher.

Prerequisites: Python 3.10+
Run:  python start.py

Automatically installs everything else (Node.js, npm deps, Python deps),
builds the frontend, and starts the server.
Works on Windows, macOS, Linux.
"""
import subprocess, sys, os, importlib, webbrowser, time, threading, shutil, platform

ROOT = os.path.dirname(os.path.abspath(__file__))
VENV = os.path.join(ROOT, ".venv")
FRONTEND = os.path.join(ROOT, "frontend")
DIST = os.path.join(FRONTEND, "dist")


# ── Virtual environment ─────────────────────────────────────────────

def in_venv():
    return hasattr(sys, "real_prefix") or (
        hasattr(sys, "base_prefix") and sys.base_prefix != sys.prefix
        and os.path.abspath(sys.prefix) == os.path.abspath(VENV)
    )

def venv_python():
    if sys.platform == "win32":
        return os.path.join(VENV, "Scripts", "python.exe")
    return os.path.join(VENV, "bin", "python")

def ensure_venv():
    if in_venv():
        return
    if not os.path.isdir(VENV):
        print("[Pebble] Creating virtual environment...")
        subprocess.check_call([sys.executable, "-m", "venv", VENV])
    os.execv(venv_python(), [venv_python(), __file__] + sys.argv[1:])


# ── Python dependencies ─────────────────────────────────────────────

def ensure_python_deps():
    reqs = os.path.join(ROOT, "requirements.txt")
    if not os.path.isfile(reqs):
        return
    missing = False
    for line in open(reqs):
        pkg = line.strip().split(">=")[0].split("==")[0].split(">")[0].strip()
        if not pkg or pkg.startswith("#"):
            continue
        mod = pkg.lower().replace("-", "_")
        if mod == "pyjwt": mod = "jwt"
        if mod == "python_multipart": mod = "multipart"
        try:
            importlib.import_module(mod)
        except ImportError:
            missing = True
            break
    if missing:
        print("[Pebble] Installing Python dependencies...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "-r", reqs])


# ── Node.js & frontend build ────────────────────────────────────────

def find_npm():
    """Return npm command, or None if not found."""
    npm = shutil.which("npm")
    if npm:
        return npm
    # Windows: check common install locations
    if sys.platform == "win32":
        for candidate in [
            os.path.expandvars(r"%ProgramFiles%\nodejs\npm.cmd"),
            os.path.expandvars(r"%APPDATA%\nvm\current\npm.cmd"),
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\nodejs\npm.cmd"),
        ]:
            if os.path.isfile(candidate):
                return candidate
    return None

def install_node():
    """Install Node.js automatically."""
    print("[Pebble] Node.js not found. Installing...")
    if sys.platform == "win32":
        # Download and install Node.js LTS via MSI (silent)
        import urllib.request
        url = "https://nodejs.org/dist/v22.15.0/node-v22.15.0-x64.msi"
        msi = os.path.join(os.environ.get("TEMP", "."), "node-install.msi")
        print(f"[Pebble] Downloading Node.js from {url}...")
        urllib.request.urlretrieve(url, msi)
        print("[Pebble] Installing Node.js (this may take a minute)...")
        subprocess.check_call(["msiexec", "/i", msi, "/qn", "/norestart"])
        os.remove(msi)
        # Add to PATH for current process
        node_dir = os.path.expandvars(r"%ProgramFiles%\nodejs")
        if node_dir not in os.environ["PATH"]:
            os.environ["PATH"] = node_dir + os.pathsep + os.environ["PATH"]
    elif sys.platform == "darwin":
        if shutil.which("brew"):
            subprocess.check_call(["brew", "install", "node"])
        else:
            print("[Pebble] ERROR: Install Node.js from https://nodejs.org")
            sys.exit(1)
    else:
        # Linux — try common package managers
        if shutil.which("apt-get"):
            subprocess.check_call(["sudo", "apt-get", "update", "-qq"])
            subprocess.check_call(["sudo", "apt-get", "install", "-y", "-qq", "nodejs", "npm"])
        elif shutil.which("dnf"):
            subprocess.check_call(["sudo", "dnf", "install", "-y", "nodejs", "npm"])
        else:
            print("[Pebble] ERROR: Install Node.js from https://nodejs.org")
            sys.exit(1)

def ensure_frontend():
    """Build frontend if dist/ doesn't exist."""
    if os.path.isdir(DIST) and os.listdir(DIST):
        return  # Already built

    print("[Pebble] Frontend not built — building now...")

    npm = find_npm()
    if not npm:
        install_node()
        npm = find_npm()
        if not npm:
            print("[Pebble] ERROR: npm still not found after install.")
            print("[Pebble] Please install Node.js manually: https://nodejs.org")
            sys.exit(1)

    # npm install
    node_modules = os.path.join(FRONTEND, "node_modules")
    if not os.path.isdir(node_modules):
        print("[Pebble] Installing frontend dependencies...")
        subprocess.check_call([npm, "install"], cwd=FRONTEND)

    # npm run build
    print("[Pebble] Building frontend...")
    subprocess.check_call([npm, "run", "build"], cwd=FRONTEND)
    print("[Pebble] Frontend built successfully.")


# ── Launch ───────────────────────────────────────────────────────────

def open_browser(port):
    time.sleep(2)
    webbrowser.open(f"http://localhost:{port}")

def main():
    ensure_venv()
    ensure_python_deps()
    ensure_frontend()

    port = int(os.environ.get("PORT", 8000))

    print()
    print("  ========================================")
    print(f"   Pebble: http://localhost:{port}")
    print("  ========================================")
    print()

    threading.Thread(target=open_browser, args=(port,), daemon=True).start()

    os.chdir(ROOT)
    import uvicorn
    uvicorn.run("backend.main:app", host="0.0.0.0", port=port)

if __name__ == "__main__":
    main()
