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

NODE_MIN = 18
NODE_MAX = 22  # LTS versions: 18, 20, 22 — avoid odd/bleeding-edge releases

def _find_node():
    """Return path to node executable, checking common Windows paths too."""
    node = shutil.which("node")
    if node:
        return node
    if sys.platform == "win32":
        for d in [
            os.path.expandvars(r"%ProgramFiles%\nodejs"),
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\nodejs"),
        ]:
            candidate = os.path.join(d, "node.exe")
            if os.path.isfile(candidate):
                # Add to PATH so child processes (npm) also see it
                if d not in os.environ["PATH"]:
                    os.environ["PATH"] = d + os.pathsep + os.environ["PATH"]
                return candidate
    return None

def get_node_version():
    """Return (major_version, node_path) or (None, None)."""
    node = _find_node()
    if not node:
        return None, None
    try:
        out = subprocess.check_output([node, "--version"], text=True).strip()
        return int(out.lstrip("v").split(".")[0]), node
    except Exception:
        return None, None

def _refresh_win_path():
    """On Windows, re-read the system/user PATH from the registry so the
    current process sees freshly-installed programs (winget doesn't update
    the running shell's PATH)."""
    if sys.platform != "win32":
        return
    try:
        import winreg
        parts = []
        for root, sub in [
            (winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment"),
            (winreg.HKEY_CURRENT_USER, r"Environment"),
        ]:
            try:
                with winreg.OpenKey(root, sub) as key:
                    val, _ = winreg.QueryValueEx(key, "Path")
                    parts.append(val)
            except FileNotFoundError:
                pass
        if parts:
            os.environ["PATH"] = os.pathsep.join(parts)
    except Exception:
        pass

def install_node():
    """Install Node.js LTS automatically."""
    print("[Pebble] Installing Node.js LTS...")
    if sys.platform == "win32":
        if shutil.which("winget"):
            # Uninstall any incompatible version first
            print("[Pebble] Removing old Node.js (if any)...")
            subprocess.run(["winget", "uninstall", "OpenJS.NodeJS",
                            "--accept-source-agreements", "--silent"],
                           capture_output=True)
            subprocess.run(["winget", "uninstall", "OpenJS.NodeJS.LTS",
                            "--accept-source-agreements", "--silent"],
                           capture_output=True)
            print("[Pebble] Installing Node.js LTS via winget...")
            subprocess.check_call(["winget", "install", "OpenJS.NodeJS.LTS",
                                   "--accept-source-agreements",
                                   "--accept-package-agreements"])
        else:
            print("[Pebble] ERROR: winget not found. Install Node.js LTS manually: https://nodejs.org")
            sys.exit(1)
        # Refresh PATH from registry so we see the new install
        _refresh_win_path()
        # Also add default location explicitly
        node_dir = os.path.expandvars(r"%ProgramFiles%\nodejs")
        if os.path.isdir(node_dir) and node_dir not in os.environ["PATH"]:
            os.environ["PATH"] = node_dir + os.pathsep + os.environ["PATH"]
    elif sys.platform == "darwin":
        if shutil.which("brew"):
            subprocess.check_call(["brew", "install", "node@20"])
        else:
            print("[Pebble] ERROR: Install Node.js LTS from https://nodejs.org")
            sys.exit(1)
    else:
        # Linux — use NodeSource for proper LTS version
        if shutil.which("apt-get"):
            subprocess.check_call(["sudo", "apt-get", "update", "-qq"])
            subprocess.check_call(["sudo", "apt-get", "install", "-y", "-qq",
                                   "ca-certificates", "curl", "gnupg"])
            subprocess.check_call(
                "curl -fsSL https://deb.nodesource.com/setup_20.x | sudo -E bash -",
                shell=True)
            subprocess.check_call(["sudo", "apt-get", "install", "-y", "-qq", "nodejs"])
        elif shutil.which("dnf"):
            subprocess.check_call(
                "curl -fsSL https://rpm.nodesource.com/setup_20.x | sudo bash -",
                shell=True)
            subprocess.check_call(["sudo", "dnf", "install", "-y", "nodejs"])
        else:
            print("[Pebble] ERROR: Install Node.js LTS from https://nodejs.org")
            sys.exit(1)

def ensure_node():
    """Make sure a compatible Node.js is available; install if needed."""
    ver, _ = get_node_version()
    need_install = ver is None or ver < NODE_MIN or ver > NODE_MAX
    if need_install:
        if ver is not None:
            print(f"[Pebble] Node.js v{ver} detected — need v{NODE_MIN}-{NODE_MAX} (LTS).")
        install_node()
        # Wipe node_modules — native modules built for wrong Node ABI
        node_modules = os.path.join(FRONTEND, "node_modules")
        if os.path.isdir(node_modules):
            print("[Pebble] Removing old node_modules (incompatible)...")
            shutil.rmtree(node_modules, ignore_errors=True)
    # Verify node actually works
    ver2, node_path = get_node_version()
    if ver2 is None or ver2 < NODE_MIN or ver2 > NODE_MAX:
        print(f"[Pebble] ERROR: Node.js LTS not available (found: v{ver2}).")
        print("[Pebble] Please install Node.js LTS (v20) manually: https://nodejs.org")
        sys.exit(1)
    print(f"[Pebble] Using Node.js v{ver2} ({node_path})")
    npm = find_npm()
    if not npm:
        print("[Pebble] ERROR: npm not found.")
        print("[Pebble] Please install Node.js LTS manually: https://nodejs.org")
        sys.exit(1)
    return npm

def ensure_frontend():
    """Build frontend if dist/ doesn't exist."""
    if os.path.isdir(DIST) and os.listdir(DIST):
        return  # Already built

    print("[Pebble] Frontend not built — building now...")

    npm = ensure_node()

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
