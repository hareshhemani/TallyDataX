"""
Build script to compile TallyDataX into a standalone Windows .exe
Using PyInstaller 6.x
"""

import os
import sys
import subprocess

import zipfile
import shutil

def build():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(base_dir)

    print("=======================================================")
    print("  Building TallyDataX Standalone Executable (.exe)...")
    print("=======================================================")

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name=TallyDataX_App",
        "--onefile",
        "--clean",
        "--icon=app_icon.ico",
        "--add-data=index.html;.",
        "--add-data=app_icon.ico;.",
        "--add-data=app_icon.png;.",
        "--hidden-import=uvicorn.logging",
        "--hidden-import=uvicorn.loops",
        "--hidden-import=uvicorn.loops.auto",
        "--hidden-import=uvicorn.protocols",
        "--hidden-import=uvicorn.protocols.http",
        "--hidden-import=uvicorn.protocols.http.auto",
        "--hidden-import=uvicorn.protocols.http.h11_impl",
        "--hidden-import=uvicorn.protocols.websockets",
        "--hidden-import=uvicorn.protocols.websockets.auto",
        "--hidden-import=uvicorn.lifespan",
        "--hidden-import=uvicorn.lifespan.on",
        "--hidden-import=uvicorn.lifespan.off",
        "--hidden-import=multipart",
        "--hidden-import=multipart.multipart",
        "--hidden-import=openpyxl",
        "--hidden-import=xlsxwriter",
        "--hidden-import=requests",
        "--hidden-import=tally_client",
        "--hidden-import=report_engine",
        "app.py"
    ]

    print("Executing command:")
    print(" ".join(cmd))
    print("-------------------------------------------------------")

    result = subprocess.run(cmd)

    # Clean temporary build directory and spec file
    build_dir = os.path.join(base_dir, "build")
    spec_file = os.path.join(base_dir, "TallyDataX_App.spec")
    if os.path.exists(build_dir):
        shutil.rmtree(build_dir, ignore_errors=True)
    if os.path.exists(spec_file):
        try:
            os.remove(spec_file)
        except OSError:
            pass

    if result.returncode == 0:
        dist_exe = os.path.join(base_dir, "dist", "TallyDataX_App.exe")
        if os.path.exists(dist_exe):
            size_mb = os.path.getsize(dist_exe) / (1024 * 1024)

            # Create updated zip
            dist_zip = os.path.join(base_dir, "dist", "TallyDataX_App.zip")
            with zipfile.ZipFile(dist_zip, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.write(dist_exe, arcname="TallyDataX_App.exe")

            zip_size_mb = os.path.getsize(dist_zip) / (1024 * 1024)

            print("=======================================================")
            print(f"  BUILD SUCCESSFUL!")
            print(f"  Output File: {dist_exe} ({size_mb:.2f} MB)")
            print(f"  Zip Archive: {dist_zip} ({zip_size_mb:.2f} MB)")
            print("=======================================================")
            return True
    print("\n[ERROR] Build failed. Check errors above.")
    return False

if __name__ == "__main__":
    build()
