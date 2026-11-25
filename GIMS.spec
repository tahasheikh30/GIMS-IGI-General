# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller build spec for GIMS (onedir build).
#
#   python -m PyInstaller --noconfirm --clean GIMS.spec
#
# Output: dist\GIMS\GIMS.exe  (copy the WHOLE dist\GIMS folder, not just the
# .exe). A folder build is used on purpose: a onefile exe re-unpacks itself
# on every launch and antivirus / OneDrive scanning flags that far more
# often than a plain folder of files. UPX is off for the same reason.
#
# Icon: optional. Put a real .ico file at assets\GIMS.ico and it is used
# automatically; without it the default icon is used.

import os

# SPECPATH is provided by PyInstaller: the folder this spec file lives in.
ICON_PATH = os.path.join(SPECPATH, "assets", "GIMS.ico")

a = Analysis(
    ["app.py"],
    pathex=[SPECPATH],
    binaries=[],
    # app.py loads its logos from <app folder>\assets at startup.
    datas=[("assets", "assets")],
    hiddenimports=[
        # Outlook email sending / native pivot (loaded lazily inside functions).
        "win32com.client",
        "pythoncom",
        "pywintypes",
        # Excel readers chosen at runtime by name.
        "xlrd",
        "python_calamine",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="GIMS",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,  # windowed app: no console window
    icon=ICON_PATH if os.path.exists(ICON_PATH) else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="GIMS",
)
