# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [("static", "static"), ("config.example.json", ".")]
hiddenimports = []
for package in ("uvicorn", "pymodbus", "serial", "sqlalchemy", "pandas"):
    d, b, h = collect_all(package)
    datas += d
    hiddenimports += h

a = Analysis(["launcher.py"], pathex=[], binaries=[], datas=datas,
             hiddenimports=hiddenimports, hookspath=[], hooksconfig={},
             runtime_hooks=[], excludes=["pytest"], noarchive=False)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, a.binaries, a.datas, [], name="OZDEKAN",
          debug=False, bootloader_ignore_signals=False, strip=False,
          upx=True, console=True, disable_windowed_traceback=False)
