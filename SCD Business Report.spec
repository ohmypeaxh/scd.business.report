# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
import sys

python_root = Path(sys.base_prefix)
tk_binaries = [
    (str(python_root / 'DLLs' / '_tkinter.pyd'), '.'),
    (str(python_root / 'DLLs' / 'tcl86t.dll'), '.'),
    (str(python_root / 'DLLs' / 'tk86t.dll'), '.'),
]
tk_datas = [
    (str(python_root / 'Lib' / 'tkinter'), 'tkinter'),
    (str(python_root / 'tcl' / 'tcl8.6'), 'tcl/tcl8.6'),
    (str(python_root / 'tcl' / 'tk8.6'), 'tcl/tk8.6'),
    ('assets/scd.ico', 'assets'),
]

a = Analysis(['main.py'], pathex=['src'], binaries=tk_binaries, datas=tk_datas, hiddenimports=['tkinter'],
             hookspath=[], hooksconfig={}, runtime_hooks=[], excludes=[], noarchive=False)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, a.binaries, a.datas, [], name='SCD Business Report',
          debug=False, bootloader_ignore_signals=False, strip=False, upx=True,
          console=False, disable_windowed_traceback=False, icon='assets/scd.ico')
