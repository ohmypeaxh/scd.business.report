from __future__ import annotations

import ctypes
from pathlib import Path
import subprocess
import sys


APP_NAME = "SCD Business Report"


def resource_path(name: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[2]))
    return base / "assets" / name


def desktop_path() -> Path:
    buffer = ctypes.create_unicode_buffer(260)
    result = ctypes.windll.shell32.SHGetFolderPathW(None, 0x10, None, 0, buffer)
    if result == 0 and buffer.value:
        return Path(buffer.value)
    return Path.home() / "Desktop"


def _ps_quote(value: Path | str) -> str:
    return "'" + str(value).replace("'", "''") + "'"


def create_desktop_shortcut() -> Path | None:
    """Create or refresh a desktop shortcut when running as a packaged EXE."""
    if not getattr(sys, "frozen", False):
        return None
    executable = Path(sys.executable).resolve()
    shortcut = desktop_path() / f"{APP_NAME}.lnk"
    icon = executable
    script = (
        "$w=New-Object -ComObject WScript.Shell;"
        f"$s=$w.CreateShortcut({_ps_quote(shortcut)});"
        f"$s.TargetPath={_ps_quote(executable)};"
        f"$s.WorkingDirectory={_ps_quote(executable.parent)};"
        f"$s.IconLocation={_ps_quote(str(icon) + ',0')};"
        "$s.Description='SCD Business Report';$s.Save()"
    )
    subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script],
        check=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        timeout=15,
    )
    return shortcut
