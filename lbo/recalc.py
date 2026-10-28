"""Recalculate an .xlsx with headless LibreOffice so cached values exist (Excel itself recalculates on open)."""
from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path


def soffice_path() -> str | None:
    return shutil.which("soffice") or shutil.which("libreoffice")


def recalc(path: str, timeout: int = 120) -> bool:
    exe = soffice_path()
    if not exe:
        return False
    path = str(Path(path).resolve())
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([exe, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", tmp, path],
                       check=True, timeout=timeout, capture_output=True)
        shutil.move(str(Path(tmp) / Path(path).name), path)
    return True
