import PyInstaller.__main__
import rtfParse
import os
from pathlib import Path
rtfParse.programDir = Path(os.getcwd())

PyInstaller.__main__.run([
    "__main__.py",
    "--onefile",
    "--noconfirm",
    "--console",
    "--clean",
    "--name=ParseTubeProLog",
    "--hidden-import=openpyxl.cell._writer",
    "--icon=./src/sticky-note.ico",
])
