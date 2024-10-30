import config

import PyInstaller.__main__
import os
from pathlib import Path
config.PROGRAM_DIR = Path(os.getcwd())

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
