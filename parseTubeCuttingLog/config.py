import os
from pathlib import Path

SILENTMODE = False
PROGRAMDIR = Path(os.getcwd())
LOCALEXPORTDIR = Path(PROGRAMDIR, "export")
LASERFILESTEMMATCH = r"^(\d{3}[-a-zA-Z]{,3}(\(.+?\))?) (([^_]+? )?[^_]+?) (.*?)_(.+?)(\d+?支 \+ (.*?)\s\d+?支)?( L(\d{4}))?_X1"
