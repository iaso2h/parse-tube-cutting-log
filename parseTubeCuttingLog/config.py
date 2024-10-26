# File: parseTubeProLog
# Author: iaso2h
# Description: Parsing Log files(.rtf) from TubePro and split them into separated files
# Version: 0.0.38
# Last Modified: 2024-10-24

import os
from pathlib import Path

VERSION     = "0.0.39"
LASTUPDATED = "2024-10-25"
AUTHOR      = "阮焕"
SILENTMODE = False
PROGRAMDIR = Path(os.getcwd())
LOCALEXPORTDIR = Path(PROGRAMDIR, "export")
LASERFILESTEMMATCH = r"^(\d{3}[-a-zA-Z]{,3}(\(.+?\))?) (([^_]+? )?[^_]+?) (.*?)_(.+?)(\d+?支 \+ (.*?)\s\d+?支)?( L(\d{4}))?_X1"
