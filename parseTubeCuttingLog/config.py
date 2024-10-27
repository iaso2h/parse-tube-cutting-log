# File: parseTubeProLog
# Author: iaso2h
# Description: Parsing Log files(.rtf) from TubePro and split them into separated files
# Version: 0.0.41
# Last Modified: 2024-10-27

import os
from pathlib import Path

VERSION     = "0.0.41"
LASTUPDATED = "2024-10-27"
AUTHOR      = "阮焕"
SILENT_MODE = False
PROGRAM_DIR = Path(os.getcwd())
LOCAL_EXPORT_DIR = Path(PROGRAM_DIR, "export")
RE_LASER_FILES_MATCH = r"^(\d{3}[-a-zA-Z]{,3}(\(.+?\))?) (([^_]+? )?[^_]+?) (.*?)_(.+?)(\d+?支 \+ (.*?)\s\d+?支)?( L(\d{4}))?_X1"

LASER_FILE_DIR_PATH  = Path(r"D:\欧拓图纸\切割文件")
DISPATCH_FILE_PATH   = Path(r"D:\欧拓图纸\派工单（模板+空表）.xlsx")
SCREENSHOT_DIR_PATH  = Path(r"D:\欧拓图纸\存档\截图")
CUT_RECORD_PATH      = Path(r"D:\欧拓图纸\存档\开料记录.xlsx")
LASER_PORFILING_PATH = Path(r"D:\欧拓图纸\存档\开料耗时.xlsx")
LASER_LOG_PATH       = Path(r"D:\欧拓图纸\存档\耗时计算")
