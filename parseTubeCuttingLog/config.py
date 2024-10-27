# File: parseTubeProLog
# Author: iaso2h
# Description: Parsing Log files(.rtf) from TubePro and split them into separated files
# Version: 0.0.42
# Last Modified: 2024-10-27

import os
from pathlib import Path

VERSION     = "0.0.42"
LASTUPDATED = "2024-10-27"
AUTHOR      = "阮焕"
SILENT_MODE = False
PROGRAM_DIR = Path(os.getcwd())
LOCAL_EXPORT_DIR = Path(PROGRAM_DIR, "export")
RE_LASER_FILES_MATCH = r"^(\d{3}[-a-zA-Z]{,3}(\(.+?\))?) (([^_]+? )?[^_]+?) ([^_]{,8})[_*x]([^\u4e00-\u9fff]+?(([RT])?([0-9.]+?))[_*x]L([.0-9]{2,5}))( \d+?支 \+ (.*?) \d+?支)?( L(\d{4}))?_X1$"

LASER_FILE_DIR_PATH  = Path(r"D:\欧拓图纸\切割文件")
DISPATCH_FILE_PATH   = Path(r"D:\欧拓图纸\派工单（模板+空表）.xlsx")
SCREENSHOT_DIR_PATH  = Path(r"D:\欧拓图纸\存档\截图")
CUT_RECORD_PATH      = Path(r"D:\欧拓图纸\存档\开料记录.xlsx")
LASER_PORFILING_PATH = Path(r"D:\欧拓图纸\存档\开料耗时.xlsx")
LASER_LOG_PATH       = Path(r"D:\欧拓图纸\存档\耗时计算")
LASER_OCR_FIX_PATH                   = Path(r"D:\欧拓图纸\存档\辅助程序\激光名称OCR修复规则.json")
PRODUCT_ID_CATERGORY_CONVENTION_PATH = Path(r"D:\欧拓图纸\存档\辅助程序\型号类别对照规则.json")
