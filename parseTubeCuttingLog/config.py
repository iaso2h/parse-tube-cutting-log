# File: parseTubeProLog
# Author: iaso2h
# Description: Parsing Log files(.rtf) from TubePro and split them into separated files
VERSION     = "0.0.60"
LASTUPDATED = "2024-12-04"

import os
from pathlib import Path

GUI_MODE    = False
SILENT_MODE = False
DEV_MODE    = False
PROGRAM_DIR = Path(os.getcwd())
LOCAL_EXPORT_DIR = Path(PROGRAM_DIR, "export")
RE_LASER_FILES_MATCH = r"^(\d{3}[-a-zA-Z]{,3}(\(.+?\))?) (([^_]+? )?[^_]+?) ([^_]{,8})[_*x]([^\u4e00-\u9fff]+?(([RT])?([0-9.]+?))[_*x]L([.0-9]{2,5}))( \d+?支 \+ (.*?) \d+?支)?( L(\d{4}))?_X1$"

PARENT_DIR_PATH = Path(r"D:\欧拓图纸")

LASER_FILE_DIR_PATH  = None
DISPATCH_FILE_PATH   = None
SCREENSHOT_DIR_PATH  = None
CUT_RECORD_PATH      = None
LASER_PORFILING_PATH = None
LASER_LOG_PATH       = None
LASER_OCR_FIX_PATH                   = None
PRODUCT_ID_CATERGORY_CONVENTION_PATH = None
GUI_GEOMETRY_PATH = None

def updaPath():
    global LASER_FILE_DIR_PATH
    global DISPATCH_FILE_PATH
    global SCREENSHOT_DIR_PATH
    global CUT_RECORD_PATH
    global LASER_PORFILING_PATH
    global LASER_LOG_PATH
    global LASER_OCR_FIX_PATH
    global PRODUCT_ID_CATERGORY_CONVENTION_PATH
    global GUI_GEOMETRY_PATH

    LASER_FILE_DIR_PATH  = Path(PARENT_DIR_PATH, r"切割文件")
    DISPATCH_FILE_PATH   = Path(PARENT_DIR_PATH, r"派工单（模板+空表）.xlsx")
    SCREENSHOT_DIR_PATH  = Path(PARENT_DIR_PATH, r"存档\截图")
    CUT_RECORD_PATH      = Path(PARENT_DIR_PATH, r"存档\开料记录.xlsx")
    LASER_PORFILING_PATH = Path(PARENT_DIR_PATH, r"存档\开料耗时.xlsx")
    LASER_LOG_PATH       = Path(PARENT_DIR_PATH, r"存档\耗时计算")
    LASER_OCR_FIX_PATH                   = Path(PARENT_DIR_PATH, r"辅助程序\激光名称OCR修复规则.json")
    PRODUCT_ID_CATERGORY_CONVENTION_PATH = Path(PARENT_DIR_PATH, r"辅助程序\型号类别对照规则.json")
    GUI_GEOMETRY_PATH = Path(PARENT_DIR_PATH, r"辅助程序\程序几何.json")
