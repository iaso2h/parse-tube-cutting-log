from os.path import isfile
import config
import console
import os

import shutil
import datetime
import win32api, win32con
import re
from pathlib import Path


print = console.print


def getTimeStamp() -> str:
    now = datetime.datetime.now()
    return str(now.strftime("%H:%M:%S"))
    # return str(now.strftime(f"%Y/{now.month}/%d %H:%M:%S"))


def saveWorkbook(wb, dstPath=None, openAfterSaveChk=False): # {{{
    os.makedirs(config.LOCAL_EXPORT_DIR, exist_ok=True)
    timeStr = str(datetime.datetime.now().strftime("%Y-%m-%d %H%M%S%f"))

    if dstPath and (os.getlogin() == "OT03" or config.DEV_MODE):
        if dstPath.exists():
            backupPath = Path(
                config.LOCAL_EXPORT_DIR,
                dstPath.stem + "_backup_" + timeStr + ".xlsx"
            )
            shutil.copy2(dstPath, backupPath)

        try:
            wb.save(str(dstPath))
            print(f"\n[{getTimeStamp()}]:[bold green]Saving Excel file at: [/bold green][bright_black]{dstPath}")
            if openAfterSaveChk:
                os.startfile(dstPath)
        except PermissionError:
            if win32con.IDYES == win32api.MessageBox(
                None,
                f"是否要重新写入该路径？\n\"{str(dstPath)}\"",
                "写入权限不足",
                win32con.MB_YESNO | win32con.MB_ICONQUESTION
                ):
                saveWorkbook(wb, dstPath, openAfterSaveChk)
            else:
                fallbackExcelPath = Path(
                    config.LOCAL_EXPORT_DIR,
                    dstPath.stem + "_fallback_" + timeStr + ".xlsx")
                wb.save(str(fallbackExcelPath))
                print(f"\n[{getTimeStamp()}]:[bold green]Saving fallback Excel file at: [/bold green][bright_black]{fallbackExcelPath}")

    else:
        newExcelPath = Path(
            config.LOCAL_EXPORT_DIR,
            timeStr + ".xlsx")
        wb.save(str(newExcelPath))
        print(f"\n[{getTimeStamp()}]:[bold green]Saving new Excel file at: [/bold green][bright_black]{newExcelPath}")
        if openAfterSaveChk:
            os.startfile(newExcelPath)

    print(f"[{getTimeStamp()}]:[bold white]Done[/bold white]") # }}}



def strStandarize(old: Path):
    if old.is_file():
        new = str(old)
        # old = old.replace("∅", "∅")
        new = new.replace("Ø", "∅")
        new = new.replace("Φ", "∅")
        new = new.replace("φ", "∅")
        new = re.sub(r"\s{2,}", " ", new)
        try:
            os.rename(old, new)
            return Path(new)
        except PermissionError as e:
            print(str(e))
            return old
    else:
        return old


def getAllLaserFiles(): # {{{
    laserFilePaths = []

    if not config.LASER_FILE_DIR_PATH.exists():
        return

    for p in config.LASER_FILE_DIR_PATH.iterdir():
        p = strStandarize(p)
        if p.is_file() and "demo" not in p.stem.lower() and (p.suffix == ".zx" or p.suffix == ".zzx" or p.suffix == ""):
            laserFilePaths.append(p)

    return laserFilePaths # }}}
