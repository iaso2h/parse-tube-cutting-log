import config
import console
import os

import shutil
import datetime
from pathlib import Path


print = console.print


def getTimeStamp() -> str:
    now = datetime.datetime.now()
    return str(now.strftime(f"%Y/{now.month}/%d %H:%M:%S"))


def saveWorkbook(wb, dstPath=None, openAfterSaveChk=False): # {{{
    os.makedirs(config.LOCAL_EXPORT_DIR, exist_ok=True)

    if dstPath and (os.getlogin() == "OT03" or config.DEV_MODE):
        if dstPath.exists():
            backupPath = Path(
                config.LOCAL_EXPORT_DIR,
                dstPath.stem + str(
                    datetime.datetime.now().strftime("%Y-%m-%d %H%M%S%f")
                    ) + ".xlsx"
            )
            shutil.copy2(dstPath, backupPath)

        try:
            wb.save(str(dstPath))
            print(f"\n[{getTimeStamp()}]:[bold green]Saving Excel file at: [/bold green][bright_black]{dstPath}")
            if openAfterSaveChk:
                os.startfile(dstPath)
        except Exception as e:
            print(e)
            fallbackExcelPath = Path(
                config.LOCAL_EXPORT_DIR,
                str(
                    datetime.datetime.now().strftime("%Y-%m-%d %H%M%S%f")
                    ) + ".xlsx"
            )
            wb.save(str(fallbackExcelPath))
            print(f"\n[{getTimeStamp()}]:[bold green]Saving Excel file at: [/bold green][bright_black]{fallbackExcelPath}")
            if openAfterSaveChk:
                os.startfile(fallbackExcelPath)
    else:
        fallbackExcelPath = Path(
            config.LOCAL_EXPORT_DIR,
            str(
                datetime.datetime.now().strftime("%Y-%m-%d %H%M%S%f")
                ) + ".xlsx"
        )
        wb.save(str(fallbackExcelPath))
        print(f"\n[{getTimeStamp()}]:[bold green]Saving Excel file at: [/bold green][bright_black]{fallbackExcelPath}")
        if openAfterSaveChk:
            os.startfile(fallbackExcelPath)

    print(f"[{getTimeStamp()}]:[bold white]Done[/bold white]") # }}}


def getAllLaserFiles(): # {{{
    laserFilePaths = []

    if not config.LASER_FILE_DIR_PATH.exists():
        return

    for p in config.LASER_FILE_DIR_PATH.iterdir():
        if p.is_file() and p.suffix == ".zx" or p.suffix == ".zzx" and "demo" not in p.stem.lower():
            laserFilePaths.append(p)

    return laserFilePaths # }}}
