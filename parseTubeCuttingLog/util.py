import config
import console

import datetime
from pathlib import Path


print = console.print


def getTimeStamp() -> str:
    now = datetime.datetime.now()
    return str(now.strftime(f"%Y/{now.month}/%d %H:%M:%S"))


def saveWorkbook(dstPath, wb): # {{{
    try:
        wb.save(str(dstPath))
        print(f"\n[{getTimeStamp()}]:[bold green]Saving Excel file at: [/bold green][bright_black]{dstPath}")
    except Exception as e:
        print(e)
        fallbackExcelPath  = Path(
            config.LOCALEXPORTDIR,
            str(
                datetime.datetime.now().strftime("%Y-%m-%d %H%M%S%f")
                ) + ".xlsx"
        )
        wb.save(str(fallbackExcelPath))
        print(f"\n[{getTimeStamp()}]:[bold green]Saving Excel file at: [/bold green][bright_black]{fallbackExcelPath}")

    print(f"[{getTimeStamp()}]:[bold white]Done[/bold white]") # }}}
