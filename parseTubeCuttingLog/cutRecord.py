import util
import console
import config

import datetime
import os
from openpyxl import Workbook, load_workbook
from pathlib import Path


screenshotParentPath = Path(r"D:\欧拓图纸\存档\截图")
cutRecordPath = Path(r"D:\欧拓图纸\存档\开料记录.xlsx")
if cutRecordPath.exists():
    wb = load_workbook(str(cutRecordPath))
else:
    wb = Workbook()

print = console.print

screenshotPaths = []
sheetNames = []
def initFiles(wb):
    for p in screenshotParentPath.iterdir():
        if p.suffix == ".png":
            screenshotPaths.append(p)
            dateStamp = p.stem[5:12]
            if dateStamp not in sheetNames:
                sheetNames.append(dateStamp)

    for n in sheetNames:
        try:
            ws = wb[n]
        except Exception:
            ws = wb.create_sheet(n, 0)
            ws["A1"].value = "排样文件"
            ws["B1"].value = "长料长度"
            ws["C1"].value = "完成时间"
            ws["D1"].value = "单号"
            ws["E1"].value = "型号(数量)"
            ws["F1"].value = "已切量/需求量"
            ws["G1"].value = "截图文件"



def addScreendshots():

    initFiles(wb)
    # now = datetime.datetime.now()
    # sheetNameRightnow = now.strftime(f"%Y-{now.month}")

    for p in screenshotPaths:
        sheetName = p.stem[5:12]
        ws = wb[sheetName]
        if ws.max_row != 1:
            lastPath = Path(ws[f"G{ws.max_row}"].value)
            lastDatetime = datetime.datetime.strptime(str(lastPath.stem)[5:], "%Y-%m-%d %H%M%S")
            currentDatetime = datetime.datetime.strptime(str(p.stem)[5:], "%Y-%m-%d %H%M%S")
            if lastDatetime < currentDatetime:
                ws[f"G{ws.max_row + 1}"].hyperlink = str(p)
        else:
            ws[f"G{ws.max_row + 1}"].hyperlink = str(p)

    saveWorkbook()


def saveWorkbook(): # {{{
    try:
        wb.save(str(cutRecordPath))
        print(f"\n[{util.getTimeStamp()}]:[bold green]Saving Excel file at: [/bold green][bright_black]{cutRecordPath}")
    except Exception as e:
        print(e)
        excelFilePath  = Path(
            config.PROGRAMDIR,
            str(
                datetime.datetime.now().strftime("%Y-%m-%d %H%M%S%f")
                ) + ".xlsx"
        )
        wb.save(str(excelFilePath))
        print(f"\n[{util.getTimeStamp()}]:[bold green]Saving Excel file at: [/bold green][bright_black]{excelFilePath}")

    print(f"[{util.getTimeStamp()}]:[bold white]Done[/bold white]") # }}}
