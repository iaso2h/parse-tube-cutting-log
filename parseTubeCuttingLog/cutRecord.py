import util
import console
import config

import datetime
import os
import re
import easyocr
import numpy
from PIL import Image, ImageFilter
from openpyxl import Workbook, load_workbook
from pathlib import Path


screenshotParentPath = Path(r"D:\欧拓图纸\存档\截图")
cutRecordPath = Path(r"D:\欧拓图纸\存档\开料记录.xlsx")
if cutRecordPath.exists():
    wb = load_workbook(str(cutRecordPath))
else:
    wb = Workbook()

print = console.print
reader = easyocr.Reader(["ch_sim", "en"])

screenshotPaths = []
sheetNames = []
def initSheetImg(wb):
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


def takeScreenshot():
    pass


def getImgInfo(p:Path):
    with Image.open(p) as img:
        imgTitle        = img.crop((91, 0, 900, 25))
        imgProcessCount = img.crop((550, 1665, 765, 1685))
        cvTitle = numpy.array(imgTitle)[:, :, ::-1].copy()
        cvProcessCount = numpy.array(imgProcessCount)[:, :, ::-1].copy()

        imgRGB = img.convert("RGB")
        targetCompletedPixel = imgRGB.getpixel((15, 1810))
        if targetCompletedPixel == (170, 170, 0):
            targetCompletedChk = True
        else:
            targetCompletedChk = False

        if targetCompletedChk:
            imgTimeStamp = img.crop((104, 1777, 240, 1792))
            cvTimeStamp  = numpy.array(imgTimeStamp)[:, :, ::-1].copy()
        else:
            imgTimeStamp = img.crop((91, 1755, 185, 1864)).filter(ImageFilter.EDGE_ENHANCE)
            cvTimeStamp  = numpy.array(imgTimeStamp)[:, :, ::-1].copy()

    titleRead = reader.readtext(cvTitle)
    processCountRead = reader.readtext(cvProcessCount)
    timeStampRead = reader.readtext(cvTimeStamp)
    partFileName = ""
    partProcessCount = ""
    timeStamp = p.stem[5:] # Default time stamp
    if titleRead:
        for text in titleRead:
            partFileName = partFileName + " " + text[1]
            suffixMatch = re.search(r"\.zzx", partFileName, flags=re.IGNORECASE)
            if suffixMatch:
                partFileName = partFileName[:suffixMatch.span()[1]]
            partFileName = partFileName.strip()


    if processCountRead:
        partProcessCount = processCountRead[1][1]

    if timeStampRead:
        timeStamp = timeStampRead[len(timeStampRead) - 1][1]
        if not targetCompletedChk:
            timeStamp = p.stem[5:9] + "/" + timeStamp # Add year prefix

    return partFileName, partProcessCount, timeStamp


def saveWorkbook(): # {{{
    try:
        wb.save(str(cutRecordPath))
        print(f"\n[{util.getTimeStamp()}]:[bold green]Saving Excel file at: [/bold green][bright_black]{cutRecordPath}")
    except Exception as e:
        print(e)
        excelFilePath  = Path(
            config.LOCALEXPORTDIR,
            str(
                datetime.datetime.now().strftime("%Y-%m-%d %H%M%S%f")
                ) + ".xlsx"
        )
        wb.save(str(excelFilePath))
        print(f"\n[{util.getTimeStamp()}]:[bold green]Saving Excel file at: [/bold green][bright_black]{excelFilePath}")

    print(f"[{util.getTimeStamp()}]:[bold white]Done[/bold white]") # }}}



def writeNewRecord():

    initSheetImg(wb)
    # now = datetime.datetime.now()
    # sheetNameRightnow = now.strftime(f"%Y-{now.month}")
    def writeColumn(p):
        partFileName, partProcessCount, timeStamp = getImgInfo(p)
        longTubeLengthMatch = re.search(r"(?<=L)\d{4}(?=.{0,3}\.zz?x$)", partFileName, flags=re.IGNORECASE)
        newRow = ws.max_row + 1
        ws[f"A{newRow}"].value = partFileName
        if longTubeLengthMatch:
            ws[f"B{newRow}"].value = int(longTubeLengthMatch.group())
        ws[f"C{newRow}"].value = timeStamp
        ws[f"C{newRow}"].number_format = "yyyy/m/d h:mm:ss"
        ws[f"F{newRow}"].value = str(partProcessCount)
        ws[f"G{newRow}"].hyperlink = str(p)

    for p in screenshotPaths:
        sheetName = p.stem[5:12]
        ws = wb[sheetName]
        if ws.max_row != 1:
            lastPath = Path(ws[f"G{ws.max_row}"].value)
            lastDatetime = datetime.datetime.strptime(str(lastPath.stem)[5:], "%Y-%m-%d %H%M%S")
            currentDatetime = datetime.datetime.strptime(str(p.stem)[5:], "%Y-%m-%d %H%M%S")
            # Only save screenshots that are newer than the last one
            if lastDatetime < currentDatetime:
                writeColumn(p)
        else:
            # Start in a new worksheet
            writeColumn(p)

    saveWorkbook()
