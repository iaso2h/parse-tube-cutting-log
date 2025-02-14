import config
import console
import util

import re
import os
import datetime
import win32api, win32con
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.comments import Comment
from pathlib import Path
from openpyxl.styles.numbers import BUILTIN_FORMATS
# https://openpyxl.readthedocs.io/en/3.1.3/_modules/openpyxl/styles/numbers.html


print = console.print


def removeDummyLaserFile(p: Path) -> None:
    if p.suffix == "" and p.stat().st_size == 0:
        try:
            os.remove(p)
        except:
            pass


def workpieceNamingVerification() -> None:
    laserFilePaths = util.getAllLaserFiles()
    if not laserFilePaths:
        print("All files match the naming convention!")
    invalidFilePathFoundChk = False
    for _, p in enumerate(laserFilePaths):
        if p.suffix == ".zx" or p.suffix == ".zzx":
            fileNameMatch = re.match(
                    config.RE_LASER_FILES_MATCH,
                    str(p.stem)
                    )
            if not fileNameMatch:
                invalidFilePathFoundChk = True
                print(f'------------------------\n({_}): "{p.stem}"')
    if not invalidFilePathFoundChk:
        print("没有不规范的工件名称")

def removeRedundantLaserFile() -> None:
    rawLaserFile = []

    if not config.LASER_FILE_DIR_PATH.exists():
        return

    for p in config.LASER_FILE_DIR_PATH.iterdir():
        p = util.strStandarize(p)
        if p.is_file() and "demo" not in p.stem.lower():
            rawLaserFile.append(p)

    pDeletedStr = []
    for p in rawLaserFile:
        laserFile = Path(p.parent, p.stem + ".zzx")
        if laserFile.exists() and laserFile.stat().st_mtime > p.stat().st_mtime:
            try:
                os.remove(p)
                pDeletedStr.append(str(p))
            except:
                pass

    if len(pDeletedStr) > 0:
        print(f"{len(pDeletedStr)} redundant .zx files has been deleted:")
        for pStr in pDeletedStr:
            print(pStr)
        win32api.MessageBox(
                    None,
                    f"{len(pDeletedStr)}个冗余文件已经被删除",
                    "Info",
                    4096 + 64 + 0
                )
                #   MB_SYSTEMMODAL==4096
                ##  Button Styles:
                ### 0:OK  --  1:OK|Cancel -- 2:Abort|Retry|Ignore -- 3:Yes|No|Cancel -- 4:Yes|No -- 5:Retry|No -- 6:Cancel|Try Again|Continue
                ##  To also change icon, add these values to previous number
                ### 16 Stop-sign  ### 32 Question-mark  ### 48 Exclamation-point  ### 64 Information-sign ('i' in a circle)
    else:
        print("No redundant .zx files")




def exportDimensions():
    laserFilePaths = util.getAllLaserFiles()
    wb = Workbook()
    ws = wb.create_sheet("Sheet1", 0)
    ws["A1"] = "更新时间:" + str(datetime.datetime.now().strftime("%Y-%m-%d %H%M%S%f"))
    ws.merge_cells("B1:F1")
    ws["A2"].value = "零件名称"
    ws.column_dimensions["A"].width = 45
    ws["B2"].value = "外发别名"
    ws.column_dimensions["B"].width = 14
    ws["C2"].value = "规格"
    ws.column_dimensions["C"].width = 18
    ws["D2"].value = "材料"
    ws.column_dimensions["D"].width = 8
    ws["E2"].value = "第二规格"
    ws.column_dimensions["D"].width = 12
    ws["F2"].value = "第二规格数值"
    ws.column_dimensions["D"].width = 15.5
    ws["G2"].value = "长度"
    ws.column_dimensions["G"].width = 7.5
    ws["H2"].value = "方数m²"
    ws.column_dimensions["H"].width = 9.5
    workpieceFullNames = []
    workpieceNickNames = {
            # <fullPartName>: ["<nickName>", "<comment>"]
            "513L 主体管": ["515L 主体管", "1)因为开厂以来第一款移动式是515L而不是513L，因此用在委外时用515L来泛指代表移动式助行器\n2)移动式助行器的插销孔套在固定式助行器的H架上面会被H架遮挡住，所以固定式助行器的主体管统一用移动式助行器的主体管(带有插销孔)是无所谓的"],
            "513L(移动式) 大开关管": ["515L 大开关管", "因为开厂以来第一款移动式是515L而不是513L，因此用在委外时用515L来泛指代表移动式助行器"],
            "513L(移动式) 小开关管": ["515L 小开关管", "因为开厂以来第一款移动式是515L而不是513L，因此用在委外时用515L来泛指代表移动式助行器"],
            "734L 底座 焊接组合": ["734L 四脚架", ""],
            }
    fileNamePat      = re.compile(config.RE_LASER_FILES_MATCH)
    tubeDimensionPat = re.compile(r"(∅.*?)\*(T.*?)\*(L.*)")
    for _, p in enumerate(laserFilePaths):
        if p.suffix == ".zx" or p.suffix == ".zzx":
            fileNameMatch = fileNamePat.match(str(p.stem))
        else:
            fileNameMatch = fileNamePat.match(str(p.name))

        rowMax = ws.max_row + 1

        if not fileNameMatch:
            if p.suffix == ".zx" or p.suffix == ".zzx":
                workpieceFullName = p.stem
            else:
                workpieceFullName = p.name
            ws[f"A{rowMax}"].value = workpieceFullName
            ws[f"A{rowMax}"].number_format = "@"
            # namingly ws[f"A{rowMax}"].number_format = BUILTIN_FORMATS[49]
            if workpieceFullName in workpieceNickNames:
                ws[f"B{rowMax}"].value = workpieceNickNames[workpieceFullName][0]
                ws[f"B{rowMax}"].number_format = "@"
                if workpieceNickNames[workpieceFullName][1]:
                    comment = Comment(workpieceNickNames[workpieceFullName][1], "阮焕")
                    comment.width = 300
                    comment.height = 150
                    ws[f"B{rowMax}"].comment = comment

            if workpieceFullName in workpieceFullNames:
                removeDummyLaserFile(p)
                continue
            else:
                workpieceFullNames.append(workpieceFullName)
        else:
            productId          = fileNameMatch.group(1)
            productIdNote      = fileNameMatch.group(2) # name
            workpieceName           = fileNameMatch.group(3)
            if workpieceName and "飞切" in workpieceName:
                workpieceName = re.sub(r"[有无]飞切", "", workpieceName)
                workpieceName = workpieceName.replace("()", "")
            workpieceComponentName  = fileNameMatch.group(4) # Optional
            workpieceMaterial       = fileNameMatch.group(5)
            workpieceDimension               = fileNameMatch.group(6)
            workpiece2ndDimensionInccator    = fileNameMatch.group(7) # Optional
            workpiece2ndDimensionInccatorNum = fileNameMatch.group(9) # Optional
            workpieceLength    = fileNameMatch.group(10)
            workpieceDimension = workpieceDimension.replace("_", "*")
            workpieceDimension = workpieceDimension.replace("x", "*")
            # workpieceDimension = workpieceDimension.replace("∅", "∅")
            workpieceDimension = workpieceDimension.replace("Ø", "∅")
            workpieceDimension = workpieceDimension.replace("Φ", "∅")
            workpieceDimension = workpieceDimension.replace("φ", "∅")
            workpieceDimension = workpieceDimension.strip()
            workpieceFullName = "{} {}".format(productId, workpieceName)
            if workpieceFullName in workpieceFullNames:
                removeDummyLaserFile(p)
                continue
            else:
                workpieceFullNames.append(workpieceFullName)

            tailingWorkpiece = fileNameMatch.group(12)          # Optional
            workpieceLongTubeLength = fileNameMatch.group(14) # Optional

            ws[f"A{rowMax}"].value = workpieceFullName
            ws[f"A{rowMax}"].number_format = "@"
            # namingly ws[f"A{rowMax}"].number_format = BUILTIN_FORMATS[49]
            if workpieceFullName in workpieceNickNames:
                ws[f"B{rowMax}"].value = workpieceNickNames[workpieceFullName][0]
                ws[f"B{rowMax}"].number_format = "@"
                if workpieceFullName[1]:
                    comment = Comment(workpieceNickNames[workpieceFullName][1], "阮焕")
                    comment.width = 300
                    comment.height = 150
                    ws[f"B{rowMax}"].comment = comment

            ws[f"C{rowMax}"].value = workpieceDimension
            ws[f"C{rowMax}"].number_format = "@"
            ws[f"D{rowMax}"].value = workpieceMaterial
            ws[f"D{rowMax}"].number_format = "@"
            ws[f"E{rowMax}"].value = workpiece2ndDimensionInccator
            ws[f"E{rowMax}"].number_format = "@"
            ws[f"F{rowMax}"].value = workpiece2ndDimensionInccatorNum
            ws[f"F{rowMax}"].number_format = BUILTIN_FORMATS[2]
            ws[f"G{rowMax}"].value = workpieceLength
            ws[f"G{rowMax}"].number_format = BUILTIN_FORMATS[2]
            # Calculate the surface area
            if "∅" in workpieceDimension and "T" in workpieceDimension and "L" in workpieceDimension:
                m = tubeDimensionPat.match(workpieceDimension)
                if m:
                    dia       = float(m.group(1)[1:])
                    length    = float(m.group(3)[1:])
                    surfaceArea = 3.14 * dia * length / 1000 / 1000
                    ws[f"H{rowMax}"].value = surfaceArea
                    ws[f"H{rowMax}"].number_format = BUILTIN_FORMATS[2]







    # Add table
    tab = Table(displayName="Table1", ref=f"A2:H{ws.max_row}")

    # Add a default style with striped rows and banded columns
    style = TableStyleInfo(
            name="TableStyleMedium16",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
            )
    tab.tableStyleInfo = style

    ws.add_table(tab)

    util.saveWorkbook(wb, Path(config.PARENT_DIR_PATH, r"存档\零件规格总览.xlsx"), True)


