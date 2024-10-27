import config
import console
import util

import re
import json
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo


with open(config.PRODUCT_ID_CATERGORY_CONVENTION_PATH, "r", encoding="utf-8") as pat:
    productIdCatergoryConvention = json.load(pat)

wb = Workbook()
print = console.print


def invalidNamingParts():
    laserFilePaths = util.getAllLaserFiles()
    if not laserFilePaths:
        print("All files match the naming convention!")
    for _, p in enumerate(laserFilePaths):
        fileNameMatch = re.match(
                config.RE_LASER_FILES_MATCH,
                str(p.stem)
                )
        if not fileNameMatch:
            print(f'--------\n{_}: "{p.stem}"')


def exportDimensions():
    laserFilePaths = util.getAllLaserFiles()
    ws = wb.create_sheet("Sheet1", 0)
    ws["A1"].value = "零件名称"
    ws["B1"].value = "规格"
    ws["C1"].value = "材料"
    ws["D1"].value = "第二规格指标"
    ws["E1"].value = "第二规格指标数值"
    ws["F1"].value = "长度"
    partFullNames = []
    for _, p in enumerate(laserFilePaths):
        fileNameMatch = re.match(
                config.RE_LASER_FILES_MATCH,
                str(p.stem)
                )
        if not fileNameMatch:
            continue

        productId          = fileNameMatch.group(1)
        productIdNote      = fileNameMatch.group(2) # name
        partName           = fileNameMatch.group(3)
        partComponentName  = fileNameMatch.group(4)  # Optional
        partMaterial       = fileNameMatch.group(5)
        partDimension                  = fileNameMatch.group(6)
        part2ndDimensionInccator    = fileNameMatch.group(7) # Optional
        part2ndDimensionInccatorNum = fileNameMatch.group(9) # Optional
        partLength    = fileNameMatch.group(10)
        partDimension = partDimension.replace("_", "*")
        partDimension = partDimension.replace("x", "*")
        partDimension = partDimension.replace("∅", "D")
        partDimension = partDimension.replace("Ø", "D")
        partDimension = partDimension.replace("Φ", "D")
        partDimension = partDimension.strip()
        partFullName = "{} {}\n({}/{})".format(productId, partName, partMaterial, partDimension)
        if partFullName in partFullNames:
            continue
        else:
            partFullNames.append(partFullName)

        otherPart = fileNameMatch.group(12)          # Optional
        partLongTubeLength = fileNameMatch.group(14) # Optional

        rowMax = ws.max_row + 1
        ws[f"A{rowMax}"].value = partFullName
        ws[f"A{rowMax}"].number_format = "@"
        ws[f"B{rowMax}"].value = partDimension
        ws[f"B{rowMax}"].number_format = "@"
        ws[f"C{rowMax}"].value = partMaterial
        ws[f"C{rowMax}"].number_format = "@"
        ws[f"D{rowMax}"].value = part2ndDimensionInccator
        ws[f"D{rowMax}"].number_format = "@"
        ws[f"E{rowMax}"].value = part2ndDimensionInccatorNum
        ws[f"E{rowMax}"].number_format = "0"
        ws[f"F{rowMax}"].value = partLength
        ws[f"F{rowMax}"].number_format = "0"


    # Add table
    tab = Table(displayName="Table1", ref=f"A1:F{ws.max_row}")

    # Add a default style with striped rows and banded columns
    style = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
            )
    tab.tableStyleInfo = style

    ws.add_table(tab)

    util.saveWorkbook(wb)


