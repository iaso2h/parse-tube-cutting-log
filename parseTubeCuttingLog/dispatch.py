import config
import util
import console

import re
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Border, Side

from openpyxl.worksheet.cell_range import CellRange

print = console.print
laserCutFileParentPath = Path(r"D:\欧拓图纸\切割文件")
dispatchFilePath = Path(r"D:\欧拓图纸\派工单（模板+空表） - 副本.xlsx")
partColumnLetter = "E"
partColumnNum = 5
wb = load_workbook(str(dispatchFilePath))

laserFilePaths = []
def getAllLaserFiles():
    if not laserCutFileParentPath.exists():
        return

    for p in laserCutFileParentPath.iterdir():
        if p.is_file() and p.suffix == ".zx" or p.suffix == ".zzx":
            laserFilePaths.append(p)


def fillPartInfo(): # {{{
    getAllLaserFiles()
    ws = wb["OT计件表"]
    if not laserFilePaths:
        print(f"[red]No laser files found in: {str(laserCutFileParentPath)}[/red]")
        raise SystemExit(1)


    for _, p in enumerate(laserFilePaths):
        # https://regex101.com
        fileNameMatch = re.match(
                config.LASERFILESTEMMATCH,
                str(p.stem)
                )
        if not fileNameMatch:
            continue

        productId          = fileNameMatch.group(1)
        productIdNote      = fileNameMatch.group(2) # Optional
        partName           = fileNameMatch.group(3)
        partComponentName  = fileNameMatch.group(4)  # Optional
        partMaterial       = fileNameMatch.group(5)
        partDimension = fileNameMatch.group(6)
        partDimension = partDimension.replace("_", "*")
        partDimension = partDimension.replace("x", "*")
        partDimension = partDimension.replace("∅", "D")
        partDimension = partDimension.replace("Ø", "D")
        partDimension = partDimension.strip()
        partFullName = "{} {}\n({}/{})".format(productId, partName, partMaterial, partDimension)
        otherPart = fileNameMatch.group(8)           # Optional
        partLongTubeLength = fileNameMatch.group(10) # Optional

        # DEBUG: # {{{

        # print("---------------------------------------------")
        # print(_)
        # print("Laser File:", fileNameMatch.group(0))
        # print("productId 1:", productId)
        # if productIdNote:
        #     print("productIdNote 2:", productIdNote)
        # print("partName 3:", partName)
        # if partComponentName:
        #     print("partComponentName 4:", partComponentName)
        # print("partMaterial 5:", partMaterial)
        # print("partDimension 6:", partDimension)
        # if otherPart:
        #     print("otherPart 8:", otherPart)
        # if partLongTubeLength:
        #     print("partLongTubeLength 10", partLongTubeLength)
        # print("\n") # }}}

        def writePartInfo(mergedRng) -> int: # {{{
            newRow = None
            for rowPart in ws.iter_rows(min_col=partColumnNum, max_col=partColumnNum, min_row=int(mergedRng[0][1:]), max_row=int(mergedRng[1][1:])):
                for cellPart in rowPart:
                    if partFullName == cellPart.value:
                        return ""

                # If not exsting part info, then insert new row at the last row
                # NOTE: mergedRng is a static list whose range has been expaned over time
                lastRow = int(mergedRng[1][1:])
                if ws[f"{partColumnLetter}{lastRow}"].value == partFullName:
                    # Avoid overlapping part info
                    return ""
                newRow = lastRow + 1
                ws.insert_rows(newRow)
                ws[f"{partColumnLetter}{newRow}"].value = partFullName
                return newRow # }}}

        productExistsChk = False
        for rowProductId in ws.iter_rows(min_col=3, max_col=3, min_row=4, max_row=ws.max_row):

            rowMax = ws.max_row
            for cellProductId in rowProductId:
                # NOTE: merged cell doesn't have value
                if cellProductId.value and cellProductId.value[2:] == productId:
                    # Existing product ID
                    productExistsChk = True

                    # Find existing merged product ID
                    mergedProudctRng = []
                    for rng in ws.merged_cells.ranges:
                        if f"C{cellProductId.row}" in rng:
                            if ":" in rng.coord:
                                mergedProudctRng = rng.coord.split(":")
                            else:
                                # Make up range for merged cell consists of only one cell
                                mergedProudctRng = [rng, rng]

                            break


                    if mergedProudctRng:
                        rowNew = writePartInfo(mergedProudctRng)
                        if rowNew:
                            ws.unmerge_cells(mergedProudctRng[0] + ":" + mergedProudctRng[1])
                            ws.merge_cells(mergedProudctRng[0] + ":C" + str(rowNew))

                    else:
                        # Find part info range
                        unMergedProductRng = [cellProductId.coordinate]
                        for rowPart in ws.iter_rows(min_col=partColumnNum, max_col=partColumnNum, min_row=cellProductId.row, max_row=rowMax):
                            for cellPart in rowPart:
                                if cellPart.value[:2] != productId:
                                    unMergedProductRng.append(f"{cellProductId.column_letter}{cellPart.row-1}")
                                    rowNew = writePartInfo(unMergedProductRng)
                                    if rowNew:
                                        ws.merge_cells(unMergedProductRng[0] + ":C" + str(rowNew))
                                    break

                        # If no other product ID exists, append the max row
                        unMergedProductRng.append(f"{cellProductId.column_letter}{rowMax}")
                        rowNew = writePartInfo(unMergedProductRng)
                        # if rowNew:
                            # ws.merge_cells(unMergedProductRng[0] + ":C" + str(rowNew))
                        break

                    break

            if not productExistsChk:
                # If no product ID matches, write new product
                rowNew = rowMax + 1
                ws[f"C{rowNew}"] = "OT" + productId
                ws[f"{partColumnLetter}{rowNew}"] = partFullName
                ws.merge_cells(f"C{rowNew}:C{rowNew}")
                break



    util.saveWorkbook(dispatchFilePath, wb) # }}}


def getRowSections(ws, rowFirst, rowLast):
    sections = []

    # NOTE: ws["C"] won't yield row greater than the maximum row
    for i, cell in enumerate(ws["C"]):
        if sections:
            lastSectionPair = sections[len(sections) - 1]

        rowNum = i + 1

        if rowNum < rowFirst:
            continue
        if rowNum == rowLast:
            if len(lastSectionPair) % 2 == 1:
                lastSectionPair.append(rowNum)
            break
        if cell.value:
            if not sections:
                sections.append([rowNum])
            else:
                if len(lastSectionPair) % 2 == 1:
                    if cell.value != ws["C{}".format(lastSectionPair[0])].value:
                        lastSectionPair.append(rowNum - 1)
                        sections.append([rowNum])
                else:
                    if cell.value != ws["C{}".format(lastSectionPair[0])].value:
                        sections.append([rowNum])

    return sections

def beautifyCells():
    ws = wb["OT计件表"]
    rowMax = ws.max_row
    colMax = ws.max_column
    # Get product id row sections
    productIdRowSections = getRowSections(ws, 4, rowMax)


    # Merge product Id rows
    for rowPair in productIdRowSections:
        ws.merge_cells(f"C{rowPair[0]}:C{rowPair[1]}")


    # Fill sequence number
    for rowPair in productIdRowSections:
        for i, rowNum in enumerate(range(rowPair[0], rowPair[1] + 1)):
            sequenceNum = i + 1
            ws[f"A{rowNum}"].value = sequenceNum


    # Merge order number rows
    for rowPair in productIdRowSections:
        orderNum = ""
        for rowOrder in ws.iter_rows(min_col=2, max_col=2, min_row=rowPair[0], max_row=rowPair[1]):
            for cellOrder in rowOrder:
                if cellOrder.value:
                    orderNum = cellOrder.value
                    break
        if orderNum:
            ws.merge_cells(f"B{rowPair[0]}:B{rowPair[1]}")
            ws[f"B{rowPair[0]}"].value = orderNum


    # Merge part info rows
    # Get part info row sections
    for rowPair in productIdRowSections:
        partInfoRowSections = getRowSections(ws, rowPair[0], rowPair[1])
        for rowPartPair in partInfoRowSections:
            ws.merge_cells(f"D{rowPartPair[0]}:D{rowPartPair[1]}")

    # Add border to
    thin = Side(border_style="thin", color="FF000000")
    for row in ws[f"A3:O{rowMax}"]:
        for cell in row:
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
