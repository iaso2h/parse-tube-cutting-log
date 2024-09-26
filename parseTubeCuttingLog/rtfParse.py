import console
import chardet
import os
import re
import datetime
import time
from collections import Counter
from pathlib import Path
from striprtf.striprtf import rtf_to_text
from openpyxl import Workbook, load_workbook


speedTrackFilePath = Path("D:\欧拓图纸\存档\耗时计算\统计表格.xlsx")
if speedTrackFilePath.exists():
    wb = load_workbook(str(speedTrackFilePath))
else:
    wb = Workbook()
programPath = Path(__file__).resolve()
# programDir = programPath.parent
programDir = Path(os.getcwd())
exportDir  = Path(programDir, "export")
excelCreatedChk = False
print = console.print
rtfCandidates = []
rtfTarget = ""


def getTimeStamp():
    now = datetime.datetime.now()
    return str(now.strftime(f"%Y/{now.month}/%d %H:%M:%S"))


def getProperSheetName(name: str) -> str:
    conciseFileNameMatch = re.search(r"^.*?(?=A3)|^.*?(?=6063(-T\d)?)|^.*?(?=6061(-T\d)?)", name, re.I)
    if conciseFileNameMatch:
        conciseFileName = conciseFileNameMatch.group().strip()
    else:
        conciseFileName = name

    if len(conciseFileName) > 31:
        return conciseFileName[:31]
    else:
        return conciseFileName


laserCutKeywords = {}


def getEncoding(filePath):
    # Create a magic object
    with open(filePath, "rb") as f:
        # Detect the encoding
        rawData = f.read()
        result = chardet.detect(rawData)
        return result["encoding"]


def convertLogTimeStamp(timeStampStr):
    return datetime.datetime.strptime(timeStampStr, "%m/%d %H:%M:%S")


def parseStart():
    # https://rich.readthedocs.io/en/stable/appendix/colors.html
    print(f"[{getTimeStamp()}]:[yellow]Parsing rtf file:[/yellow] [bright_black]{rtfTarget}")
    openIndexes = []
    # loop through file convert rtf content to plain text and find keywords of opening files
    with open(rtfTarget, "r", encoding=getEncoding(rtfTarget)) as f:
        content = rtf_to_text(f.read())
        lines = content.split("\n")
        for openIndex, l in enumerate(lines):
            if re.match(r".*\)打开文件.*", l):
                openIndexes.append(openIndex)

    if not content or not openIndexes:
        print("[red]No available content to parse inside this .rtf file[/red]")
        return


    # Loop through each laser cut file section starting with keywords about opening a file
    for i, openIndex in enumerate(openIndexes):
        # MATCH:  0 : (09/10 09:06:45)打开文件：D:\欧拓图纸\切割文件\608GC样品\608GC 车架前管 A3_Φ22_T1.0_L230_X1.zx
        laserFileNameMatch = re.match(r"^\(\d+.\d+\ \d+:\d+:\d+\)打开文件：(.*)", lines[openIndex])
        # Skip current openIdex if the heading doesn't match the regex pattern
        if not laserFileNameMatch:
            print(f"[{getTimeStamp()}]:[bright_black]No laser file opened keywords found. Skip")
            continue

        laserFilePath = Path(laserFileNameMatch.groups()[0])


        # Determine whether a laserFile has been opened and completed two loops
        if str(laserFilePath) not in laserCutKeywords:
            # Initialization
            laserCutKeywords[str(laserFilePath)] = []
            laserCutKeywordsPreviousCount = 0
        else:
            laserCutKeywordsPreviousCount = len(laserCutKeywords[str(laserFilePath)])


        print(f"\n[{getTimeStamp()}]:[white]Parsing log for laser file:[/white] [bright_black]{laserFilePath}")

        # Exact exact log section for each laser file
        if i + 1 == len(openIndexes):
            # The last open index
            linesSplited = lines[openIndex:]
        else:
            linesSplited = lines[openIndex:openIndexes[i+1]]

        # Get part loop info and concatenate the line list into a string variable
        lineSplitedConcatenated = ""
        for l in linesSplited:
            # match example:
            # (09/10 09:08:00)总零件数:24, 当前零件序号:1
            partLoopMatch = re.match(r"^\((\d+.\d+) (\d+:\d+:\d+)\)总零件数:(\d+), 当前零件序号:1$", l)
            if partLoopMatch:
                partLoopDateStamp = partLoopMatch.groups()[0]
                partLoopTimeStamp = partLoopMatch.groups()[1]
                partLoopYield = partLoopMatch.groups()[2]
                laserCutKeywords[str(laserFilePath)].append(f"{partLoopDateStamp} {partLoopTimeStamp} 总零件数:{partLoopYield}，当前零件序号:1")

            # Concatenate laser cut seciion in dictionary
            lineSplitedConcatenated = lineSplitedConcatenated + l + "\n"


        # Go for next laser cut file session if current session doesn't complete two loops
        if not laserCutKeywords[str(laserFilePath)] or len(laserCutKeywords[str(laserFilePath)]) == 1:
            print(f"[{getTimeStamp()}]:[bright_black]Current laser file haven't completed two loops yet. Skip")
            continue

        os.makedirs(exportDir, exist_ok=True)
        # Write laser cut record about the first part being cut in a .txt file
        txtFilePath = Path(exportDir, laserFilePath.stem + ".txt")
        print(f"[{getTimeStamp()}]:[bold purple]Saving txt file: [/bold purple][bright_black]{txtFilePath}")
        if not txtFilePath.exists():
            txtWriteMode = "w"
            txtEncoding = "utf-8"

            with open(txtFilePath, txtWriteMode, encoding=txtEncoding) as f:
                for l in laserCutKeywords[str(laserFilePath)]:
                    f.write(f"{l}\n")
        else:
            txtWriteMode = "a"
            txtEncoding = getEncoding(txtFilePath)

            # Convert file to UTF-8
            if txtEncoding != "utf-8":
                with open(txtFilePath, "rb") as f:
                    byteContent = f.read().decode(txtEncoding)
                with open(txtFilePath, "w", encoding="utf-8") as f:
                    f.write(byteContent)

            with open(txtFilePath, txtWriteMode, encoding="utf-8") as f:
                for l in laserCutKeywords[str(laserFilePath)]:
                    f.write(f"{l}\n")


        # Split rtf file based on laserfile name
        rtfFilePath = Path(exportDir, laserFilePath.stem + ".rtf")
        print(f"[{getTimeStamp()}]:[bold blue]Saving rtf file: [/bold blue][bright_black]{rtfFilePath}")
        if not rtfFilePath.exists():
            rtfWriteMode = "w"
            rtfEncoding = "utf-8"

            with open(rtfFilePath, rtfWriteMode, encoding=rtfEncoding) as f:
                f.write(lineSplitedConcatenated)
        else:
            rtfWriteMode = "a"
            rtfEncoding = getEncoding(rtfFilePath)

            # Convert file to UTF-8
            if rtfEncoding != "utf-8":
                with open(rtfFilePath, "rb") as f:
                    byteContent = f.read().decode(rtfEncoding)
                with open(rtfFilePath, "w", encoding="utf-8") as f:
                    f.write(byteContent)

            with open(rtfFilePath, rtfWriteMode, encoding="utf-8") as f:
                f.write(lineSplitedConcatenated)


        # Generate excel file for analysis

        # Write info in gross sheet
        # Skip in current loop if current laser file info has been recorded
        grossWritenChk = False
        partWorksheet = wb["总表"]
        for row in grossWorksheet.iter_rows(min_row=1, max_col=1, max_row=grossWorksheet.max_row):
            for cell in row:
                if not cell.value:
                    break
                if cell.value == laserFilePath.name:
                    grossWritenChk = True
                    grossRowNum = cell.row
                    break

        if not grossWritenChk:
            grossRowNum = grossWorksheet.max_row + 1
            grossWorksheet[f"A{grossRowNum}"] = laserFilePath.name
            grossWorksheet[f"C{grossRowNum}"] = partLoopYield
            grossWorksheet[f"E{grossRowNum}"] = f"=D{grossRowNum}/B{grossRowNum}"
            grossWorksheet[f"E{grossRowNum}"].number_format = "h:mm:ss"
            grossWorksheet[f"F{grossRowNum}"] = 100
            grossWorksheet[f"G{grossRowNum}"] = f"=F{grossRowNum}*E{grossRowNum}"
        else:
            pass # The gross info has been written before

        grossWorksheet[f"H{grossRowNum}"] = getTimeStamp()
        grossWorksheet[f"E{grossRowNum}"].number_format = "yyyy/m/d h:mm:ss"

        # Write specific cut time in new sheet
        partWorksheetName = getProperSheetName(laserFilePath.stem)
        if laserCutKeywordsPreviousCount <= 1:
            try:
                # Even though sheet name may be duplicated after truncating
                partWorksheet = wb[partWorksheetName]
                startRow = partWorksheet.max_row + 1
            except Exception:
                partWorksheet = wb.create_sheet(partWorksheetName, 1)
                startRow = 1 #NOTE: 1 based
        else:
            partWorksheet = wb[partWorksheetName]
            startRow = partWorksheet.max_row + 1

        timeDifferences = []
        for row in partWorksheet.iter_rows(min_row=startRow, max_col=3, max_row=len(laserCutKeywords[str(laserFilePath)])):
            for cell in row:
                if cell.row == 1:
                    if cell.column_letter == "A":
                        cell.value = "时间节点"
                    if cell.column_letter == "B":
                        cell.value = "零件信息"
                    if cell.column_letter == "C":
                        cell.value = "时间差"
                else:
                    loopIdx = cell.row - 1
                    loopContent = laserCutKeywords[str(laserFilePath)][loopIdx]
                    partLoopMatch = re.match(r"^(.+) (.+) (.+)$", loopContent)
                    if partLoopMatch:
                        partLoopDateStamp = partLoopMatch.groups()[0]
                        partLoopTimeStamp = partLoopMatch.groups()[1]
                        partLoopYield     = partLoopMatch.groups()[2]

                    if cell.column_letter == "A":
                        cell.value = f"{partLoopDateStamp} {partLoopTimeStamp}"
                        cell.number_format = "yyyy/m/d h:mm:ss"
                    elif cell.column_letter == "B":
                        cell.value = partLoopYield
                    elif cell.column_letter == "C":
                        if cell.row != 2:
                            timeDifferenceFormula = f"=A{cell.row}-A{cell.row-1}"
                            timeDifferenceDatetimeObj = convertLogTimeStamp(partWorksheet[f"A{cell.row}"].value) - convertLogTimeStamp(partWorksheet[f"A{cell.row-1}"].value)
                            timeDifferenceLiteral = time.strftime("%H:%M:%S", time.gmtime(timeDifferenceDatetimeObj.total_seconds()))
                            timeDifferences.append(timeDifferenceLiteral)
                            cell.value = timeDifferenceFormula
                            cell.number_format = "h:mm:ss"


        # Write the time cost of a long tube back in gross sheet
        counter = Counter(timeDifferences)
        timeDifferenceMostCommonLiteral = counter.most_common()[0][0]
        timeDifferenceMostCommon = datetime.datetime.strptime(f"{timeDifferenceMostCommonLiteral}", "%H:%M:%S")
        for i in range(1, 6):
            timeDifferenceDelta = datetime.timedelta(seconds=i)
            timeDifferenceNew = timeDifferenceMostCommon + timeDifferenceDelta
            timeDifferenceNewLiteral = timeDifferenceNew.strftime("%H:%M:%S")

            if timeDifferenceNewLiteral not in timeDifferences:
                grossWorksheet[f"D{grossRowNum}"] = timeDifferenceNewLiteral
                grossWorksheet[f"D{grossRowNum}"].number_format = "h:mm:ss"
                break




def saveWorkbook():
    try:
        wb.save(str(speedTrackFilePath))
    except Exception as e:
        print(e)
        excelFilePath = Path(
            exportDir,
            str(
                datetime.datetime.now().strftime("%Y-%m-%d %H%M%S%f")
                ) + ".xlsx"
        )
        print(f"\n[{getTimeStamp()}]:[bold green]Saving Excel file at: [/bold green][bright_black]{excelFilePath}")
        wb.save(str(excelFilePath))

    print(f"[{getTimeStamp()}]:[bold white]Done[/bold white]")

