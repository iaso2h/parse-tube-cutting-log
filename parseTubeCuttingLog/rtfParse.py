import console
import chardet
import os
import re
import datetime
from pathlib import Path
from striprtf.striprtf import rtf_to_text
from openpyxl import Workbook


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
    if len(name) > 31:
        return name[:31]
    else:
        return name


laserCutKeywords = {}


def getEncoding(filePath):
    # Create a magic object
    with open(filePath, "rb") as f:
        # Detect the encoding
        rawData = f.read()
        result = chardet.detect(rawData)
        return result["encoding"]


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


    # Record how many laser files have been skipped if they haven't completed two loops
    skipCount = 0
    # Loop through each laser cut file section starting with keywords about opening a file
    for i, openIndex in enumerate(openIndexes):
        # MATCH:  0 : (09/10 09:06:45)打开文件：D:\欧拓图纸\切割文件\608GC样品\608GC 车架前管 A3_Φ22_T1.0_L230_X1.zx
        laserFileNameMatch = re.match(r"^\(\d+.\d+\ \d+:\d+:\d+\)打开文件：(.*)", lines[openIndex])
        # Skip current openIdex if the heading doesn't match the regex pattern
        if not laserFileNameMatch:
            skipCount = skipCount + 1
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
            skipCount = skipCount + 1
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
        rowNum = i + 1 - skipCount
        global excelCreatedChk
        if not excelCreatedChk:
            excelCreatedChk = True
            ws = wb.active
            ws.title = "总表(复制用)"
        else:
            ws = wb["总表(复制用)"]


        # Write info in gross sheet

        # Skip in current loop if current laser file info has been recorded
        grossWritenChk = False
        for row in partWorksheet.iter_rows(min_row=1, max_col=1, max_row=len(openIndexes)+1:
            for cell in row:
                if not cell.value:
                    break
                if cell.value == laserFilePath.name:
                    grossWritenChk = True
                    break

        if grossWritenChk:
            continue

        if laserCutKeywordsPreviousCount <= 1:
            ws[f"A{rowNum}"] = laserFilePath.name
            ws[f"C{rowNum}"] = partLoopYield
            ws[f"E{rowNum}"] = f"=D{rowNum}/B{rowNum}"
            ws[f"F{rowNum}"] = 100
            ws[f"G{rowNum}"] = f"=F{rowNum}*E{rowNum}"
        else:
            pass #skip when gross info has been written before

        ws[f"H{rowNum}"] = getTimeStamp()

        # Write specific cut time in new sheet
        partWorksheetName = getProperSheetName(laserFilePath.stem)
        if laserCutKeywordsPreviousCount <= 1:
            try:
                # Even though sheet name may be duplicated after truncating
                partWorksheet = wb[partWorksheetName]
                startRow = laserCutKeywordsPreviousCount + 1
            except Exception:
                partWorksheet = wb.create_sheet(partWorksheetName, 1)
                startRow = 1 #NOTE: 1 based
        else:
            partWorksheet = wb[partWorksheetName]
            startRow = laserCutKeywordsPreviousCount + 1

        for row in partWorksheet.iter_rows(min_row=startRow, max_col=4, max_row=len(laserCutKeywords[str(laserFilePath)])):
            for cell in row:
                if cell.row == 1:
                    if cell.column_letter == "D":

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
                        cell.value = partLoopDateStamp
                    elif cell.column_letter == "B":
                        cell.value = partLoopTimeStamp
                    elif cell.column_letter == "C":
                        cell.value = partLoopYield
                    elif cell.column_letter == "D":
                        if cell.row != 2:
                            cell.value = f"=B{cell.row}-B{cell.row-1}"
                            cell.number_format = "h:mm:ss"


def saveWorkbook():
    if excelCreatedChk:
        excelFilePath = Path(
            exportDir,
            str(
                datetime.datetime.now().strftime("%Y-%m-%d %H%M%S%f")
                ) + ".xlsx"
        )
        print(f"\n[{getTimeStamp()}]:[bold green]Saving Excel file at: [/bold green][bright_black]{excelFilePath}")
        # create a gross sheet
        wb.save(str(excelFilePath))

        print(f"[{getTimeStamp()}]:[bold white]Done[/bold white]")

