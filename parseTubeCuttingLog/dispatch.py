import config
import console

import re
from pathlib import Path

print = console.print
laserCutFileParentPath = Path(r"D:\欧拓图纸\切割文件")

laserFilePaths = []
def getAllLaserFiles():
    if not laserCutFileParentPath.exists():
        return

    for p in laserCutFileParentPath.iterdir():
        if p.is_file():
            laserFilePaths.append(p)


def fillExcel():
    getAllLaserFiles()
    for p in laserFilePaths:
        fileNameMatch = re.match(r"^(\d{3}[a-zA-Z]{0,2})\s+?(.*?)(?=A3)|^(\d{3}[a-zA-Z]{0,2})\s+?(.*?)(?=6063(-T\d)?)|^(\d{3}[a-zA-Z]{0,2})\s+?(.*?)(?=6061(-T\d)?)", str(p.stem))
        if fileNameMatch:
            print(fileNameMatch.group(1), fileNameMatch.group(2))
        # else:
            # print("miss match for:", p.name)
