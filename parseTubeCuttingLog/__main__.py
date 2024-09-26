# File: parseTubeProLog
# Author: iaso2h
# Description: Parsing Log files(.rtf) from TubePro and split them into separated files
# Version: 0.0.12
# Last Modified: 2024-09-25

import console
import cli
import sys

print = console.print


if __name__ == "__main__":
    print("[bold white]此TubePro日志分析程序由阮焕编写[bold white]")
    print("[bold white]版本号: 0.0.12[bold white]")
    print("[bold white]最后更新: 2024-09-25[bold white]\n\n")
    cli.cliStart()
    # try:
    #     cli.cliStart()
    # except Exception as e:
    #     print(e)
    # print(len(sys.argv))
    input("Press enter to proceed...")
