import rtfParse
import config
import console

import sys
import beaupy
import beaupy.spinners as sp
from rich.console import Console


print = console.print
spinnerParsing = sp.Spinner(sp.DOTS, "Parsing...\n")

def keyboardInterruptExit() -> None:
    print("[red]Interrupt by user[/red]")
    raise SystemExit(1)


def abortExit() -> None:
    print("[red]Abort by user[/red]")
    raise SystemExit(1)


def rtfFind():
    # Find rtf file to read
    for f in config.PROGRAMDIR.iterdir():
        if f.suffix == ".rtf":
            rtfParse.rtfCandidates.append(f)


def speedTrack():
    # Select rtf files to parse
    if len(sys.argv) > 1 and sys.argv[1][-4:] == ".rtf":
            rtfParse.rtfTarget = sys.argv[1]
    else:
        rtfFind()
        if len(rtfParse.rtfCandidates) == 0:
            return print("[red]No available .rtf files[/red]")
        else:
            if len(rtfParse.rtfCandidates) == 1:
                rtfParse.rtfTarget = rtfParse.rtfCandidates[0]
            else:
                print("[white]Please Select .rtf file[/white]")
                rtfCandidates = [str(p) for p in rtfParse.rtfCandidates]

                try:
                    rtfParse.rtfTarget = beaupy.select(rtfCandidates, return_index=False)
                except KeyboardInterrupt:
                    keyboardInterruptExit()
                except beaupy.Abort:
                    abortExit()
                except Exception as e:
                    print(e)
                    SystemExit(1)

    # spinnerParsing.start()
    rtfParse.parseStart()
    rtfParse.saveWorkbook()
    # spinnerParsing.stop()


def cutRecord():
    import cutRecord
    cutRecord.addScreendshots()


def dispatchFill():
    import dispatch
    dispatch.fillExcel()


def cliStart():
    functions = ["日志分析", "开料记录", "填派工单"]
    try:
        ans = beaupy.select(functions, return_index=False)
    except KeyboardInterrupt:
        keyboardInterruptExit()
    except beaupy.Abort:
        abortExit()
    except Exception as e:
        print(e)
        SystemExit(1)

    if ans == "开料记录":
        cutRecord()
    elif ans == "日志分析":
        speedTrack()
    elif ans == "填派工单":
        dispatchFill()
