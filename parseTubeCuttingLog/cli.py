import rtfParse
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
    for f in rtfParse.programDir.iterdir():
        if f.suffix == ".rtf":
            rtfParse.rtfCandidates.append(f)


def cliStart():
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
