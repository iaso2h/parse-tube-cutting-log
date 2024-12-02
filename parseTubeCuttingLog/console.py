import config
logFlow = ""

from rich.console import Console


console = Console()
def print(*args, **kwargs):
    if not config.GUI_MODE:
        if not config.SILENT_MODE:
            console.print(*args, **kwargs)
    else:
        import dearpygui.dearpygui as dpg
        global logFlow
        if logFlow == "":
            logFlow = "\n".join(args) + "\n"
        else:
            logFlow = logFlow + "\n".join(args) + "\n"
        dpg.set_value("log", value=logFlow)

