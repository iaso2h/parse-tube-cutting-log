import config

from rich.console import Console


console = Console()
def print(*args, **kwargs):
    if not config.SILENT_MODE:
        console.print(*args, **kwargs)
