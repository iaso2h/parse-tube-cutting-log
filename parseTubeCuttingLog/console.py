import config

from rich.console import Console


console = Console()
def print(*args, **kwargs):
    if not config.SILENTMODE:
        console.print(*args, **kwargs)
