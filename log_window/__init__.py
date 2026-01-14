"""
Logging package for 3ds Max batch rendering.

This package provides a console-like logging interface with a separate window
for displaying log messages with different verbosity levels.

Example usage:
    from log import Console, VerboseLevel

    # Create a console instance
    console = Console(show_log=True)

    # Log messages at different levels
    console.log(VerboseLevel.INFO, "This is an info message")
    console.log(VerboseLevel.ERROR, "This is an error message")

    # Set verbosity level
    console.verbose_level = VerboseLevel.DEBUG

    # Shutdown the console when done
    console.shutdown()
"""

import sys
from typing import Optional

from .main import Console, excepthook
from .config import VerboseLevel, AppInfo

# Initialize console as None and provide a getter function
_console: Optional[Console] = None


def get_console() -> Console:
    """Get or create the console instance."""
    global _console
    if _console is None:
        print("Initializing console...")
        _console = Console()
    return _console


# Create a property-like accessor for backward compatibility
console = get_console()

__version__ = AppInfo.APP_VERSION
__author__ = AppInfo.AUTHOR

# Set up the exception hook by default
sys.excepthook = excepthook

__all__ = [
    "console",
    "get_console",
    "VerboseLevel",
    "AppInfo",
]
