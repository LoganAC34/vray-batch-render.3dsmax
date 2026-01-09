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

from .main import Console, excepthook, test
from .config import VerboseLevel, CommandType, Commands, AppInfo

__version__ = AppInfo.APP_VERSION
__author__ = AppInfo.AUTHOR

# Set up the exception hook by default
sys.excepthook = excepthook

__all__ = [
    "Console",
    "VerboseLevel",
    "CommandType",
    "Commands",
    "AppInfo",
    "excepthook",
    "test",
]
