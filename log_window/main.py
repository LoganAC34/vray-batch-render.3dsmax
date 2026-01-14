"""
This script is a secondary window for the 3ds Max batch render dialog.
Could be used for other applications as well.

Feature wishlist:
- Rich text support
"""

import os
import subprocess
import traceback
from contextlib import contextmanager
from typing import ContextManager

CMD_ERROR = False

from .config import AppInfo, VerboseLevel
from .window import LogWindow


class Console:
    """Class for console window"""

    LOG_INDENT_SPACING = "   "
    PIPE_NAME = rf"\\.\pipe\{AppInfo.APP_NAME}_{os.getpid()}"

    def __init__(self) -> None:
        self.window = LogWindow()
        self._verbose_level = VerboseLevel.INFO
        self._indent_level = 0
        print("Console was run")

    # noinspection PyTypeChecker
    @contextmanager
    def indent(self, ignore: bool = False) -> ContextManager[None]:
        """Increases the indent level by one."""
        if ignore:
            yield
            return

        self._indent_level += 1
        try:
            yield
        finally:
            self._indent_level -= 1

    @staticmethod
    def _verbose_level_check(level: VerboseLevel) -> None:
        """Checks if the verbose level is valid.
        Args:
            level: Verbose level

        Raises:
            ValueError: If the verbose level is not valid.
        """
        if level not in list(VerboseLevel):
            raise ValueError("Invalid verbose level")

    @property
    def verbose_level(self) -> VerboseLevel:
        """Returns the current verbose level."""
        return self._verbose_level

    @verbose_level.setter
    def verbose_level(self, level: VerboseLevel) -> None:
        """Sets the verbose level.
        Args:
            level: The verbose level.
        """
        self._verbose_level_check(level)
        self._verbose_level = level

    def log(self, level: VerboseLevel, text: str) -> None:
        """String to send to log window.
        Args:
            level: Verbose level of message.
            text: Message
        """
        self._verbose_level_check(level)
        if level >= self._verbose_level:
            message = (self.LOG_INDENT_SPACING * self._indent_level) + str(text)
            self.window.log_message(message)

    def open(self) -> None:
        """Restores the log window if it was closed."""
        self.window.hide()
        self.window.show()

    def close(self) -> None:
        """Shut down the log window."""
        self.window.close()


def excepthook(error_type, error, traceback_str):
    """Handles the exception hook"""

    def escape_characters(text):
        """Escapes the characters in the text"""
        text = str(text)
        characters = ["^", "&", "|", "<", ">"]  # '^' needs to be first in list
        for character in characters:
            text = text.replace(character, f"^{character}")
        text = text.replace('"', "'")

        text = text.split("\n")
        for x, line in enumerate(text):
            if line == "":
                text[x] = "."
            else:
                text[x] = f" {line}"

        return text

    # Create a new cmd window and display the error message
    global CMD_ERROR
    if not CMD_ERROR:
        traceback_str = "".join(traceback.format_tb(error.__traceback__))
        traceback_str = escape_characters(traceback_str)
        error_type = escape_characters(str(error_type.__name__))[0]
        error_str = escape_characters(error)
        error_str[0] = error_type + ": " + error_str[0]
        command = ["Traceback:"] + traceback_str + error_str
        command = "echo " + " & echo".join(command)
        arguments = ["cmd", "/k", command]

        subprocess.Popen(arguments, creationflags=subprocess.CREATE_NEW_CONSOLE)
        CMD_ERROR = True
        raise error


def test():
    console = Console()
    console.log(VerboseLevel.DEBUG, "Debug mode enabled")
    with console.indent():
        console.log(VerboseLevel.DEBUG, "Debug mode enabled")
        with console.indent():
            console.log(VerboseLevel.DEBUG, "Debug mode enabled")
            console.log(VerboseLevel.DEBUG, "Debug mode enabled")
    console.log(VerboseLevel.DEBUG, "Debug mode enabled")
    console.open()
    console.close()
    pass


if __name__ == "__main__":
    # test() # imports broken
    pass
