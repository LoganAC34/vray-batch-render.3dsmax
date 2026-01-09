"""
This script is a secondary window for the 3ds Max batch render dialog.
Could be used for other applications as well.

Feature wishlist:
- Rich text support
"""

import argparse
import json
import os
import subprocess
import sys
import traceback
from contextlib import contextmanager
from pathlib import Path
from typing import ContextManager

import pywintypes
import win32event
import win32file
import win32pipe
import winerror
from PySide6.QtWidgets import *

from .window import LogWindow
from .config import AppInfo, VerboseLevel, CommandType, Commands, PIPE_TIMEOUT

CMD_ERROR = False


class Console:
    """Class for console window"""

    LOG_INDENT_SPACING = "   "
    PIPE_NAME = rf"\\.\pipe\{AppInfo.APP_NAME}_{os.getpid()}"

    def __init__(self, show_log: bool = False) -> None:
        self._verbose_level = VerboseLevel.INFO
        self._indent_level = 0
        self.show_log = show_log
        self.pipe_path = self._setup_log()

    # Log functions
    def _setup_log(self) -> int:
        """Initializes and starts the log window
        Returns:
            pipe_handle: Pipe path
        """
        # noinspection PyTypeChecker
        pipe_handle = win32pipe.CreateNamedPipe(
            self.PIPE_NAME,
            win32pipe.PIPE_ACCESS_DUPLEX | win32file.FILE_FLAG_OVERLAPPED,
            win32pipe.PIPE_TYPE_MESSAGE
            | win32pipe.PIPE_READMODE_MESSAGE
            | win32pipe.PIPE_WAIT,
            1,
            512,
            512,
            10,
            None,
        )

        overlapped = pywintypes.OVERLAPPED()
        overlapped.hEvent = win32event.CreateEvent(None, 0, 0, None)

        # Start the log window
        creation_flags = (
            subprocess.CREATE_NEW_CONSOLE if self.show_log else subprocess.CREATE_NO_WINDOW
        )
        script_path = Path(__file__).absolute()
        # noinspection PyUnresolvedReferences
        arguments = [sys.executable, str(script_path), "--pipe", self.PIPE_NAME]

        subprocess.Popen(arguments, creationflags=creation_flags)

        try:
            win32pipe.ConnectNamedPipe(pipe_handle, overlapped)
        except pywintypes.error as e:
            if e.winerror != winerror.ERROR_IO_PENDING:
                raise

        # Wait for 3000ms (3 seconds)
        wait_result = win32event.WaitForSingleObject(overlapped.hEvent, PIPE_TIMEOUT)

        if wait_result == win32event.WAIT_TIMEOUT:
            win32file.CloseHandle(pipe_handle)
            raise Exception("Pipe connection timed out")
        elif wait_result == win32event.WAIT_OBJECT_0:
            return pipe_handle
        raise Exception("Failed to connect to named pipe")

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
    def _check_command_type_and_data(data_type: CommandType, data: str | Commands):
        """Checks if the data type and command are valid.
        Args:
            data_type: Data type
            data: Data to send

        Raises:
            ValueError: If the data type or command is not valid.
        """
        if data_type not in list(CommandType):
            raise ValueError(f"Invalid data type: {data_type}")

        if data_type == CommandType.COMMAND and data not in list(Commands):
            raise ValueError(f"Invalid command: {data}")

    def _write_to_pipe(self, data_type: CommandType, data_str: str | Commands):
        """Writes data to the pipe.
        Args:
            data_str: Data to send
            data_type: Data type
        """
        self._check_command_type_and_data(data_type, data_str)
        data_json = json.dumps({"TYPE": data_type.value, "VALUE": data_str})
        win32file.WriteFile(self.pipe_path, (data_json + "\n").encode("utf-8"))

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
            self._write_to_pipe(CommandType.LOG, message)

    def open(self) -> None:
        """Restores the log window if it was closed."""
        self._write_to_pipe(CommandType.COMMAND, Commands.OPEN)

    def shutdown(self) -> None:
        """Shut down the log window."""
        self._write_to_pipe(CommandType.COMMAND, Commands.SHUTDOWN)
        win32file.CloseHandle(self.pipe_path)


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
    console = Console(True)
    console.log(VerboseLevel.DEBUG, "Debug mode enabled")
    with console.indent():
        console.log(VerboseLevel.DEBUG, "Debug mode enabled")
        with console.indent():
            console.log(VerboseLevel.DEBUG, "Debug mode enabled")
            console.log(VerboseLevel.DEBUG, "Debug mode enabled")
    console.log(VerboseLevel.DEBUG, "Debug mode enabled")
    console.open()
    console.shutdown()
    pass


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--debug", action="store_true")
    parser.add_argument("--show", action="store_true")
    parser.add_argument("--pipe", type=str)
    args = parser.parse_args()

    if args.debug:
        test()
    else:
        app = QApplication()
        window = LogWindow(args.pipe)
        window.show()
        sys.exit(app.exec())
