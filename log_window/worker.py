"""Worker thread for receiving log messages."""

import json
import time

import pywintypes
import win32file
import win32pipe
import winerror
from PySide6.QtCore import QObject, Signal, Slot

from .config import CommandType, Commands, PIPE_TIMEOUT


class Worker(QObject):
    """Worker thread for receiving log messages."""

    finished = Signal()
    received_message = Signal(str)

    def __init__(self, pipe_name):
        super().__init__()
        self.pipe_name = pipe_name
        self.pipe_handle = None

    def is_connection_active(self):
        """Checks if the named pipe connection is still valid"""
        try:
            win32pipe.PeekNamedPipe(self.pipe_handle, 0)
            return True
        except pywintypes.error:
            return False

    @Slot()
    def run(self):
        """Runs the worker thread."""
        pipe_handle = self.connect_to_pipe()

        while True:
            time.sleep(0.01)
            _, data = win32file.ReadFile(pipe_handle, 64 * 1024)
            # noinspection PyUnresolvedReferences
            json_objects = data.decode("utf-8").split("\n")
            print(json_objects)

            for data in json_objects:
                if data:
                    dat_dict = json.loads(data)
                    data_type = dat_dict["TYPE"]
                    text = dat_dict["VALUE"]

                    self.received_message.emit(data)

                    if data_type == CommandType.COMMAND and text == Commands.SHUTDOWN:
                        self.finished.emit()
                        break

    def connect_to_pipe(self) -> int:
        """Connects to the named pipe.
        Returns:
            int: The handle to the named pipe.
        """
        try:
            win32pipe.WaitNamedPipe(self.pipe_name, PIPE_TIMEOUT)
        except pywintypes.error as e:
            if e.winerror == winerror.ERROR_SEM_TIMEOUT:
                # Run your failure function here
                # sys.exit()
                pass

        self.pipe_handle = win32file.CreateFile(
            self.pipe_name,
            win32file.GENERIC_READ | win32file.GENERIC_WRITE,
            0,
            None,
            win32file.OPEN_EXISTING,
            0,
            None,
        )

        win32pipe.SetNamedPipeHandleState(
            self.pipe_handle.handle, win32pipe.PIPE_READMODE_MESSAGE, None, None
        )

        return self.pipe_handle.handle
