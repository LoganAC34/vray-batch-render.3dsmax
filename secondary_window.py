"""
This script is a secondary window for the 3ds Max batch render dialog.
Could be used for other applications as well.

Feature wishlist:
- Rich text support
"""
import winreg

import pywintypes

PUBLISHER = 'OrangeByte'
PUBLISHER_CREATOR = 'Logan Carrozza'
AUTHOR = f'{PUBLISHER_CREATOR} w/ assistance of GPT'
APP_NAME = 'Log Window'
APP_VERSION = '1.0.0'
""" Semantic Versioning https://semver.org/
MAJOR version when you make incompatible API changes
MINOR version when you adjustment functionality in a backward compatible manner
PATCH version when you make backward compatible bug fixes
"""

import json
import os
import subprocess
import sys
import time
import traceback

from PySide6.QtCore import QThread, Qt, QObject, Signal, Slot, QSettings
from PySide6.QtGui import QTextOption, QTextCursor, QIcon, QPainter, QPixmap
from PySide6.QtWidgets import *

current_script_dir = os.path.dirname(__file__)
venv = os.path.join(current_script_dir, '..', '.venv')
sys.path.append(venv)

import win32file
import win32pipe

pipe_name = sys.argv[1]
CMD_ERROR = False


class Worker(QObject):
    finished = Signal()
    received_message = Signal(str)

    @Slot()
    def run(self):
        abort = False
        while not abort:
            try:
                handle = win32file.CreateFile(
                    pipe_name,
                    win32file.GENERIC_READ | win32file.GENERIC_WRITE,
                    0,
                    None,
                    win32file.OPEN_EXISTING,
                    0,
                    None
                )
                res = win32pipe.SetNamedPipeHandleState(handle, win32pipe.PIPE_READMODE_MESSAGE, None, None)
                if res == 0:
                    print(f"SetNamedPipeHandleState return code: {res}")
                while True:
                    json_objects = win32file.ReadFile(handle, 64 * 1024)[1].decode('utf-8')
                    json_objects = json_objects.split('\n')

                    for data in json_objects:
                        if data:
                            dat_dict = json.loads(data)
                            data_type = dat_dict['TYPE']
                            text = dat_dict['VALUE']

                            self.received_message.emit(data)

                            if data_type == 'COMMAND' and text == 'SHUTDOWN':
                                abort = True
                                self.finished.emit()
                                break

            except pywintypes.error as e:
                if e.args[0] == 2:
                    # no pipe, trying again
                    time.sleep(0.1)
                elif e.args[0] == 109:
                    # broken pipe
                    break

class LogWindow(QMainWindow):
    GRIP_SIZE = 9
    TITLE_BAR_COLOR = '#808080'
    WINDOW_BACKGROUND_COLOR = '#2b2b2b'
    DEBUG_COLORS = False
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Batch Print Log')
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        self.setStyleSheet(f"background-color: {self.WINDOW_BACKGROUND_COLOR};")
        self.setMinimumSize(500, 200)
        self.setGeometry(100, 100, 725, 200)
        self.restoreWindowSettings()
        self.offset = None

        self.max_icon = QIcon('C:/Program Files/Autodesk/3ds Max 2025/icons/icon_main.ico')
        self._set_icon()
        self._build_ui()

        # Receive log message thread
        # https://stackoverflow.com/questions/6783194/background-thread-with-qthread-in-pyqt
        self.worker = Worker()
        self.thread = QThread()
        self.worker.received_message.connect(self.handle_message)
        self.worker.moveToThread(self.thread)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.exit)
        self.thread.started.connect(self.worker.run)
        self.thread.start()

    def _set_icon(self):
        # Get 3ds Max icon
        console_icon = QIcon(r"C:\Windows\Installer\{C1593F76-F694-448E-AD35-82DDD6203975}\PowerShellExe.ico")

        max_icon_pixmap = self.max_icon.pixmap(24, 24)
        console_icon_pixmap = console_icon.pixmap(12, 12)

        # Create a new pixmap with the same size as self.max_icon
        new_pixmap = QPixmap(max_icon_pixmap.width(), max_icon_pixmap.height())
        new_pixmap.fill(Qt.GlobalColor.transparent)

        # Paint the self.max_icon on the new pixmap
        painter = QPainter(new_pixmap)
        painter.drawPixmap(0, 0, max_icon_pixmap)
        painter.drawPixmap(max_icon_pixmap.width() - console_icon_pixmap.width(),
                           max_icon_pixmap.height() - console_icon_pixmap.height(),
                           console_icon_pixmap)
        painter.end()

        # Set the new pixmap as the window icon
        self.setWindowIcon(QIcon(new_pixmap))

    # noinspection LongLine
    def _build_ui(self):
        # Title bar icon
        icon_size = 16 # same W & H
        self.icon_label = QLabel(self)
        self.icon_label.setPixmap(self.max_icon.pixmap(icon_size, icon_size))
        self.icon_label.setFixedSize(icon_size, icon_size)
        self.icon_label.mousePressEvent = self.titlebarMousePressEvent
        self.icon_label.mouseMoveEvent = self.titlebarMouseMoveEvent

        # Title bar title
        self.title = QLabel(self)
        self.title.setText('  ' + self.windowTitle())
        self.title.setStyleSheet(f"color: {self.TITLE_BAR_COLOR}")
        self.title.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.title.mousePressEvent = self.titlebarMousePressEvent
        self.title.mouseMoveEvent = self.titlebarMouseMoveEvent

        # X Close button
        self.closeButton_size = 29
        self.closeButton = QPushButton(u'\U00010317', self)
        self.closeButton.setFlat(True)
        self.closeButton.clicked.connect(self.close_window)
        self.closeButton.setFixedSize(self.closeButton_size, self.closeButton_size)
        self.closeButton.setStyleSheet(f"background-color: {self.WINDOW_BACKGROUND_COLOR}; "
                                       f"color: {self.TITLE_BAR_COLOR}; "
                                       f"font: 12pt 'Segoe UI Historic'; text-align: center;")

        # Log
        self.log = QPlainTextEdit(self)
        self.log.setReadOnly(True)
        self.log.setWordWrapMode(QTextOption.WrapMode.NoWrap)
        self.log.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.log.setStyleSheet("background-color: #5a5a5a; color: White;")

        # https://stackoverflow.com/questions/62807295/how-to-resize-a-window-from-the-edges-after-adding-the-property-qtcore-qt-framel
        # ^Not really relevant anymore, but had some inspiration^
        # Grips
        self.grip_top = QWidget()
        self.grip_top.setFixedHeight(self.GRIP_SIZE)
        self.grip_top.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        if self.DEBUG_COLORS: self.grip_top.setStyleSheet(f"background-color: green;")

        self.grip_left = QWidget()
        self.grip_left.setFixedWidth(self.GRIP_SIZE)
        if self.DEBUG_COLORS: self.grip_left.setStyleSheet(f"background-color: yellow;")

        self.grip_right = QWidget()
        self.grip_right.setFixedWidth(self.GRIP_SIZE)
        if self.DEBUG_COLORS: self.grip_right.setStyleSheet(f"background-color: white;")

        self.grip_bottom = QWidget()
        self.grip_bottom.setFixedHeight(self.GRIP_SIZE)
        if self.DEBUG_COLORS: self.grip_bottom.setStyleSheet(f"background-color: purple;")

        self.grip_topRight = QWidget()
        self.grip_topRight.setFixedSize(self.GRIP_SIZE, self.GRIP_SIZE)
        if self.DEBUG_COLORS: self.grip_topRight.setStyleSheet(f"background-color: orange;")

        self.grip_topLeft = QWidget()
        self.grip_topLeft.setFixedSize(self.GRIP_SIZE, self.GRIP_SIZE)
        if self.DEBUG_COLORS: self.grip_topLeft.setStyleSheet(f"background-color: blue;")

        self.grip_bottomRight = QWidget()
        self.grip_bottomRight.setFixedSize(self.GRIP_SIZE, self.GRIP_SIZE)
        if self.DEBUG_COLORS: self.grip_bottomRight.setStyleSheet(f"background-color: orange;")

        self.grip_bottomLeft = QWidget()
        self.grip_bottomLeft.setFixedSize(self.GRIP_SIZE, self.GRIP_SIZE)
        if self.DEBUG_COLORS: self.grip_bottomLeft.setStyleSheet(f"background-color: blue;")

        # Layout - Window title bar (ORDER MATTERS!)
        windowTitle_And_ResizeGrips = QGridLayout()
        windowTitle_And_ResizeGrips.setSpacing(0)
        windowTitle_And_ResizeGrips.setContentsMargins(0, 0, 0, 0)
        windowTitle_And_ResizeGrips.setColumnStretch(1, 1)
        windowTitle_And_ResizeGrips.setColumnStretch(2, 1)
        windowTitle_And_ResizeGrips.setColumnMinimumWidth(3, 30)
        windowTitle_And_ResizeGrips.setRowMinimumHeight(0, 30)
        windowTitle_And_ResizeGrips.setRowStretch(2, 1)
        windowTitle_And_ResizeGrips.addWidget(self.icon_label, 0, 1)
        windowTitle_And_ResizeGrips.addWidget(self.title, 0, 2, 1, 3)
        windowTitle_And_ResizeGrips.addWidget(self.grip_left, 0, 0, 3, 0, alignment=Qt.AlignmentFlag.AlignLeft)
        windowTitle_And_ResizeGrips.addWidget(self.grip_right, 0, 3, 3, 3, alignment=Qt.AlignmentFlag.AlignRight)
        windowTitle_And_ResizeGrips.addWidget(self.grip_top, 0, 1, 0, 3, alignment=Qt.AlignmentFlag.AlignTop)
        windowTitle_And_ResizeGrips.addWidget(self.grip_topLeft, 0, 0, alignment=Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        windowTitle_And_ResizeGrips.addWidget(self.grip_topRight, 0, 3, alignment=Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight)
        windowTitle_And_ResizeGrips.addWidget(self.grip_bottom, 4, 1, 4, 3, alignment=Qt.AlignmentFlag.AlignTop)
        windowTitle_And_ResizeGrips.addWidget(self.grip_bottomLeft, 4, 0, alignment=Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignLeft)
        windowTitle_And_ResizeGrips.addWidget(self.grip_bottomRight, 4, 3, alignment=Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignRight)
        windowTitle_And_ResizeGrips.addWidget(self.closeButton, 0, 3, alignment=Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignLeft)


        # Layout - Log and title bar
        self.content_layout = QVBoxLayout()
        self.content_layout.addWidget(self.log)

        window_layout = QGridLayout()
        window_layout.setContentsMargins(0, 0, 0, 0)
        window_layout.setSpacing(0)
        window_layout.addLayout(windowTitle_And_ResizeGrips, 0, 0, 3, 3)
        window_layout.setRowStretch(1, 1)
        window_layout.setColumnStretch(1, 1)
        window_layout.setColumnMinimumWidth(0, self.GRIP_SIZE)
        window_layout.setColumnMinimumWidth(2, self.GRIP_SIZE)
        window_layout.setRowMinimumHeight(0, self.closeButton.height() + 1)
        window_layout.setRowMinimumHeight(2, self.GRIP_SIZE)
        window_layout.addLayout(self.content_layout, 1, 1)

        central_widget = QWidget()
        central_widget.setLayout(window_layout)
        self.setCentralWidget(central_widget)

        # Mouse hover events
        self.grip_left.enterEvent = lambda event: self.setCursor(Qt.CursorShape.SizeHorCursor)
        self.grip_top.enterEvent = lambda event: self.setCursor(Qt.CursorShape.SizeVerCursor)
        self.grip_right.enterEvent = lambda event: self.setCursor(Qt.CursorShape.SizeHorCursor)
        self.grip_bottom.enterEvent = lambda event: self.setCursor(Qt.CursorShape.SizeVerCursor)
        self.grip_topLeft.enterEvent = lambda event: self.setCursor(Qt.CursorShape.SizeFDiagCursor)
        self.grip_topRight.enterEvent = lambda event: self.setCursor(Qt.CursorShape.SizeBDiagCursor)
        self.grip_bottomLeft.enterEvent = lambda event: self.setCursor(Qt.CursorShape.SizeBDiagCursor)
        self.grip_bottomRight.enterEvent = lambda event: self.setCursor(Qt.CursorShape.SizeFDiagCursor)
        self.title.enterEvent = lambda event: self.setCursor(Qt.CursorShape.ArrowCursor)
        self.closeButton.enterEvent = lambda event: self.setCursor(Qt.CursorShape.ArrowCursor)
        self.icon_label.enterEvent = lambda event: self.setCursor(Qt.CursorShape.ArrowCursor)

        self.grip_topLeft.mouseMoveEvent = lambda event: self.gripMoveEvent(event, self.grip_topLeft)
        self.grip_topRight.mouseMoveEvent = lambda event: self.gripMoveEvent(event, self.grip_topRight)
        self.grip_bottomLeft.mouseMoveEvent = lambda event: self.gripMoveEvent(event, self.grip_bottomLeft)
        self.grip_bottomRight.mouseMoveEvent = lambda event: self.gripMoveEvent(event, self.grip_bottomRight)
        self.grip_left.mouseMoveEvent = lambda event: self.gripMoveEvent(event, self.grip_left)
        self.grip_top.mouseMoveEvent = lambda event: self.gripMoveEvent(event, self.grip_top)
        self.grip_right.mouseMoveEvent = lambda event: self.gripMoveEvent(event, self.grip_right)
        self.grip_bottom.mouseMoveEvent = lambda event: self.gripMoveEvent(event, self.grip_bottom)

    def restoreWindowSettings(self):
        # Restore window position & size
        settings = QSettings(PUBLISHER, APP_NAME)
        # noinspection PyTypeChecker
        self.restoreGeometry(settings.value("geometry"))

    def saveWindowSettings(self):
        # Save publisher creator
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, f"Software\\{PUBLISHER}", 0, winreg.KEY_WRITE)
        winreg.SetValueEx(key, "creator", 0, winreg.REG_SZ, PUBLISHER_CREATOR)
        winreg.CloseKey(key)

        # Save window position & size
        settings = QSettings(PUBLISHER, APP_NAME)
        settings.setValue("geometry", self.saveGeometry())
        settings.setValue("author", AUTHOR)
        settings.setValue("version", APP_VERSION)

    def gripMoveEvent(self, event, grip):
        new_x = event.globalPosition().x()
        new_y = event.globalPosition().y()

        # Left Side
        if grip in [self.grip_topLeft, self.grip_left, self.grip_bottomLeft]:
            new_width = self.width() - (new_x - self.pos().x())
            if new_width > self.minimumWidth():
                self.move(new_x, self.pos().y())
                self.resize(new_width, self.height())

        # Top Side
        if grip in [self.grip_topLeft, self.grip_top, self.grip_topRight]:
            new_height = self.height() - (new_y - self.pos().y())
            if new_height > self.minimumHeight():
                self.move(self.pos().x(), new_y)
                self.resize(self.width(), new_height)

        # Right Side
        if grip in [self.grip_topRight, self.grip_right, self.grip_bottomRight]:
            new_width = self.width() - (self.pos().x() + self.width() - new_x)
            self.resize(new_width, self.height())

        if grip in [self.grip_bottomLeft, self.grip_bottom, self.grip_bottomRight]:
            new_height = self.height() - (self.pos().y() + self.height() - new_y)
            self.resize(self.width(), new_height)

    def titlebarMousePressEvent(self, event):
        self.offset = -(self.pos() - event.globalPosition().toPoint())

    def titlebarMouseMoveEvent(self, event):
        delta = event.globalPosition() - self.offset
        self.move(delta.toPoint())

    def handle_message(self, data):
        """
        Handles messages from the other application.
        If you adjustment
        :param data:
        :type data:
        :return:
        :rtype:
        """
        data = json.loads(data)
        data_type = data['TYPE']
        data_value = data['VALUE']

        if data_type == 'LOG':
            print(data_value)
            label_text = self.log.toPlainText()
            label_text += data_value + '\n'

            self.log.setPlainText(label_text)
            self.log.moveCursor(QTextCursor.MoveOperation.End)

        elif data_type == 'COMMAND':
            if data_value == 'OPEN':
                self.hide()
                self.show()

    def close_window(self):
        self.hide()

    def exit(self):
        self.saveWindowSettings()
        self.thread.quit()
        sys.exit()

def excepthook(error_type, error, traceback_str):
    def escape_characters(text):
        text = str(text)
        characters = ['^', '&','|', '<', '>'] # '^' needs to be first in list
        for character in characters:
            text = text.replace(character, f'^{character}')
        text = text.replace('"', "'")

        text = text.split('\n')
        for x, line in enumerate(text):
            if line == '':
                text[x] = '.'
            else:
                text[x] = f' {line}'

        return text

    # Create a new cmd window and display the error message
    global CMD_ERROR
    if not CMD_ERROR:
        traceback_str = ''.join(traceback.format_tb(error.__traceback__))
        traceback_str = escape_characters(traceback_str)
        error_type = escape_characters(str(error_type.__name__))[0]
        error_str = escape_characters(error)
        error_str[0] = error_type + ': ' + error_str[0]
        command = ['Traceback:'] + traceback_str +  error_str
        command = 'echo ' + ' & echo'.join(command)
        args = ['cmd', '/k', command]

        subprocess.Popen(args, creationflags=subprocess.CREATE_NEW_CONSOLE)
        CMD_ERROR = True
        raise error


def main():
    app = QApplication(sys.argv)
    LogWindow()
    sys.exit(app.exec())


if __name__ == '__main__':
    sys.excepthook = excepthook
    main()
