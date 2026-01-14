"""Log window class"""

import os
import sys
import winreg

from PySide6 import QtWidgets
from PySide6.QtCore import Qt, QSettings
from PySide6.QtGui import QIcon, QPixmap, QPainter, QTextOption, QTextCursor
from PySide6.QtWidgets import (
    QLabel,
    QSizePolicy,
    QPushButton,
    QPlainTextEdit,
    QWidget,
    QGridLayout,
    QVBoxLayout,
)

project_root = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, project_root)
os.chdir(project_root)

if __name__ == "__main__":
    from config import AppInfo

    MAX_VERSION = 2026
else:
    from .config import AppInfo

    import pymxs

    rt = pymxs.runtime
    MAX_VERSION = rt.maxVersion()[-2]


class LogWindow(QtWidgets.QDialog):
    """Log window class"""

    GRIP_SIZE = 9
    TITLE_BAR_COLOR = "#808080"
    WINDOW_BACKGROUND_COLOR = "#2b2b2b"
    DEBUG_COLORS = False
    max_icon = f"C:/Program Files/Autodesk/3ds Max {MAX_VERSION}/icons/icon_main.ico"
    console_icon = (
        r"C:\Windows\Installer\{C1593F76-F694-448E-AD35-82DDD6203975}\PowerShellExe.ico"
    )

    def __init__(self, parent=None):
        super(LogWindow, self).__init__(parent)
        self.setModal(False)  # Make it non-modal
        self.setWindowTitle("Batch Print Log")
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint
        )
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        self.setStyleSheet(f"background-color: {self.WINDOW_BACKGROUND_COLOR};")
        self.setMinimumSize(500, 200)
        self.setGeometry(100, 100, 725, 200)
        self.offset = None

        self.max_icon = QIcon(self.max_icon)
        self._build_ui()
        self.restoreWindowSettings()
        self.show()

    def _set_icon(self):
        """Sets the icon for the window"""
        # Get 3ds Max icon
        console_icon = QIcon(self.console_icon)

        max_icon_pixmap = self.max_icon.pixmap(24, 24)
        console_icon_pixmap = console_icon.pixmap(12, 12)

        # Create a new pixmap with the same size as self.max_icon
        new_pixmap = QPixmap(max_icon_pixmap.width(), max_icon_pixmap.height())
        new_pixmap.fill(Qt.GlobalColor.transparent)

        # Paint the self.max_icon on the new pixmap
        painter = QPainter(new_pixmap)
        painter.drawPixmap(0, 0, max_icon_pixmap)
        painter.drawPixmap(
            max_icon_pixmap.width() - console_icon_pixmap.width(),
            max_icon_pixmap.height() - console_icon_pixmap.height(),
            console_icon_pixmap,
        )
        painter.end()

        # Set the new pixmap as the window icon
        self.setWindowIcon(QIcon(new_pixmap))
        self.show()

    def _build_ui(self):
        self._set_icon()
        self._setup_title_bar_elements()

        # Log
        self.log = QPlainTextEdit(self)
        self.log.setReadOnly(True)
        self.log.setWordWrapMode(QTextOption.WrapMode.NoWrap)
        self.log.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.log.setStyleSheet("background-color: #5a5a5a; color: White;")

        self._create_resize_grips()

        # Layout - Log and title bar
        self.content_layout = QVBoxLayout()
        self.content_layout.addWidget(self.log)

        window_layout = QGridLayout()
        window_layout.setContentsMargins(0, 0, 0, 0)
        window_layout.setSpacing(0)
        window_layout.addLayout(self._setup_window_title(), 0, 0, 3, 3)
        window_layout.setRowStretch(1, 1)
        window_layout.setColumnStretch(1, 1)
        window_layout.setColumnMinimumWidth(0, self.GRIP_SIZE)
        window_layout.setColumnMinimumWidth(2, self.GRIP_SIZE)
        window_layout.setRowMinimumHeight(0, self.closeButton.height() + 1)
        window_layout.setRowMinimumHeight(2, self.GRIP_SIZE)
        window_layout.addLayout(self.content_layout, 1, 1)

        self._setup_resize_events()

        # Create a central widget and set the layout
        central_widget = QWidget()
        central_widget.setLayout(window_layout)
        self.setLayout(window_layout)  # Set the layout on the dialog itself

        # Make sure the window has a proper background
        central_widget.setAutoFillBackground(True)
        palette = central_widget.palette()
        palette.setColor(central_widget.backgroundRole(), Qt.GlobalColor.transparent)
        central_widget.setPalette(palette)

        # Make sure the window is properly sized
        self.adjustSize()

    def _setup_title_bar_elements(self):
        """Builds the UI"""
        # Title bar icon
        icon_size = 16  # same W & H
        self.icon_label = QLabel(self)
        self.icon_label.setPixmap(self.max_icon.pixmap(icon_size, icon_size))
        self.icon_label.setFixedSize(icon_size, icon_size)
        self.icon_label.mousePressEvent = self.titlebarMousePressEvent
        self.icon_label.mouseMoveEvent = self.titlebarMouseMoveEvent

        # Title bar title
        self.title = QLabel(self)
        self.title.setText("  " + self.windowTitle())
        self.title.setStyleSheet(f"color: {self.TITLE_BAR_COLOR}")
        self.title.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.title.mousePressEvent = self.titlebarMousePressEvent
        self.title.mouseMoveEvent = self.titlebarMouseMoveEvent

        # X Close button
        self.closeButton_size = 29
        self.closeButton = QPushButton("\U00010317", self)
        self.closeButton.setFlat(True)
        self.closeButton.clicked.connect(self.hide)
        self.closeButton.setFixedSize(self.closeButton_size, self.closeButton_size)
        self.closeButton.setStyleSheet(
            f"background-color: {self.WINDOW_BACKGROUND_COLOR}; "
            f"color: {self.TITLE_BAR_COLOR}; "
            f"font: 12pt 'Segoe UI Historic'; text-align: center;"
        )

    def _create_resize_grips(self):
        # https://stackoverflow.com/questions/62807295/how-to-resize-a-window-from-the-edges-after-adding-the-property-qtcore-qt-framel
        # ^Not really relevant anymore, but had some inspiration^
        # Grips
        self.grip_top = QWidget()
        self.grip_top.setFixedHeight(self.GRIP_SIZE)
        self.grip_top.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        if self.DEBUG_COLORS:
            self.grip_top.setStyleSheet(f"background-color: green;")

        self.grip_left = QWidget()
        self.grip_left.setFixedWidth(self.GRIP_SIZE)
        if self.DEBUG_COLORS:
            self.grip_left.setStyleSheet(f"background-color: yellow;")

        self.grip_right = QWidget()
        self.grip_right.setFixedWidth(self.GRIP_SIZE)
        if self.DEBUG_COLORS:
            self.grip_right.setStyleSheet(f"background-color: white;")

        self.grip_bottom = QWidget()
        self.grip_bottom.setFixedHeight(self.GRIP_SIZE)
        if self.DEBUG_COLORS:
            self.grip_bottom.setStyleSheet(f"background-color: purple;")

        self.grip_topRight = QWidget()
        self.grip_topRight.setFixedSize(self.GRIP_SIZE, self.GRIP_SIZE)
        if self.DEBUG_COLORS:
            self.grip_topRight.setStyleSheet(f"background-color: orange;")

        self.grip_topLeft = QWidget()
        self.grip_topLeft.setFixedSize(self.GRIP_SIZE, self.GRIP_SIZE)
        if self.DEBUG_COLORS:
            self.grip_topLeft.setStyleSheet(f"background-color: blue;")

        self.grip_bottomRight = QWidget()
        self.grip_bottomRight.setFixedSize(self.GRIP_SIZE, self.GRIP_SIZE)
        if self.DEBUG_COLORS:
            self.grip_bottomRight.setStyleSheet(f"background-color: orange;")

        self.grip_bottomLeft = QWidget()
        self.grip_bottomLeft.setFixedSize(self.GRIP_SIZE, self.GRIP_SIZE)
        if self.DEBUG_COLORS:
            self.grip_bottomLeft.setStyleSheet(f"background-color: blue;")

    def _setup_resize_events(self):
        # Mouse hover events
        self.grip_left.enterEvent = lambda event: self.setCursor(Qt.CursorShape.SizeHorCursor)
        self.grip_top.enterEvent = lambda event: self.setCursor(Qt.CursorShape.SizeVerCursor)
        self.grip_right.enterEvent = lambda event: self.setCursor(Qt.CursorShape.SizeHorCursor)
        self.grip_bottom.enterEvent = lambda event: self.setCursor(
            Qt.CursorShape.SizeVerCursor
        )
        self.grip_topLeft.enterEvent = lambda event: self.setCursor(
            Qt.CursorShape.SizeFDiagCursor
        )
        self.grip_topRight.enterEvent = lambda event: self.setCursor(
            Qt.CursorShape.SizeBDiagCursor
        )
        self.grip_bottomLeft.enterEvent = lambda event: self.setCursor(
            Qt.CursorShape.SizeBDiagCursor
        )
        self.grip_bottomRight.enterEvent = lambda event: self.setCursor(
            Qt.CursorShape.SizeFDiagCursor
        )
        self.title.enterEvent = lambda event: self.setCursor(Qt.CursorShape.ArrowCursor)
        self.closeButton.enterEvent = lambda event: self.setCursor(Qt.CursorShape.ArrowCursor)
        self.icon_label.enterEvent = lambda event: self.setCursor(Qt.CursorShape.ArrowCursor)

        self.grip_topLeft.mouseMoveEvent = lambda event: self.gripMoveEvent(
            event, self.grip_topLeft
        )
        self.grip_topRight.mouseMoveEvent = lambda event: self.gripMoveEvent(
            event, self.grip_topRight
        )
        self.grip_bottomLeft.mouseMoveEvent = lambda event: self.gripMoveEvent(
            event, self.grip_bottomLeft
        )
        self.grip_bottomRight.mouseMoveEvent = lambda event: self.gripMoveEvent(
            event, self.grip_bottomRight
        )
        self.grip_left.mouseMoveEvent = lambda event: self.gripMoveEvent(event, self.grip_left)
        self.grip_top.mouseMoveEvent = lambda event: self.gripMoveEvent(event, self.grip_top)
        self.grip_right.mouseMoveEvent = lambda event: self.gripMoveEvent(
            event, self.grip_right
        )
        self.grip_bottom.mouseMoveEvent = lambda event: self.gripMoveEvent(
            event, self.grip_bottom
        )

    def _setup_window_title(self) -> QGridLayout:
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
        windowTitle_And_ResizeGrips.addWidget(
            self.grip_left, 0, 0, 3, 0, alignment=Qt.AlignmentFlag.AlignLeft
        )
        windowTitle_And_ResizeGrips.addWidget(
            self.grip_right, 0, 3, 3, 3, alignment=Qt.AlignmentFlag.AlignRight
        )
        windowTitle_And_ResizeGrips.addWidget(
            self.grip_top, 0, 1, 0, 3, alignment=Qt.AlignmentFlag.AlignTop
        )
        windowTitle_And_ResizeGrips.addWidget(
            self.grip_topLeft,
            0,
            0,
            alignment=Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft,
        )
        windowTitle_And_ResizeGrips.addWidget(
            self.grip_topRight,
            0,
            3,
            alignment=Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight,
        )
        windowTitle_And_ResizeGrips.addWidget(
            self.grip_bottom, 4, 1, 4, 3, alignment=Qt.AlignmentFlag.AlignTop
        )
        windowTitle_And_ResizeGrips.addWidget(
            self.grip_bottomLeft,
            4,
            0,
            alignment=Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignLeft,
        )
        windowTitle_And_ResizeGrips.addWidget(
            self.grip_bottomRight,
            4,
            3,
            alignment=Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignRight,
        )
        windowTitle_And_ResizeGrips.addWidget(
            self.closeButton,
            0,
            3,
            alignment=Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignLeft,
        )
        return windowTitle_And_ResizeGrips

    def restoreWindowSettings(self):
        """Restores the window settings."""
        # Restore window position & size
        settings = QSettings(AppInfo.PUBLISHER, AppInfo.APP_NAME)
        # noinspection PyTypeChecker
        self.restoreGeometry(settings.value("geometry"))

    def saveWindowSettings(self):
        """Saves the window settings."""
        # Save publisher creator
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, f"Software\\{AppInfo.PUBLISHER}", 0, winreg.KEY_WRITE
        )
        winreg.SetValueEx(key, "creator", 0, winreg.REG_SZ, AppInfo.PUBLISHER_CREATOR)
        winreg.CloseKey(key)

        # Save window position & size
        settings = QSettings(AppInfo.PUBLISHER, AppInfo.APP_NAME)
        settings.setValue("geometry", self.saveGeometry())
        settings.setValue("author", AppInfo.AUTHOR)
        settings.setValue("version", AppInfo.APP_VERSION)

    def gripMoveEvent(self, event, grip) -> None:
        """Handles the grip move event.
        Args:
            event: Event
            grip: Grip
        """
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
        """Handles the titlebar mouse press event.
        Args:
            event: Event
        """
        self.offset = -(self.pos() - event.globalPosition().toPoint())

    def titlebarMouseMoveEvent(self, event):
        """Handles the titlebar mouse move event.
        Args:
            event: Event
        """
        delta = event.globalPosition() - self.offset
        self.move(delta.toPoint())

    def log_message(self, message: str):
        """Handles the message from the other application.
        Args:
            message: Message
        """

        print(message)
        label_text = self.log.toPlainText()
        label_text += message + "\n"

        self.log.setPlainText(label_text)
        self.log.moveCursor(QTextCursor.MoveOperation.End)

    def exit(self):
        """Exits the application"""
        self.saveWindowSettings()
        sys.exit()


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = LogWindow()
    window.show()
    sys.exit(app.exec())
