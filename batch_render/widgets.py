from functools import partial

import PySide6
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QPushButton, QVBoxLayout

qt_widgets = PySide6.QtWidgets


class GenericDialog(qt_widgets.QDialog):
    def __init__(self, title, message, buttons, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        self.setFixedWidth(250)

        # Message label
        message_label = qt_widgets.QLabel(message)
        message_label.setWordWrap(True)
        message_label.setContentsMargins(10, 5, 10, 10)

        # Buttons
        button_layout = qt_widgets.QHBoxLayout()
        for button_text, button_result in buttons:
            button = QPushButton(button_text)
            button.setFixedWidth(75)
            button.clicked.connect(partial(self.accept_with_result, button_result))
            button_layout.addWidget(button)

        # Main layout
        main_layout = QVBoxLayout()
        main_layout.addWidget(message_label)
        main_layout.addLayout(button_layout)

        self.setLayout(main_layout)

        if parent:
            self.center_on_parent(parent)

    def center_on_parent(self, parent):
        parent_geometry = parent.geometry()
        dialog_geometry = self.geometry()

        center_x = (
            parent_geometry.x() + (parent_geometry.width() - dialog_geometry.width()) // 2
        )
        center_y = (
            parent_geometry.y() + (parent_geometry.height() - dialog_geometry.height()) // 2
        )

        self.setGeometry(center_x, center_y, dialog_geometry.width(), dialog_geometry.height())

    def accept_with_result(self, result):
        self.done(result)


class TruncateDelegateRight(qt_widgets.QStyledItemDelegate):
    def paint(self, painter, option, index):
        # Get the text and rectangle
        text = index.data(Qt.ItemDataRole.DisplayRole)
        rect = option.rect

        # Truncate the text and adjustment ellipsis
        elided_text = option.fontMetrics.elidedText(
            text, Qt.TextElideMode.ElideRight, rect.adjusted(0, 0, -5, 0).width()
        )

        # Draw the truncated text
        painter.drawText(
            rect.adjusted(5, 0, 0, 0),
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
            elided_text,
        )


class TruncateDelegateMiddle(qt_widgets.QStyledItemDelegate):
    def paint(self, painter, option, index):
        # Get the text and rectangle
        text = index.data(Qt.ItemDataRole.DisplayRole)
        rect = option.rect

        # Truncate the text and adjustment ellipsis
        elided_text = option.fontMetrics.elidedText(
            text, Qt.TextElideMode.ElideMiddle, rect.adjusted(0, 0, -5, 0).width()
        )

        # Draw the truncated text
        painter.drawText(
            rect.adjusted(5, 0, 0, 0),
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
            elided_text,
        )


class CustomTableWidgetItem(qt_widgets.QTableWidgetItem):
    def __init__(self, display_value, hidden_value):
        super().__init__(display_value)
        self.hidden_value = hidden_value
