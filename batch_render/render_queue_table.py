"""Render queue table."""

import re

from PySide6 import QtWidgets as QtWidgets
from PySide6.QtCore import Qt

from .config import Config
from .widgets import CustomTableWidgetItem, TruncateDelegateMiddle, TruncateDelegateRight
from log_window import console, VerboseLevel


class Column:
    """Column element."""

    def __init__(
        self,
        column_name: str,
        editable: bool = False,
        resizable: bool = True,
        additional_tags=None,
    ):
        if additional_tags is None:
            additional_tags = []
        else:
            self.check_tags(additional_tags)

        self.name = column_name
        self.editable = editable
        self.resizable = resizable
        self.additional_tags = additional_tags

    @staticmethod
    def check_tags(tags) -> None:
        """Check if tags contain spaces.

        Args:
            tags (list): List of tags.
        """
        for tag in tags:
            if " " in tag:
                raise ValueError("Tags cannot contain spaces")

    @property
    def tags(self):
        """Get replacement tags for file naming."""
        column_name = self.name.replace(" ", "_")

        tags = []
        for tag in self.additional_tags + [column_name]:
            tags.append("{" + tag.lower() + "}")

        return tags


class ColumnMeta(type):
    """Column meta."""

    def __new__(mcs, name, bases, attrs):
        cols = []
        for key, value in attrs.items():
            if isinstance(value, Column):
                value.attr_name = key
                cols.append(value)

        cls = super().__new__(mcs, name, bases, attrs)
        cls._all_columns = cols
        return cls


class Columns(metaclass=ColumnMeta):
    """Table columns (ORDER MATTERS)"""

    USE = Column("Use", False, True, [])
    NAME = Column("Name", True, True, [])
    CAMERA = Column("Camera", False, True, [])
    OUTPUT_PATH = Column("Output Path", True, True, [])
    RANGE = Column("Range", False, True, [])
    RESOLUTION = Column("Resolution", False, True, [])
    PIXEL_ASPECT = Column("Pixel Aspect", False, True, [])
    SCENE_STATE = Column("Scene State", False, True, ["State_Set"])
    RENDER_PRESET = Column("Render Preset", False, True, [])
    LAYER_PRESET = Column("Layer Preset", False, True, [])

    def __iter__(self):
        return iter(self._all_columns)

    def __len__(self):
        return len(self._all_columns)

    def __getitem__(self, item):
        return self._all_columns[item]

    @property
    def names(self) -> list[str]:
        """Get column names."""
        return [column.name for column in self._all_columns]

    @classmethod
    def replace_tags(cls, string: str, properties: dict[str, str]) -> tuple[str, bool]:
        """Replace tags in string."""
        blank_values = False
        for column in cls._all_columns:
            for tag in column.tags:
                if tag in string.lower():
                    property_value = properties[column.name]
                    if property_value.lower() in [None, "default", ""]:
                        property_value = ""
                        blank_values = True
                    string = re.sub(tag, property_value, string, flags=re.IGNORECASE)

        return string, blank_values

    @classmethod
    def build_naming_tooltip(cls):
        """Build naming tooltip."""
        prefix = "Valid replacement flags (not case sensitive):\n"
        all_tags = []

        for column in cls._all_columns:
            all_tags.extend(column.tags)

        return prefix + ", ".join(all_tags)


class ColumnInfo:
    """Column information."""

    MINIMUM_WIDTH = 100


columns = Columns()


class RenderQueueTable(QtWidgets.QTableWidget):
    """Render queue table."""

    def __init__(self, parent=None):
        super().__init__()
        self.parent = parent

        self.system_modified = False
        self.previously_selected = None
        self._setup_columns()
        self._setup_style()

    def setCellData(self, row, column, display_value, hidden_value):
        """Set cell data."""
        item = CustomTableWidgetItem(display_value, hidden_value)
        self.setItem(row, column, item)

    def getHiddenValue(self, row, column):
        """Get hidden value from table item."""
        item = self.item(row, column)
        if isinstance(item, CustomTableWidgetItem):
            return item.hidden_value
        else:
            return None

    """Setup functions"""

    def _setup_columns(self):
        """Setup table columns."""
        self.setColumnCount(len(columns.names))
        self.setHorizontalHeaderLabels(columns.names)

        initial_row_height = self.rowHeight(0)
        self.verticalHeader().setDefaultSectionSize(initial_row_height)
        self.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeMode.Fixed)
        self.horizontalHeaderItem(1).setToolTip(Columns.build_naming_tooltip())
        self.setItemDelegate(TruncateDelegateRight(self))  # Truncate style
        self.setItemDelegateForColumn(3, TruncateDelegateMiddle(self))

        # Table - Checkbox column
        self.setItemDelegateForColumn(0, QtWidgets.QStyledItemDelegate())
        self.setColumnWidth(0, 1)
        header = self.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeMode.Fixed)

    def _setup_style(self):
        """Setup table style."""
        self.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior(1))
        self.setStyleSheet("QTableWidget::item:selected {background-color: #3498db;}")

    """Table functions"""

    def resize_column_to_contents(self, column: int | str = None):
        """Resize column to contents."""
        if type(column) is int:
            column_names = [column]
        elif type(column) is str:
            column_names = [columns.names.index(column)]
        elif column is None:
            column_names = columns.names
        else:
            raise ValueError

        for c, _ in enumerate(column_names):
            self.resizeColumnToContents(c)
            column_width = self.columnWidth(c)
            self.setColumnWidth(c, column_width + 4)
            if column_width < ColumnInfo.MINIMUM_WIDTH:
                self.setColumnWidth(c, ColumnInfo.MINIMUM_WIDTH)

    def get_entry_values(self, row: int):
        """Returns a dictionary of entry values for a given row"""
        item_values = {}
        for column_index, column in enumerate(columns):
            column_name = column.name
            column_name = column_name.replace(" ", "_")
            item = self.item(row, column_index)
            widget = self.cellWidget(row, column_index)

            if widget:
                # noinspection PyUnresolvedReferences
                value = widget.isChecked()
                hidden_value = None
            else:
                value = item.text()
                hidden_value = self.getHiddenValue(row, column_index)

            # rowProperties[column_name] = [value, hidden_value]
            item_values[column_name] = value
            item_values[column_name + "_Hidden"] = hidden_value

        return item_values

    def table_get_selected(self, item_index: int | None = 0):
        selected_items = self.selectedItems()

        if selected_items:
            selected_items_int = [x.row() for x in selected_items]
            combined_lists = list(zip(selected_items_int, selected_items))
            sorted_lists = sorted(combined_lists, key=lambda x: x[0])

            console.log(VerboseLevel.DEBUG, "selection: " + str(selected_items))

            if item_index is not None:
                selected_items = sorted_lists[item_index]
            else:
                selected_items = sorted_lists
        else:
            selected_items = (None, None)

        return selected_items

    """Table actions"""

    def add_row(self, suppress_output: bool = False):
        self.system_modified = True

        row_position = self.rowCount()
        self.previously_selected = row_position
        self.insertRow(row_position)
        # self.item(row_position, 1).setToolTip(self.naming_tooltip)

        # Add checkbox to the first column
        checkbox = QtWidgets.QCheckBox()
        checkbox.setChecked(True)
        self.setCellWidget(row_position, 0, checkbox)

        # You can populate the cells with default values if needed
        for col in range(1, self.columnCount() - 2):
            item = QtWidgets.QTableWidgetItem("Default")
            if col == columns.names.index("Name"):
                item.setToolTip(columns.build_naming_tooltip())
            self.setItem(row_position, col, item)

        # Default - Camera
        default_camera = self.parent.cmbCamera.itemText(0)
        default_camera_ID = self.parent.cmbCamera.itemData(0)
        default = QtWidgets.QTableWidgetItem(default_camera)
        col = columns.names.index("Camera")
        self.setCellData(row_position, col, default, default_camera_ID)

        # Default - State, Render Preset, & Layer Preset
        # State
        default = QtWidgets.QTableWidgetItem("")
        col = columns.names.index("Scene State")
        self.setItem(row_position, col, default)
        # Render Preset
        default = QtWidgets.QTableWidgetItem("")
        col = columns.names.index("Render Preset")
        self.setItem(row_position, col, default)
        # Layer Preset
        default = QtWidgets.QTableWidgetItem("")
        col = columns.names.index("Layer Preset")
        self.setItem(row_position, col, default)

        # Default - Output Path
        default = QtWidgets.QTableWidgetItem(Config.DEFAULT_PATH_TEXT)
        col = columns.names.index("Output Path")
        self.setItem(row_position, col, default)

        # Check if the current column is the one you want to make non-editable
        for col, column in enumerate(columns):
            item = self.item(row_position, col)
            if item is not None and column.editable:
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)

        self.resize_column_to_contents()
        self.system_modified = False
        self.selectRow(row_position)
        if not suppress_output:
            console.log(VerboseLevel.INFO, "New row added")

    def duplicate_row(self):
        self.system_modified = True

        selected_rows = self.selectedItems()
        if selected_rows:
            rows_to_duplicate = sorted(list(set(x.row() for x in selected_rows)))
            x = 0
            for row_int in reversed(rows_to_duplicate):
                # Insert a new row below the selected row
                newRow_int = rows_to_duplicate[-1] + 1 + x
                self.insertRow(newRow_int)

                # Duplicate the items from the selected row to the new row
                for col in range(self.columnCount()):
                    column_name = self.horizontalHeaderItem(col).text()
                    item = self.item(row_int, col)
                    item_cell = self.cellWidget(row_int, col)

                    hidden_value = None
                    if isinstance(item_cell, QtWidgets.QCheckBox):
                        value_type = "bool"
                        value = item_cell.isChecked()
                        new_item = QtWidgets.QCheckBox()
                        new_item.setChecked(value)
                        self.setCellWidget(newRow_int, col, new_item)
                    else:
                        new_item = item.clone()
                        value = item.text()
                        hidden_value = self.getHiddenValue(row_int, col)
                        value_type = "string"
                        new_item.setText(value)
                        self.setCellData(newRow_int, col, value, hidden_value)

                    with console.indent():
                        console.log(
                            VerboseLevel.INFO,
                            f"{column_name}: {value} [{value_type}] ({hidden_value})",
                        )
                    if column_name == "Camera" and hidden_value is None or "":
                        raise ValueError("Camera ID is missing")

            # Deselect old cells
            start_row, start_col = rows_to_duplicate[0], 0
            end_row, end_col = rows_to_duplicate[-1], self.columnCount() - 1
            rows_to_deselect = QtWidgets.QTableWidgetSelectionRange(
                start_row, start_col, end_row, end_col
            )
            self.setRangeSelected(rows_to_deselect, False)

            # Select new duplicated cells
            rows_to_select = [x + len(rows_to_duplicate) for x in rows_to_duplicate]
            start_row, start_col = rows_to_select[0], 0
            end_row, end_col = rows_to_select[-1], self.columnCount() - 1
            rows_to_select = QtWidgets.QTableWidgetSelectionRange(
                start_row, start_col, end_row, end_col
            )
            self.setRangeSelected(rows_to_select, True)

            duplicated_row_nums = [x + 1 for x in rows_to_duplicate]
            console.log(VerboseLevel.INFO, "Duplicated: Rows " + str(duplicated_row_nums))

        self.system_modified = False

    def delete_row(self):
        selected_items = self.selectedItems()

        if selected_items:
            self.system_modified = True
            rows_to_delete = sorted(list(set(x.row() for x in selected_items)), reverse=True)
            rows_deleted = []
            x = 0
            for x in rows_to_delete:
                self.removeRow(x)
                rows_deleted.append(x + 1)

            console.log(VerboseLevel.INFO, "Deleted: Rows " + str(rows_deleted))

            self.previously_selected = x - 1
            if self.previously_selected < 0:
                self.previously_selected = 0
            self.system_modified = False
            self.selectRow(self.previously_selected)

    def move_up(self):
        selected_items = self.selectedItems()
        if selected_items:
            rows_to_move = sorted(list(set(x.row() for x in selected_items)))
            one_above_selection = rows_to_move[0] - 1
            x = 0
            if one_above_selection >= 0:
                for row_id in rows_to_move:
                    self.move_row(row_id, one_above_selection + x)
                    x += 1

                # Select moved sells
                rows_to_select = [x - 1 for x in rows_to_move]
                start_row, start_col = rows_to_select[0], 0
                end_row, end_col = rows_to_select[-1], self.columnCount() - 1
                rows_to_select = QtWidgets.QTableWidgetSelectionRange(
                    start_row, start_col, end_row, end_col
                )
                self.setRangeSelected(rows_to_select, True)

    def move_down(self):
        selected_items = self.selectedItems()
        if selected_items:
            rows_to_move = sorted(list(set(x.row() for x in selected_items)))
            one_below_selection = rows_to_move[-1] + 1
            x = 0
            if one_below_selection < self.rowCount():
                for row_id in reversed(rows_to_move):
                    self.move_row(row_id, one_below_selection - x)
                    x += 1

                rows_to_select = [x + 1 for x in rows_to_move]
                start_row, start_col = rows_to_select[0], 0
                end_row, end_col = rows_to_select[-1], self.columnCount() - 1
                rows_to_select = QtWidgets.QTableWidgetSelectionRange(
                    start_row, start_col, end_row, end_col
                )
                self.setRangeSelected(rows_to_select, True)

    def move_row(self, from_row: int, to_row: int):
        self.system_modified = True

        # Copy items from the source row to a dictionary
        items_dict = {}
        for col, column_name in enumerate(columns.names):
            column_name = column_name.replace(" ", "_")
            item = self.item(from_row, col)
            item_cell = self.cellWidget(from_row, col)

            if item_cell:
                # noinspection PyUnresolvedReferences
                value = item_cell.isChecked()
                hidden_value = None
            else:
                value = item.text()
                hidden_value = self.getHiddenValue(from_row, col)

            items_dict[column_name] = value
            items_dict[column_name + "_Hidden"] = hidden_value

        # Remove the source row
        self.removeRow(from_row)

        # Insert a new row at the destination
        self.insertRow(to_row)

        # Populate the destination row with the copied items
        for col, column_name in enumerate(columns.names):
            column_name = column_name.replace(" ", "_")
            value = items_dict[column_name.replace(" ", "_")]
            if isinstance(value, bool):  # Checkbox
                new_item_cell = QtWidgets.QCheckBox()
                new_item_cell.setChecked(value)
                self.setCellWidget(to_row, col, new_item_cell)
            else:  # Text or some other value
                new_item = QtWidgets.QTableWidgetItem(str(value))
                self.setCellData(to_row, col, new_item, items_dict[column_name + "_Hidden"])

        # Select the new row
        self.selectRow(to_row)
        self.system_modified = False

    """Table selection"""

    def selection_changed(self):
        console.log(VerboseLevel.DEBUG, "table_selection_changed")

        if not self.system_modified:
            selected_items = self.table_get_selected()[0]

            # Log
            display_selected = selected_items
            if selected_items:
                display_selected = selected_items + 1
            console.log(VerboseLevel.DEBUG, "Selected: Row " + str(display_selected))

            if selected_items is not None:
                self.previously_selected = selected_items
                console.log(
                    VerboseLevel.DEBUG,
                    "Selected: Row " + str(self.previously_selected + 1),
                )

            elif self.previously_selected is not None:
                console.log(
                    VerboseLevel.DEBUG,
                    "No element selected, but previous element selected",
                )
                self.system_modified = True
                self.selectRow(self.previously_selected)

            else:
                console.log(
                    VerboseLevel.DEBUG,
                    "No element selected, and no previous element selected",
                )
                self.previously_selected = self.rowCount()
                self.system_modified = True
                self.selectRow(self.previously_selected)

            self.system_modified = False
            self.parent.update_element_values()


if __name__ == "__main__":
    pass
