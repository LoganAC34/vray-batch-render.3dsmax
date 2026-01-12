"""
This script is an extended functionality version of the standard 3ds Max batch render dialog.

It adds:
- Multi-row editing such as renaming and deleting
- The ability to move rows up/down (render order).
- V-Ray layer presets to be applied before rendering.
- Better output path control (base output path + row name + frame #).
- Override toggles for each frame range, resolution, pixel aspect, output path.
- Macro to adjustment unlisted cameras.

Feature wishlist:
- Render start/end scripts
- Change frame range override value to string similar to standard render window frame range
- Render pre-check to check all render entries instead of stopping at first error.
- Render summary of successful/errors vs total renders.
- Be able to toggle multiple render entries at once
"""

from .config import Config
from .utils import *
from .widgets import GenericDialog
from .render_queue_table import RenderQueueTable, Columns

import os

# https://forums.autodesk.com/t5/3ds-max-programming/installing-python-modules/td-p/6522899
import sys

script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.append(script_dir)

import winreg
from pathlib import Path
from typing import Tuple, Any

from log_window import VerboseLevel, Console

import json
import re
import tempfile
import time
from datetime import datetime

# noinspection PyUnresolvedReferences
import pymxs

from PySide6.QtWidgets import QPushButton, QVBoxLayout, QWidget
from PySide6.QtCore import *
from PySide6 import QtWidgets as QtWidgets


# noinspection PyUnresolvedReferences
rt = pymxs.runtime
columns = Columns()

MAX_VERSION_YEAR = rt.maxVersion()[-2]
RENDER_OUTPUT_PATH = Path(rt.GetDir(rt.Name("renderoutput")))
PROJECT_ROOT_PATH = rt.pathConfig.getCurrentProjectFolder()


class BatchRenderDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super(BatchRenderDialog, self).__init__(parent)
        self.stateSet_prefix = "State Set: "
        self.sceneState_prefix = "Scene State: "
        self.script_path = os.path.dirname(__file__)

        # App settings
        self.setWindowTitle(f"{Config.APP_NAME} {Config.APP_VERSION}")
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        self.setGeometry(100, 100, 1000, 500)
        self.restoreWindowSettings()

        # Log verbose settings
        self.DEBUG_disable_rendering = False

        # Initialize variable default
        self.system_modified = False

        # Build the GUI
        self._build_gui()
        console.open()

    """Dialog Functions"""

    # noinspection PyAttributeOutsideInit
    def _build_gui(self):
        self.dropdown_width = 350

        # Buttons
        self.btnAdd = QPushButton("Add", self)
        self.btnDuplicate = QPushButton("Duplicate", self)
        self.btnDelete = QPushButton("Delete", self)
        self.btnUp = QPushButton("\u25b2", self)
        self.btnDown = QPushButton("\u25bc", self)
        self.btnAddUnlistedCameras = QPushButton("+ unlisted", self)
        self.btnAddUnlistedCameras.setToolTip("Add unlisted Cameras")
        self.btnAddCameraSceneCombos = QPushButton("+ scene combos", self)
        self.btnAddCameraSceneCombos.setToolTip("Add camera + scene state/set combos")

        # Create a widget to hold the buttons
        self.buttonWidget = QWidget(self)
        # self.buttonWidget.setFixedWidth(400)

        # Layout for the buttons
        buttonLayout = QtWidgets.QHBoxLayout(self.buttonWidget)
        buttonLayout.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        buttonLayout.addWidget(self.btnAdd)
        buttonLayout.addWidget(self.btnDuplicate)
        buttonLayout.addWidget(self.btnDelete)
        buttonLayout.addWidget(self.btnUp)
        buttonLayout.addWidget(self.btnDown)
        buttonLayout.addStretch()
        buttonLayout.addWidget(self.btnAddUnlistedCameras)
        buttonLayout.addWidget(self.btnAddCameraSceneCombos)

        # Default Output Path field
        self.lblDefaultOutputPath = QtWidgets.QLabel("Default Output Path:", self)
        self.lblDefaultOutputPath.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        self.txtDefaultOutputPath = QtWidgets.QLineEdit(self)
        self.renderOutput = str(RENDER_OUTPUT_PATH.relative_to(PROJECT_ROOT_PATH))
        self.txtDefaultOutputPath.setText(self.renderOutput)

        # Browse button
        self.btnDefaultBrowse = QPushButton("...", self)
        self.btnDefaultBrowse.setFixedWidth(30)

        # Clear button
        self.btnDefaultClear = QPushButton("X", self)
        self.btnDefaultClear.setFixedWidth(30)

        # Layout for Default Output Path
        self.defaultOutputPathWidget = QWidget(self)
        defaultOutputPathLayout = QtWidgets.QHBoxLayout(self.defaultOutputPathWidget)
        defaultOutputPathLayout.addWidget(self.lblDefaultOutputPath)
        defaultOutputPathLayout.addWidget(self.btnDefaultBrowse)
        defaultOutputPathLayout.addWidget(self.txtDefaultOutputPath)
        defaultOutputPathLayout.addWidget(self.btnDefaultClear)

        # Table
        self.tableWidget = RenderQueueTable()  # QTableWidget(self)

        # Selected Parameters GroupBox
        self.groupBoxSelectedParams = QtWidgets.QGroupBox(
            "Selected Batch Render Parameters", self
        )

        # Override checkboxes
        self.frameRangeOverride = QtWidgets.QCheckBox("", self.groupBoxSelectedParams)
        self.imageSizeOverride = QtWidgets.QCheckBox("", self.groupBoxSelectedParams)
        self.pixelAspectOverride = QtWidgets.QCheckBox("", self.groupBoxSelectedParams)
        self.outputPathOverride = QtWidgets.QCheckBox("", self.groupBoxSelectedParams)

        # Override Frame Start
        self.lblFrameStart = QtWidgets.QLabel("Frame Start:", self.groupBoxSelectedParams)
        self.lblFrameStart.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        self.spnFrameStart = QtWidgets.QSpinBox(self.groupBoxSelectedParams)
        self.spnFrameStart.setFixedWidth(100)
        self.spnFrameStart.setMinimum(-9999)

        # Override Frame End
        self.lblFrameEnd = QtWidgets.QLabel("Frame End:", self.groupBoxSelectedParams)
        self.lblFrameEnd.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        self.spnFrameEnd = QtWidgets.QSpinBox(self.groupBoxSelectedParams)
        self.spnFrameEnd.setFixedWidth(100)
        self.spnFrameEnd.setMaximum(9999)

        # Override Image Width
        self.lblWidth = QtWidgets.QLabel("Width:", self.groupBoxSelectedParams)
        self.lblWidth.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.spnWidth = QtWidgets.QSpinBox(self.groupBoxSelectedParams)
        self.spnWidth.setFixedWidth(100)
        self.spnWidth.setMaximum(9999)
        self.spnWidth.setValue(rt.renderHeight)

        # Override Image Height
        self.lblHeight = QtWidgets.QLabel("Height:", self.groupBoxSelectedParams)
        self.lblHeight.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        self.spnHeight = QtWidgets.QSpinBox(self.groupBoxSelectedParams)
        self.spnHeight.setFixedWidth(100)
        self.spnHeight.setMaximum(9999)
        self.spnHeight.setValue(rt.renderWidth)

        # Override Pixel Aspect
        self.lblPixelAspect = QtWidgets.QLabel("Pixel Aspect:", self.groupBoxSelectedParams)
        self.lblPixelAspect.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        self.spnPixelAspect = QtWidgets.QDoubleSpinBox(self.groupBoxSelectedParams)
        self.spnPixelAspect.setFixedWidth(100)
        self.spnPixelAspect.setMaximum(2)
        self.spnPixelAspect.setSingleStep(0.1)
        self.spnPixelAspect.setValue(1.0)

        # Name field
        self.lblName = QtWidgets.QLabel("Name:", self.groupBoxSelectedParams)
        self.lblName.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.txtName = QtWidgets.QLineEdit(self.groupBoxSelectedParams)
        self.txtName.setFixedWidth(self.dropdown_width)
        self.txtName.setToolTip(columns.build_naming_tooltip())

        # Output Path field
        self.lblOutputPath = QtWidgets.QLabel("Output Path:", self.groupBoxSelectedParams)
        self.lblOutputPath.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        self.txtOutputPath = QtWidgets.QLineEdit(self.groupBoxSelectedParams)

        # Browse button
        self.btnBrowse = QPushButton("...", self.groupBoxSelectedParams)
        self.btnBrowse.setFixedWidth(30)

        # Clear button
        self.btnClear = QPushButton("X", self.groupBoxSelectedParams)
        self.btnClear.setFixedWidth(30)

        # Camera Dropdown
        self.lblCamera = QtWidgets.QLabel("Camera:", self.groupBoxSelectedParams)
        self.lblCamera.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        self.cmbCamera = QtWidgets.QComboBox(self.groupBoxSelectedParams)
        self.cmbCamera.setFixedWidth(self.dropdown_width)

        cameras = []
        camera_unique_ids = []  # Just for ensuring unique IDs
        for camera in rt.cameras:
            if hasattr(camera, "type"):
                # unique_id is this script's custom ID for each camera
                # internal_id is 3ds Max's internal ID (not reliable)
                try:
                    node_unique_id = get_item_unique_id(camera)
                except ValueError:
                    node_unique_id = None

                # Make sure there are no duplicate unique IDs (camera copied)
                while True:
                    if node_unique_id and node_unique_id not in camera_unique_ids:
                        camera_unique_ids.append(node_unique_id)
                        break
                    node_unique_id = generate_unique_id(camera)
                cameras.append((camera.name, node_unique_id))

        # Sort cameras by name then ID
        cameras.sort(key=lambda c: (c[0], c[1]))
        for camera_name, camera in cameras:
            self.cmbCamera.addItem(camera_name, camera)
        # log(cameras)

        # Scene State Dropdown
        self.lblSceneState = QtWidgets.QLabel(
            "Scene State & Sets:", self.groupBoxSelectedParams
        )
        self.lblSceneState.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        self.cmbSceneState = QtWidgets.QComboBox(self.groupBoxSelectedParams)
        self.cmbSceneState.setFixedWidth(self.dropdown_width)
        self.cmbSceneState.addItem(Config.DEFAULT_TEXT)
        for x in range(rt.sceneStateMgr.GetCount()):
            sceneState = rt.sceneStateMgr.GetSceneState(x + 1)
            self.cmbSceneState.addItem(f"{self.sceneState_prefix}{sceneState}")

        # Add State Sets
        # noinspection PyUnresolvedReferences
        stateSetsDotNetObject = pymxs.runtime.dotNetObject(
            "Autodesk.Max.StateSets.Plugin"
        ).Instance
        self.masterState = stateSetsDotNetObject.EntityManager.RootEntity.MasterStateSet
        for stateSet_name in self.get_state_sets(self.masterState):
            self.cmbSceneState.addItem(f"{self.stateSet_prefix}{stateSet_name}")

        # Render Preset Dropdown
        self.lblPreset = QtWidgets.QLabel("Render Preset:", self.groupBoxSelectedParams)
        self.lblPreset.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        self.cmbPreset = QtWidgets.QComboBox(self.groupBoxSelectedParams)
        self.cmbPreset.setFixedWidth(self.dropdown_width)
        preset_dir = os.listdir(rt.GetDir(rt.Name("renderPresets")))
        self.cmbPreset.addItem(Config.DEFAULT_TEXT)
        for renderPreset in preset_dir:
            self.cmbPreset.addItem(renderPreset)

        # Layer Preset Dropdown
        self.lblLayerPreset = QtWidgets.QLabel("Layer Preset:", self.groupBoxSelectedParams)
        self.lblLayerPreset.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        self.cmbLayerPreset = QtWidgets.QComboBox(self.groupBoxSelectedParams)
        self.cmbLayerPreset.setFixedWidth(self.dropdown_width)
        layerPreset_dir = os.listdir(rt.GetDir(rt.Name("vpost")))
        self.cmbLayerPreset.addItem(Config.DEFAULT_TEXT)
        for layerPreset in layerPreset_dir:
            self.cmbLayerPreset.addItem(layerPreset)

        # Layout for the group box (part 1)
        groupBoxLayout1 = QtWidgets.QGridLayout()
        groupBoxLayout1.addWidget(self.frameRangeOverride, 1, 0)
        groupBoxLayout1.addWidget(self.lblFrameStart, 1, 1)
        groupBoxLayout1.addWidget(self.spnFrameStart, 1, 2)
        groupBoxLayout1.addWidget(self.lblFrameEnd, 1, 3)
        groupBoxLayout1.addWidget(self.spnFrameEnd, 1, 4)
        # Add resize spacer to keep elements tot the left
        groupBoxLayout1.addItem(
            QtWidgets.QSpacerItem(1, 1, QtWidgets.QSizePolicy.Policy.Expanding), 1, 4
        )
        groupBoxLayout1.addWidget(self.imageSizeOverride, 2, 0)
        groupBoxLayout1.addWidget(self.lblWidth, 2, 1)
        groupBoxLayout1.addWidget(self.spnWidth, 2, 2)
        groupBoxLayout1.addWidget(self.lblHeight, 2, 3)
        groupBoxLayout1.addWidget(self.spnHeight, 2, 4)

        groupBoxLayout1.addWidget(self.pixelAspectOverride, 3, 0)
        groupBoxLayout1.addWidget(self.lblPixelAspect, 3, 3)
        groupBoxLayout1.addWidget(self.spnPixelAspect, 3, 4)

        # Layout for the group box (part 2)
        groupBoxLayout2 = QtWidgets.QGridLayout()
        groupBoxLayout2.addWidget(self.outputPathOverride, 1, 0)
        groupBoxLayout2.addWidget(self.lblOutputPath, 1, 1)
        groupBoxLayout2.addWidget(self.btnBrowse, 1, 2)
        groupBoxLayout2.addWidget(self.txtOutputPath, 1, 3)
        groupBoxLayout2.addWidget(self.btnClear, 1, 4)

        # Layout for the group box (part 3)
        groupBoxLayout3 = QtWidgets.QGridLayout()
        groupBoxLayout3.addWidget(self.lblName, 2, 0)
        groupBoxLayout3.addWidget(self.txtName, 2, 1, 1, 1)
        groupBoxLayout3.addItem(
            QtWidgets.QSpacerItem(1, 1, QtWidgets.QSizePolicy.Policy.Expanding), 2, 3
        )  # Add resize spacer to keep elements to the left
        groupBoxLayout3.addWidget(self.lblCamera, 3, 0)
        groupBoxLayout3.addWidget(self.cmbCamera, 3, 1, 1, 1)
        groupBoxLayout3.addWidget(self.lblSceneState, 4, 0)
        groupBoxLayout3.addWidget(self.cmbSceneState, 4, 1, 1, 1)
        groupBoxLayout3.addWidget(self.lblPreset, 5, 0)
        groupBoxLayout3.addWidget(self.cmbPreset, 5, 1, 1, 1)
        groupBoxLayout3.addWidget(self.lblLayerPreset, 6, 0)
        groupBoxLayout3.addWidget(self.cmbLayerPreset, 6, 1, 1, 1)

        # Combine groupBoxLayout1 and groupBoxLayout2 vertically
        vbox = QVBoxLayout(self.groupBoxSelectedParams)
        vbox.addLayout(groupBoxLayout1)
        vbox.addLayout(groupBoxLayout2)
        vbox.addLayout(groupBoxLayout3)

        # Export to batt and render Buttons
        self.btnBatt = QPushButton("Export to.bat", self)
        self.btnBatt.setFixedWidth(90)
        self.btnBatt.setVisible(
            False
        )  ### Disabled button since IDK how to implement the function right now.

        self.btnRender = QPushButton("Render", self)
        self.btnRender.setFixedWidth(90)

        self.btnLog = QPushButton("Open Log", self)
        self.btnLog.setFixedWidth(90)

        # Create a widget to hold the buttons
        self.renderWidget = QWidget(self)

        # Layout for the buttons
        renderLayout = QtWidgets.QHBoxLayout(self.renderWidget)
        renderLayout.addWidget(self.btnLog)
        renderLayout.addWidget(self.btnBatt)
        renderLayout.addWidget(self.btnRender)

        # Main layout
        mainLayout = QVBoxLayout(self)
        mainLayout.addWidget(self.defaultOutputPathWidget, 0)  # Align buttons to the top left
        mainLayout.addWidget(self.buttonWidget)
        mainLayout.addWidget(self.tableWidget, 1)
        mainLayout.addWidget(self.groupBoxSelectedParams, 0)
        mainLayout.addWidget(
            self.renderWidget,
            alignment=Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight,
        )  # Align buttons to the top left

        # Binds - action buttons
        self.btnDefaultBrowse.clicked.connect(self.browse_default_output_path)
        self.btnDefaultClear.clicked.connect(self.clear_default_output_path)
        self.btnAdd.clicked.connect(self.tableWidget.add_row)
        self.btnAdd.setFocusPolicy(Qt.FocusPolicy.NoFocus)  # Ignore ENTER and RETURN keys
        self.btnDuplicate.clicked.connect(self.tableWidget.duplicate_row)
        self.btnDelete.clicked.connect(self.tableWidget.delete_row)
        self.btnUp.clicked.connect(self.tableWidget.move_up)
        self.btnDown.clicked.connect(self.tableWidget.move_down)
        self.btnBrowse.clicked.connect(self.browse_output_path)
        self.btnClear.clicked.connect(self.clear_output_path)
        self.btnRender.clicked.connect(self.start_batch_render)  # batch_render
        self.btnLog.clicked.connect(console.open)
        self.btnAddUnlistedCameras.clicked.connect(self.add_unlisted_cameras)
        self.btnAddCameraSceneCombos.clicked.connect(self.add_camera_sceneStateSet_combos)

        # Binds - Windows Events
        self.finished.connect(self.on_close)

        # Disable overrides (default state)
        self.toggle_override_fields(False)

        # Restore dialog info
        try:
            self.restoreDialogData()
            self.update_element_values()
        except RuntimeError as e:
            console.log(VerboseLevel.ERROR, str(e))

        # Bind - table
        self.tableWidget.itemSelectionChanged.connect(self.tableWidget.selection_changed)
        self.tableWidget.itemChanged.connect(self.update_element_values)

        # Binds - parameters changed
        self.outputPathOverride.stateChanged.connect(self.output_path_override_toggled)
        self.frameRangeOverride.stateChanged.connect(self.frame_range_override_toggled)
        self.imageSizeOverride.stateChanged.connect(self.image_size_override_toggled)
        self.pixelAspectOverride.stateChanged.connect(self.pixel_aspect_override_toggled)
        self.spnFrameStart.valueChanged.connect(self.frame_range_changed)
        self.spnFrameEnd.valueChanged.connect(self.frame_range_changed)
        self.spnWidth.valueChanged.connect(self.resolution_changed)
        self.spnHeight.valueChanged.connect(self.resolution_changed)
        self.spnPixelAspect.valueChanged.connect(self.pixel_aspect_changed)
        self.txtName.textChanged.connect(self.name_changed)
        self.txtOutputPath.textChanged.connect(self.output_path_changed)
        self.cmbCamera.currentIndexChanged.connect(self.camera_changed)
        self.cmbSceneState.currentIndexChanged.connect(self.scene_state_changed)
        self.cmbPreset.currentIndexChanged.connect(self.preset_changed)
        self.cmbLayerPreset.currentIndexChanged.connect(self.layer_preset_changed)

    # noinspection PyAttributeOutsideInit
    def get_state_sets(self, state_set_parent):
        # https://help.autodesk.com/view/MAXDEV/2025/ENU/?guid=GUID-A2F530D0-628D-48DA-A126-6752A525E1FF
        # https://www.scriptspot.com/forums/3ds-max/general-scripting/a-question-on-scripting-state-set-render-output
        # stateSet_count = state_set_parent.DescendantStateCount
        stateSet_count = state_set_parent.Children.Count
        for state_set_index in range(-1, stateSet_count + stateSet_count):
            try:
                stateSet = state_set_parent.Children.Item[state_set_index]
                stateSet_name = stateSet.Name
            except (AttributeError, SystemError):
                continue
            state_set_type = "Autodesk.Max.StateSets.Entities.StateSets.StateSet"
            if stateSet_name != "Objects" and stateSet.GetType().ToString() == state_set_type:
                stateSetChildren = state_set_parent.Children.Count
                if stateSetChildren > 0:
                    yield from self.get_state_sets(stateSet)

                yield stateSet_name

    def warn_render_settings_open(self, title: str, message: str, buttons: dict = None):
        """
        Display a warning message with buttons.
        :param title: Warning message title
        :param message: Message to display
        :param buttons: Buttons names (keys) and their return value
        :return: Button press value OR None if the was window closed.
        """

        # Default buttons
        if not buttons:
            buttons = {"Yes": True, "No": False}

        # Construct buttons and their return value
        button_constructor = []
        button_results = {}
        for x, (key, value) in enumerate(buttons.items()):
            x += 1
            button_constructor.append((key, x))
            button_results[x] = [key, value]
        buttons = button_constructor

        # Display message
        dialog = GenericDialog(title=title, message=message, buttons=buttons, parent=self)
        dialog.setMinimumWidth(300)

        # Return result
        result = dialog.exec_()
        if result == QtWidgets.QDialog.DialogCode.Rejected:
            # if self.log_minor_actions:
            console.log(VerboseLevel.DEBUG, "User closed window.")
            result = None
        else:
            # if self.log_minor_actions:
            console.log(VerboseLevel.DEBUG, f"User clicked '{button_results[result][0]}'")
            result = button_results[result][1]

        return result

    def keyPressEvent(self, event):
        """Capture and ignore all key press events.

        This is used so that return key event does not trigger the exit button
        from the dialog. We need to allow the return key to be used in filters
        in the widget."""
        console.log(VerboseLevel.DEBUG, "Ignoring Enter/Return key")

    def on_close(self):
        # Do things on window close
        self.saveDialogData()

    def restoreWindowSettings(self):
        # Restore window position & size
        settings = QSettings(Config.PUBLISHER, Config.APP_NAME)
        # noinspection PyTypeChecker
        self.restoreGeometry(settings.value("geometry"))

    def restoreDialogData(self):
        try:
            # Retrieve table data
            data = rt.globalVars.get(rt.name(Config.BATCH_RENDER_SETTINGS))
            data = json.loads(data)
            console.log(VerboseLevel.DEBUG, data)

            for row_index, row_data in enumerate(data["table_data"]):
                self.tableWidget.add_row(True)
                for column_index, column_name in enumerate(columns.names):
                    value = row_data[column_name]
                    hidden_value = None
                    if column_name in row_data:
                        try:  # For previous versions that didn't have hidden values
                            value = row_data[column_name][0]
                            hidden_value = row_data[column_name][1]
                        except TypeError:
                            pass
                        except IndexError:
                            value = ""
                            hidden_value = ""
                    else:
                        # Catch for if new columns are added in a new script version
                        existing_item = self.tableWidget.cellWidget(row_index, column_index)
                        if existing_item:
                            value = False
                        else:
                            value = ""

                    if type(value) is str:
                        if hidden_value and type(hidden_value) in [int, str]:
                            # If hidden value, use hidden value to get node from ID
                            node = None
                            try:
                                node = get_item_by_id(hidden_value)
                            except ValueError:
                                try:
                                    node = get_item_by_name(value)
                                    hidden_value = get_item_unique_id(node)
                                    console.log(
                                        VerboseLevel.WARNING,
                                        f'Warning: Found "{value}" by name instead of internal ID.',
                                    )
                                except ValueError:
                                    console.log(
                                        VerboseLevel.ERROR,
                                        f"Error: Can't find node: {value} ({hidden_value})",
                                    )
                                    value = f"*** ERROR ({value}) ***"

                            if node and node in rt.Cameras:
                                value = node.name

                        existing_item = self.tableWidget.item(row_index, column_index)
                        new_item = QtWidgets.QTableWidgetItem(existing_item)
                        new_item.setText(value)
                        self.tableWidget.setCellData(
                            row_index, column_index, new_item, hidden_value
                        )

                    elif type(value) is bool:
                        existing_item = self.tableWidget.cellWidget(row_index, column_index)
                        # noinspection PyUnresolvedReferences
                        existing_item.setChecked(value)
                        self.tableWidget.setCellWidget(row_index, column_index, existing_item)

            # Restore default output path
            default_output_path = data["default_output_path"]
            if not default_output_path:
                default_output_path = self.renderOutput
            self.txtDefaultOutputPath.setText(default_output_path)

            # Expand columns to fit column contents
            self.tableWidget.resize_column_to_contents()

            console.log(VerboseLevel.INFO, "Restored dialog data")
        except KeyError as e:
            console.log(VerboseLevel.ERROR, "Save data corrupt:")
            console.log(VerboseLevel.ERROR, str(e))

    def saveDialogData(self):
        # Create config.PUBLISHER values
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, f"Software\\{Config.PUBLISHER}", 0, winreg.KEY_WRITE
        )
        winreg.SetValueEx(key, "creator", 0, winreg.REG_SZ, Config.PUBLISHER_CREATOR)
        winreg.CloseKey(key)

        # Save window position & size
        settings = QSettings(Config.PUBLISHER, Config.APP_NAME)
        settings.setValue("geometry", self.saveGeometry())
        settings.setValue("author", Config.AUTHOR)
        settings.setValue("version", Config.APP_VERSION)

        # Get table data
        table_data = []
        for row_num in range(self.tableWidget.rowCount()):
            # log('----------------------------------')
            row_data = {}
            for column_num in range(self.tableWidget.columnCount()):
                column_name = self.tableWidget.horizontalHeaderItem(column_num).text()

                item = self.tableWidget.item(row_num, column_num)
                hidden_value = self.tableWidget.getHiddenValue(row_num, column_num)
                item_cell = self.tableWidget.cellWidget(row_num, column_num)

                if isinstance(item_cell, QtWidgets.QCheckBox):
                    val = item_cell.isChecked()
                    # log(f'{column_name} = bool = {val}')
                else:
                    val = item.text()
                    # log(f'{column_name} = text = {[val, hidden_value]}')

                row_data[column_name] = [val, hidden_value]
            table_data.append(row_data)
            # log(row_data)

        # Get default output path
        default_output_path = self.txtDefaultOutputPath.text()

        # Save data
        data = str(
            json.dumps({"default_output_path": default_output_path, "table_data": table_data})
        )

        try:
            old_data = rt.globalVars.get(rt.name(Config.BATCH_RENDER_SETTINGS))
            old_data = json.loads(old_data)
        except RuntimeError:
            # In case it's a new project and there isn't old data
            old_data = None

        if data != old_data:
            rt.batchRenderSettings = data
            rt.persistents.make(rt.name(Config.BATCH_RENDER_SETTINGS))
            rt.setSaveRequired(True)  # Makes 3ds max know that changes occurred to the file
            console.log(VerboseLevel.DEBUG, data)
            console.log(VerboseLevel.INFO, "Data saved!")

    """General Functions """

    def update_element_values(self):
        console.log(VerboseLevel.DEBUG, "update_element_values")

        if not self.system_modified:
            self.system_modified = True
            selected_item = self.tableWidget.table_get_selected()[0]

            if selected_item is not None:
                default_values = ["Default", "", None]

                # Set "Name"
                name_value = self.tableWidget.item(selected_item, 1).text()
                self.txtName.setText(name_value)

                # Set "Camera"
                value = self.tableWidget.item(selected_item, 2).text()
                self.cmbCamera.setCurrentText(value)

                # Set "Output Path"
                value = self.tableWidget.item(selected_item, 3).text()
                if value in [Config.DEFAULT_PATH_TEXT, "", None]:
                    path = self.txtDefaultOutputPath.text()
                    value = os.path.join(str(path), name_value + ".exr")
                    self.toggle_override_fields(False, "Output Path")
                else:
                    self.toggle_override_fields(True, "Output Path")
                self.txtOutputPath.setText(value)

                # Set "Frame Start" & "Frame End"
                value = self.tableWidget.item(selected_item, 4).text()
                if value in default_values:
                    value = [rt.rendStart, rt.rendEnd]
                    self.frameRangeOverride.setCheckState(Qt.CheckState.Unchecked)
                    self.toggle_override_fields(False, "Frame Range")
                else:
                    value = value.split(":")
                    self.frameRangeOverride.setCheckState(Qt.CheckState.Checked)
                    self.toggle_override_fields(True, "Frame Range")
                self.spnFrameStart.setValue(int(value[0]))
                self.spnFrameEnd.setValue(int(value[1]))

                # Set "Width" & "Height"
                value = self.tableWidget.item(selected_item, 5).text()
                if value in default_values:
                    value = [rt.renderWidth, rt.renderHeight]
                    self.imageSizeOverride.setCheckState(Qt.CheckState.Unchecked)
                    self.toggle_override_fields(False, "Image Size")
                else:
                    value = value.split("x")
                    self.imageSizeOverride.setCheckState(Qt.CheckState.Checked)
                    self.toggle_override_fields(True, "Image Size")
                self.spnWidth.setValue(int(value[0]))
                self.spnHeight.setValue(int(value[1]))

                # Set "Pixel Aspect"
                value = self.tableWidget.item(selected_item, 6).text()
                if value in default_values:
                    value = rt.renderPixelAspect
                    self.pixelAspectOverride.setCheckState(Qt.CheckState.Unchecked)
                    self.toggle_override_fields(False, "Pixel Aspect")
                else:
                    self.pixelAspectOverride.setCheckState(Qt.CheckState.Checked)
                    self.toggle_override_fields(True, "Pixel Aspect")
                self.spnPixelAspect.setValue(float(value))

                # Set "Scene State"
                value = self.tableWidget.item(selected_item, 7).text()
                if value:
                    # value = value.text()
                    # index = self.cmbSceneState.findText(value)
                    self.cmbSceneState.setCurrentText(value)
                else:
                    self.cmbSceneState.setCurrentText(Config.DEFAULT_TEXT)

                # Set "Render Preset"
                value = self.tableWidget.item(selected_item, 8).text()
                if value:
                    # value = value.text()
                    # index = self.cmbPreset.findText(value)
                    self.cmbPreset.setCurrentText(value)
                else:
                    self.cmbPreset.setCurrentText(Config.DEFAULT_TEXT)

                # Set "Layer Preset"
                value = self.tableWidget.item(selected_item, 9).text()
                if value:
                    # value = value.text()
                    # index = self.cmbLayerPreset.findText(value)
                    self.cmbLayerPreset.setCurrentText(value)
                else:
                    self.cmbLayerPreset.setCurrentText(Config.DEFAULT_TEXT)

            self.system_modified = False

    def toggle_override_fields(self, disable_fields: bool, fields: str = "All"):
        # if self.log_functions: log("toggle_override_fields")

        system_modified = self.system_modified
        self.system_modified = True

        if disable_fields:
            state_string = "Enabled"
        else:
            state_string = "Disabled"

        # Disable/enable fields based on the state
        if fields in ["Frame Range", "All"]:
            # self.frameRangeOverride.setChecked(disable_fields)
            self.lblFrameStart.setEnabled(disable_fields)
            self.spnFrameStart.setEnabled(disable_fields)
            self.lblFrameEnd.setEnabled(disable_fields)
            self.spnFrameEnd.setEnabled(disable_fields)
            console.log(VerboseLevel.DEBUG, f"{state_string} Frame Range")
        if fields in ["Image Size", "All"]:
            # self.imageSizeOverride.setChecked(disable_fields)
            self.lblWidth.setEnabled(disable_fields)
            self.spnWidth.setEnabled(disable_fields)
            self.lblHeight.setEnabled(disable_fields)
            self.spnHeight.setEnabled(disable_fields)
            console.log(VerboseLevel.DEBUG, f"{state_string} Image Size")
        if fields in ["Pixel Aspect", "All"]:
            # self.pixelAspectOverride.setChecked(disable_fields)
            self.lblPixelAspect.setEnabled(disable_fields)
            self.spnPixelAspect.setEnabled(disable_fields)
            console.log(VerboseLevel.DEBUG, f"{state_string} Pixel Aspect")
        if fields in ["Output Path", "All"]:
            self.outputPathOverride.setChecked(disable_fields)
            self.lblOutputPath.setEnabled(disable_fields)
            self.btnBrowse.setEnabled(disable_fields)
            self.txtOutputPath.setEnabled(disable_fields)
            self.btnClear.setEnabled(disable_fields)
            console.log(VerboseLevel.DEBUG, f"{state_string} Output Path")

        self.system_modified = system_modified

    """Action functions when buttons pushed"""

    def clear_default_output_path(self):
        self.txtDefaultOutputPath.setText(self.renderOutput)

    def browse_default_output_path(self):
        # Open the 3ds Max Save Folder dialog
        current_path = self.txtDefaultOutputPath.text()
        if not os.path.exists(current_path):
            current_path = self.renderOutput

        folder_path = rt.getSavePath(
            # caption="Select Folder Location",
            initialDir=current_path
        )

        # Update the text field with the selected file path
        if folder_path:
            self.txtDefaultOutputPath.setText(folder_path)

        return folder_path

    def browse_output_path(self):
        # Open the 3ds Max Save File dialog
        value = self.txtOutputPath.displayText()

        file_path = rt.getSaveFileName(
            caption="Select Image Save Location",
            filename=value,
            types="Image Files (*.png *.jpg *.bmp *.tga *.exr)|*.png;*.jpg;*.bmp;*.tga;*.exr|All Files (*.*)|*.*",
        )

        # Deconstruct file path (to remove additional file extensions)
        file_type = rt.getFilenameType(file_path)
        file_name = rt.getFilenameFile(file_path)
        file_name = file_name.split(file_type)[0]
        file_path = rt.getFilenamePath(file_path)

        # Reconstruct file path
        file_path = os.path.join(file_path, file_name + file_type)

        # Update the text field with the selected file path
        if file_path:
            self.txtOutputPath.setText(str(file_path))

        return file_path

    def clear_output_path(self):
        name_value = self.txtName.text()
        path = self.txtDefaultOutputPath.text()
        value = os.path.join(str(path), name_value + ".exr")
        self.txtOutputPath.setText(value)

    def start_batch_render(self, pre_check: bool = False):
        def parse_number_string(number_string: str) -> list[int]:
            """
            Parses a string of numbers and ranges into a list of integers.

            Args:
                number_string (str): String containing numbers and ranges (e.g., '1, 3-5, 8').

            Returns:
                list: List of integers parsed from the input string.
            """
            result = []
            parts = number_string.split(",")
            for part in parts:
                if "-" in part:  # Check if the part represents a range
                    start, end = map(int, part.split("-"))
                    result.extend(
                        range(start, end + 1)
                    )  # Add the range of numbers to the result
                else:
                    result.append(int(part.strip()))  # Add the single number to the result
            return result

        def were_there_render_errors(time_of_render: datetime):
            """
            Checks if there were any V-Ray log errors during the render.

            Args:
                time_of_render (datetime): Datetime object for the time of render.

            Returns:
                str: Error message if an error was found, None otherwise.
            """
            time_format = "[%Y/%b/%d|%H:%M:%S]"  # EX: [2024/Aug/21|11:37:54]
            vray_error_log_path = os.path.join(tempfile.gettempdir(), "vraylog.txt")
            with open(vray_error_log_path) as file:
                for line in file:
                    line = line.strip()
                    line_timestamp_formatted = re.search("\[.*?]", line).group()
                    line_timestamp_datetime = datetime.strptime(
                        line_timestamp_formatted, time_format
                    )
                    if line_timestamp_datetime >= time_of_render and "error: " in line:
                        error_message = line.replace(
                            line_timestamp_formatted + " error", VerboseLevel.ERROR.name
                        ).strip()
                        return error_message
            return None  # If no error message found

        def get_valid_filename(item_values: dict):
            """
            Checks and returns valid file name.
            If render name is the default value, returns camera name with "_VRayPhysicalCamera" removed.

            Checks if a file name is valid based on common restrictions.
            Check if file name is empty
            Check if file name starts with a dot (hidden file on some systems)
            Check if file name only contains valid characters (alphanumeric, underscore, hyphen)

            :param item_values: dictionary containing row properties
            :return: Valid file name.
            """
            item_values = item_values.copy()
            item_values["Camera"] = item_values["Camera"].replace("_VRayPhysicalCamera", "")
            original_render_name = render_name = item_values["Name"]

            if render_name == "Default":
                render_name = "{Camera}"

            replace_flags = {
                "{Camera}": item_values["Camera"],
                "{Scene State}": item_values["Scene_State"],
                "{State Set}": item_values["Scene_State"],
                "{Render Preset}": item_values["Render_Preset"],
                "{Layer Preset}": item_values["Layer_Preset"],
                "{Resolution}": item_values["Resolution_Display"],
                "{Pixel Aspect}": item_values["Pixel_Aspect"],
            }
            blank_values = False
            for replace, replace_with in replace_flags.items():
                if replace.lower() in render_name.lower():
                    if replace_with is None:
                        replace_with = ""
                    if replace_with.lower() in ["default", ""]:
                        replace_with = ""
                        blank_values = True
                    render_name = re.sub(
                        replace, replace_with, render_name, flags=re.IGNORECASE
                    )

            if blank_values and pre_check:
                message = (
                    "Some of the replacement flags in the render name are blank or default. "
                    f"This may result in weird file names.\n\n"
                    f"Original render name: {original_render_name}\n"
                    f"Render name: {render_name}\n\nDo you want to proceed?"
                )
                continue_with_render = self.warn_render_settings_open(
                    "Warning!", message, {"Yes": True, "No": False}
                )
                if not continue_with_render:
                    raise ValueError(
                        f'Canceled because of blank replacement flags in "{original_render_name}"'
                    )

            invalid_chars = '\/:*?<>|"\\'  # Commonly restricted characters
            if (
                not render_name
                or any(char in render_name for char in invalid_chars)
                or render_name.startswith(".")
            ):
                raise ValueError(f"Image name invalid!: {render_name}")

            return render_name

        def get_output_path(output_path: str, name: str) -> Tuple[Path, Path]:
            """Get image output paths of render

            Args:
                output_path: Output path or config.DEFAULT_PATH_TEXT
                name: Render name

            Returns:

            """
            if output_path == Config.DEFAULT_PATH_TEXT:
                path = Path(self.txtDefaultOutputPath.text() + name)
            else:
                path = Path(output_path)
            path = PROJECT_ROOT_PATH / path

            output_path_exr = path.with_suffix(".exr")
            output_path_png = path.with_suffix(".png")

            if not path.parent.exists():
                raise ValueError(f"Output directory doesnt exist!: {path.parent}")

            return output_path_exr, output_path_png

        def get_frame_range(frame_range_str: str) -> list[int]:
            """Get frame range of render

            Args:
                frame_range_str: String representation of frames to render

            Returns:
                List of frames to render
            """
            if frame_range_str == "Default":
                renderTimeType = rt.rendTimeType
                if renderTimeType == 1:  # Single frame
                    return [rt.currentTime.frame]
                elif renderTimeType == 2:  # Active time segment
                    return list(
                        range(
                            int(rt.animationRange.start.frame),
                            int(rt.animationRange.end.frame) + 1,
                        )
                    )
                elif renderTimeType == 3:  # Frame ange
                    return list(
                        range(
                            int(rt.rendStart.frame),
                            int(rt.rendEnd.frame) + 1,
                            int(rt.rendNThFrame),
                        )
                    )
                elif renderTimeType == 4:  # Picked frames
                    return parse_number_string(rt.rendPickupFrames)
                else:
                    raise ValueError(f"Invalid render time type: {renderTimeType}")

            else:
                frame_range_array = frame_range_str.split(":")
                if frame_range_array[0] == frame_range_array[-1]:  # Single frame "EX: 1:1"
                    return [int(frame_range_array[0])]
                else:
                    frame_range = [int(x) for x in frame_range_array]
                    return list(range(frame_range[0], frame_range[-1] + 1))

        def get_frame_range_display(frame_range):
            if isinstance(frame_range, range):
                if frame_range.step == 1:
                    frame_range_disp = f"{frame_range.start}-{frame_range.stop}"
                else:
                    frame_range_disp = [str(int(x)) for x in list(frame_range)]
                    frame_range_disp = ", ".join(frame_range_disp)
            else:
                frame_range_disp = [str(int(x)) for x in list(frame_range)]
                frame_range_disp = ", ".join(frame_range_disp)
            return frame_range_disp

        def get_resolution_disp(resolution, original_width, original_height):
            if resolution == "Default":
                resolution = [original_width, original_height]
            else:
                resolution = resolution.split("x")
            resolution_disp = "x".join(str(x) for x in resolution)

            return resolution, resolution_disp

        def get_pixel_aspect(pixel_aspect):
            if pixel_aspect == "Default":
                pixel_aspect = rt.renderPixelAspect
            else:
                rt.renderPixelAspect = pixel_aspect

            return pixel_aspect

        def get_and_set_scene_state(scene_state_disp):
            # https://help.autodesk.com/view/MAXDEV/2025/ENU/?guid=GUID-9BF5F52D-105D-46BA-8F24-D6B56529ED99
            not_found = True
            if scene_state_disp:
                if self.stateSet_prefix in scene_state_disp:
                    potential_stateSet = scene_state_disp.replace(self.stateSet_prefix, "")
                    stateSet = self.masterState.GetDescendant(potential_stateSet)
                    if stateSet:
                        self.masterState.SetCurrentStateSet([stateSet])  # Restore state set
                        scene_state_disp = potential_stateSet
                        not_found = False

                elif self.sceneState_prefix in scene_state_disp:
                    sceneState = scene_state_disp.replace(self.sceneState_prefix, "")
                    is_SceneState = rt.sceneStateMgr.FindSceneState(sceneState)
                    if is_SceneState:
                        rt.sceneStateMgr.RestoreAllParts(sceneState)  # Restore scene state
                        scene_state_disp = sceneState
                        not_found = False

                if not_found:
                    raise ValueError(
                        f"Scene State or State Set '{scene_state_disp}' does not exist."
                    )

            else:
                scene_state_disp = None

            return scene_state_disp

        def get_and_set_render_preset(preset_disp):
            if preset_disp:
                file_path = rt.GetDir(rt.Name("renderPresets"))
                file_path = os.path.join(file_path, preset_disp)
                rt.renderPresets.LoadAll(0, file_path)  # Restore render preset
            else:
                preset_disp = None
            return preset_disp

        def get_and_set_layer_preset(layer_preset_disp):
            if layer_preset_disp:
                file_path = rt.GetDir(rt.Name("vpost"))
                file_path = os.path.join(file_path, layer_preset_disp)
                rt.vfbLayerMgr.loadLayersFromFile(file_path)  # Set layer preset
            else:
                layer_preset_disp = None

            return layer_preset_disp

        def prevent_duplicate_files(output_path, frame, formatted_time):
            # Check if the image file already exists, if so, adjustment a duplicate number to not overwrite the original
            directory, file_name = os.path.split(output_path)
            filename, file_extension = os.path.splitext(file_name)
            dupe_count = 0
            dupe_count_str = ""
            original_full_filename = None
            while True:
                # Image file name EX: View 1_0000_24-06-18T14.18.45 (1).exr
                file_path = os.path.join(
                    directory,
                    f"{filename}_{int(frame):04d}_"
                    f"{formatted_time}{dupe_count_str}{file_extension}",
                )

                # Retain original name for the warning message
                if not original_full_filename:
                    original_full_filename = file_path

                if os.path.exists(file_path):
                    dupe_count += 1
                    dupe_count_str = f" ({dupe_count})"
                else:
                    break

            directory, file_name = os.path.split(file_path)
            if pre_check and dupe_count > 0:
                title = "Image already exists!"
                message = (
                    "An image with the same name already exists. "
                    "A number will automatically be append to the end of the file name to prevent overwriting. "
                    "If this is not the desired result, move or rename the existing file."
                    f"\nOriginal name: {original_full_filename}"
                    f"\nNew name:      {file_name}"
                )
                self.warn_render_settings_open(title, message, {"OK": True})

            return file_path, file_name

        def check_VFB_settings():
            # noinspection LongLine
            """
            Checks if VRay GPU is the current renderer and if test settings that could be unintentionally left on are on.
            Cancels render if VRay GPU is not set as current renderer.
            Warns user if test settings are on and asks if they want to continue

            Returns:
                bool: True to continue with render, False otherwise.
                bool: True to close renderSceneDialog after checking settings.
            """
            message = None
            close_renderSceneDialog = False
            vr = rt.renderers.current
            if "V_Ray" in str(vr):  # or "V_Ray_GPU"
                region_enabled = rt.vrayVFBGetRegionEnabled()
                testRes_enabled = "1" in str(rt.execute("vfbControl #testresolution"))
                followMouse_enabled = "1" in str(rt.execute("vfbControl #trackmouse"))
                debugShading_enabled = "1" in str(rt.execute("vfbControl #debugshading"))
                renderSettings_open = rt.renderSceneDialog.isOpen()
                problem_settings = {
                    "Region render": region_enabled,
                    "Test resolution": testRes_enabled,
                    "Follow mouse": followMouse_enabled,
                    "Debug shading": debugShading_enabled,
                }
                if True in problem_settings.values():
                    # Format warning message
                    message = [key for key, value in problem_settings.items() if value]
                    if len(message) == 1:
                        message = message[0] + " is enabled"
                    elif len(message) == 2:
                        message = " and ".join(message) + " are enabled"
                    elif len(message) >= 3:
                        message = (
                            ", ".join(message[:-1]) + ", and " + message[-1] + " are enabled"
                        )

                if renderSettings_open:
                    rt.renderSceneDialog.commit()
                else:
                    # Open renderSceneDialog to prevent override settings from sticking.
                    close_renderSceneDialog = True
                    rt.renderSceneDialog.open()

            else:
                message = "V-Ray is not set as current renderer!"

            # Display warning
            continue_with_render = True
            if message and pre_check:
                message += ". Do you want to proceed?"
                continue_with_render = self.warn_render_settings_open(
                    "Warning!", message, {"Yes": True, "No": False}
                )

            return continue_with_render, close_renderSceneDialog

        def check_if_image_was_created(image_path):
            vray_log_error = None
            if os.path.exists(image_path):
                # If file exists, check weather it was modified recently (May not be necessary)
                last_modified_time = os.path.getmtime(image_path)
                current_time = time.time()
                time_difference = current_time - last_modified_time
                if time_difference >= 60:  # Seconds
                    console.log(VerboseLevel.WARNING, f"Time difference: {time_difference}")
                    vray_log_error = (
                        "Waring: File seems too old in relation to when the image finished rendering. "
                        "The image may not have written to the file correctly."
                    )
            else:
                vray_log_error = f"Error: File doesn't exist!"

            return vray_log_error

        def do_render(
            camera: object,
            frame: int,
            outputfile_exr: str,
            outputfile_png,
            pixel_aspect: int,
            image_width: int,
            image_height: int,
        ):
            """
            Initiates rendering and checks for errors and if user cancels the render.
            Args:
                camera (object): 3ds max camera object.
                frame (int): Frame to render.
                outputfile (str): Absolute image file path to save to.
                pixel_aspect (int): Image pixel aspect.
                image_width (int): Image pixel width.
                image_height (int): Image pixel height.

            Returns:
                int:
                    0 - If rendering was successful.
                    1 - If the user canceled rendering.
                    2 - If rendering failed.
            """
            vray_log_error = None
            time_of_render = datetime.now()
            self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, False)

            rt.execute("vfbControl #clearimage")

            r, canceled = rt.render(
                camera=camera,
                frame=frame,
                outputwidth=image_width,
                outputheight=image_height,
                pixelaspect=pixel_aspect,
                outputfile=outputfile_exr,
                vfb=False,
                progressbar=False,
                cancelled=pymxs.byref(None),
                # to=rt.Bitmap(1, 1)
            )

            # Get render elapsed time
            render_elapsed_time = datetime.now() - time_of_render
            time_difference_str = (
                f"{render_elapsed_time.seconds // 3600}:"
                f"{(render_elapsed_time.seconds // 60) % 60}:"
                f"{render_elapsed_time.seconds % 60}"
            )

            # Check if render was canceled
            if not canceled:
                vray_log_error = check_if_image_was_created(outputfile_exr)
            else:
                vray_log_error = were_there_render_errors(time_of_render)

            while True:
                if not canceled and not vray_log_error:
                    n = 0
                    while n < 3:
                        n += 1
                        rt.execute("vfbControl #setchannel 0")
                        outputfile_png = outputfile_png.replace("\\", "/")
                        rt.execute(f'vfbControl #saveimage "{outputfile_png}"')
                        vray_log_error = None
                        vray_log_error = check_if_image_was_created(outputfile_png)

                        if not vray_log_error:
                            break

                    if vray_log_error:
                        continue  # Loop back to the top to handle error

                    # noinspection LongLine
                    console.log(
                        VerboseLevel.INFO,
                        f"Image rendered successfully! Elapsed time: {time_difference_str}",
                    )
                    console.log(VerboseLevel.INFO, f"Output path: {outputfile_exr}")
                    return 0
                elif canceled and not vray_log_error:
                    console.log(
                        VerboseLevel.WARNING,
                        f"Render was canceled! Elapsed time: {time_difference_str}",
                    )
                    return 1
                else:
                    if not vray_log_error:
                        vray_log_error = "Unknown error!"
                    console.log(
                        VerboseLevel.ERROR,
                        f"Render failed! Elapsed time: {time_difference_str}",
                    )
                    console.log(VerboseLevel.ERROR, vray_log_error)
                    return 2

        def main():
            if not pre_check:
                console.log(VerboseLevel.INFO, "Starting render pre-check...")
                total_rows_to_render = self.start_batch_render(True)
                if not total_rows_to_render:
                    return  # Cancel if total_rows_to_render is 0 or False
                else:
                    console.log(VerboseLevel.INFO, "Starting batch render...")
            else:
                total_rows_to_render = 0

            self.saveDialogData()  # Save dialog data
            render_queue_canceled = False
            python_render_error = None

            continue_with_render, close_renderSceneDialog = check_VFB_settings()
            if not continue_with_render:
                return

            # Render or skip each row
            total_rows = self.tableWidget.rowCount()
            rendered_rows = 0
            canceled_renders = 0
            max_canceled_before_exit = (
                2  # Set to 0 to disable canceling whole queue (not recommended)
            )
            with console.indent():
                for row_id in range(total_rows):
                    try:
                        rowProperties = self.tableWidget.get_entry_values(row_id)
                        if rowProperties["Use"]:
                            # Get parameters
                            name = rowProperties["Name"]
                            camera_name = rowProperties["Camera"]
                            camera_id = rowProperties["Camera_Hidden"]
                            output_path_exr = rowProperties["Output_Path"]
                            frame_range = rowProperties["Range"]
                            resolution = rowProperties["Resolution"]
                            pixel_aspect = rowProperties["Pixel_Aspect"]
                            scene_state_disp = rowProperties["Scene_State"]
                            preset_disp = rowProperties["Render_Preset"]
                            layer_preset_disp = rowProperties["Layer_Preset"]

                            if pre_check:
                                console.log(
                                    VerboseLevel.INFO, f"Checking: [{str(row_id + 1)}] {name}"
                                )
                                total_rows_to_render += 1

                            with console.indent(not pre_check):
                                # Prepare parameters
                                try:
                                    camera = get_item_by_id(camera_id)
                                except ValueError as e:
                                    console.log(
                                        VerboseLevel.ERROR,
                                        f"Couldn't find camera: " + str(camera_id),
                                    )
                                    raise e
                                frame_range = get_frame_range(frame_range)
                                frame_range_disp = get_frame_range_display(frame_range)

                                original_renderWidth = rt.renderWidth
                                original_renderHeight = rt.renderHeight
                                # noinspection LongLine
                                resolution, resolution_disp = get_resolution_disp(
                                    resolution, original_renderWidth, original_renderHeight
                                )
                                rowProperties["Resolution_Display"] = resolution_disp
                                pixel_aspect = get_pixel_aspect(pixel_aspect)

                                # Prepare scene
                                rowProperties["Scene_State"] = scene_state_disp = (
                                    get_and_set_scene_state(scene_state_disp)
                                )
                                preset_disp = get_and_set_render_preset(preset_disp)
                                layer_preset_disp = get_and_set_layer_preset(layer_preset_disp)

                                name = get_valid_filename(rowProperties)
                                output_path_exr, output_path_png = get_output_path(
                                    output_path_exr, name
                                )

                                rendered_rows += 1
                                if not pre_check:
                                    console.log(
                                        VerboseLevel.INFO,
                                        f"Rendering: [{str(row_id + 1)}] "
                                        f"({rendered_rows}/{total_rows_to_render}) {name} | "
                                        f"{camera_name} | {resolution_disp} | {frame_range_disp} | {str(pixel_aspect)} | "
                                        f"{scene_state_disp} | {preset_disp} | {layer_preset_disp}",
                                    )

                                with console.indent(pre_check):
                                    canceled_renders = render_frames(
                                        camera,
                                        canceled_renders,
                                        frame_range,
                                        max_canceled_before_exit,
                                        output_path_exr,
                                        output_path_png,
                                        pixel_aspect,
                                        resolution,
                                    )

                            # Reset render settings back to their original values
                            rt.renderWidth = original_renderWidth
                            rt.renderHeight = original_renderHeight

                            if (
                                max_canceled_before_exit
                                and canceled_renders >= max_canceled_before_exit
                                and not pre_check
                            ):
                                render_queue_canceled = True
                                console.log(VerboseLevel.WARNING, "Canceled render queue!")
                                break
                    except ValueError as e:
                        console.log(VerboseLevel.ERROR, str(e))
                        python_render_error = True
                        continue

            if close_renderSceneDialog:
                rt.renderSceneDialog.cancel()

            if not pre_check and not render_queue_canceled:
                console.log(VerboseLevel.INFO, "Rendering done!")
                return True
            elif pre_check and not python_render_error:
                console.log(VerboseLevel.INFO, "No errors found!")
                return total_rows_to_render
            elif pre_check and python_render_error:
                console.log(VerboseLevel.ERROR, "Error(s) found. Canceled render!")
                return False

        def render_frames(
            camera,
            canceled_renders: int,
            frame_range: list[int],
            max_canceled_before_exit: int,
            output_path_exr: Path,
            output_path_png: Path,
            pixel_aspect,
            resolution: list[Any] | Any,
        ) -> int:
            formatted_time = datetime.now().strftime("%y-%m-%dT%H.%M.%S")
            total_frames = len(frame_range)
            for x, frame in enumerate(frame_range):
                outputfile_exr, full_filename = prevent_duplicate_files(
                    output_path_exr, frame, formatted_time
                )
                outputfile_png, none = prevent_duplicate_files(
                    output_path_png, frame, formatted_time
                )
                if not pre_check:
                    # noinspection LongLine
                    console.log(
                        VerboseLevel.INFO,
                        f"Rendering frame {frame} ({x + 1}/{total_frames}) {full_filename}",
                    )
                    if not self.DEBUG_disable_rendering:
                        with console.indent():
                            render_result = do_render(
                                camera,
                                frame,
                                outputfile_exr,
                                outputfile_png,
                                pixel_aspect,
                                int(resolution[0]),
                                int(resolution[1]),
                            )
                    else:
                        render_result = 0
                    if render_result == 1:
                        canceled_renders += 1

                if max_canceled_before_exit and canceled_renders >= max_canceled_before_exit:
                    break
            return canceled_renders

        return main()

    """Macro functions"""

    def add_unlisted_cameras(self):
        # Get all cameras in table
        cameras = []
        col = columns.names.index("Camera")
        for row in range(self.tableWidget.rowCount()):
            camera_name = self.tableWidget.item(row, col).text()
            cameras.append(camera_name)

        # Add missing cameras
        added_rows = False
        for i in range(self.cmbCamera.count()):
            camera_name = self.cmbCamera.itemText(i)
            if camera_name not in cameras:
                self.tableWidget.add_row()
                self.cmbCamera.setCurrentText(camera_name)
                added_rows = True
                console.log(VerboseLevel.INFO, f"Added {camera_name}")

        if not added_rows:
            console.log(VerboseLevel.INFO, "No cameras to adjustment")

    def add_camera_sceneStateSet_combos(self):
        # Get all existing camera/scene state combos
        existing_cameras_sceneStates = []
        col_camera = columns.names.index("Camera")
        col_sceneState = columns.names.index("Scene State")
        for row in range(self.tableWidget.rowCount()):
            existing_camera = self.tableWidget.item(row, col_camera).text()
            existing_sceneState = self.tableWidget.item(row, col_sceneState).text()
            existing_cameras_sceneStates.append(existing_camera + existing_sceneState)

        # Get all cameras and scene states/sets but remove default entries
        all_cameras = [self.cmbCamera.itemText(i) for i in range(self.cmbCamera.count())]
        all_state_sets = [
            self.cmbSceneState.itemText(i) for i in range(self.cmbSceneState.count())
        ][1:]

        added_rows = False
        for camera_name in all_cameras:
            for state_set in all_state_sets:
                if camera_name + state_set not in existing_cameras_sceneStates:
                    self.tableWidget.add_row()
                    self.txtName.setText("{camera}_{state set}")
                    self.cmbCamera.setCurrentText(camera_name)
                    self.cmbSceneState.setCurrentText(state_set)
                    added_rows = True
                    console.log(VerboseLevel.INFO, f"Added {camera_name} + {state_set}")

        if not added_rows:
            console.log(
                VerboseLevel.INFO, "No cameras and scene state/sets combos to adjustment"
            )

    """Functions for when "Selected Batch Render Parameters" changed"""

    def output_path_override_toggled(self, state: bool):
        """Triggers when field in parameters (Not table) gets toggled"""
        console.log(VerboseLevel.DEBUG, "output_path_override_toggled")

        if not self.system_modified:
            current_value = self.txtOutputPath.text()
            if not current_value:
                self.clear_output_path()

            self.output_path_changed()
            self.toggle_override_fields(state, "Output Path")

    def frame_range_override_toggled(self, state: bool):
        """Triggers when field in parameters (Not table) gets toggled"""
        console.log(VerboseLevel.DEBUG, "frame_range_override_toggled")

        if not self.system_modified:
            self.frame_range_changed(state)
            self.toggle_override_fields(state, "Frame Range")

    def image_size_override_toggled(self, state: bool):
        """Triggers when field in parameters (Not table) gets toggled"""
        console.log(VerboseLevel.DEBUG, "image_size_override_toggled")

        if not self.system_modified:
            self.resolution_changed()
            self.toggle_override_fields(state, "Image Size")

    def pixel_aspect_override_toggled(self, state: bool):
        """Triggers when field in parameters (Not table) gets toggled"""
        console.log(VerboseLevel.DEBUG, "pixel_aspect_override_toggled")

        if not self.system_modified:
            self.pixel_aspect_changed()
            self.toggle_override_fields(state, "Pixel Aspect")

    def frame_range_changed(self, state: bool):
        """Triggers when field in parameters (Not table) gets changed"""
        if not self.system_modified:
            self.system_modified = True
            value = "Default"
            if state:
                value = str(self.spnFrameStart.value()) + ":" + str(self.spnFrameEnd.value())

            selected_rows = self.tableWidget.table_get_selected(None)
            if selected_rows[0]:
                for selection in selected_rows:
                    row = selection[0]
                    col = columns.names.index("Range")
                    existing_item = self.tableWidget.item(row, col)

                    if existing_item is not None:
                        # Create a new item with the existing attributes and set the new text
                        new_item = QtWidgets.QTableWidgetItem(existing_item)
                        new_item.setText(value)
                        self.tableWidget.setItem(row, col, new_item)

                console.log(VerboseLevel.DEBUG, "Frame Range changed")
                self.system_modified = False

    def resolution_changed(self):
        """Triggers when field in parameters (Not table) gets changed"""
        if not self.system_modified:
            self.system_modified = True
            value = "Default"
            if self.imageSizeOverride.checkState():
                value = str(self.spnWidth.value()) + "x" + str(self.spnHeight.value())

            selected_rows = self.tableWidget.table_get_selected(None)
            if selected_rows[0]:
                for selection in selected_rows:
                    row = selection[0]
                    col = columns.names.index("Resolution")
                    existing_item = self.tableWidget.item(row, col)

                    if existing_item is not None:
                        # Create a new item with the existing attributes and set the new text
                        new_item = QtWidgets.QTableWidgetItem(existing_item)
                        new_item.setText(value)
                        self.tableWidget.setItem(row, col, new_item)

                console.log(VerboseLevel.DEBUG, "Resolution changed")
                self.system_modified = False

    def pixel_aspect_changed(self):
        """Triggers when field in parameters (Not table) gets changed"""
        if not self.system_modified:
            self.system_modified = True
            value = "Default"
            if self.pixelAspectOverride.checkState():
                value = str(self.spnPixelAspect.value())

            selected_rows = self.tableWidget.table_get_selected(None)
            if selected_rows[0]:
                for selection in selected_rows:
                    row = selection[0]
                    col = columns.names.index("Pixel Aspect")
                    existing_item = self.tableWidget.item(row, col)

                    if existing_item is not None:
                        # Create a new item with the existing attributes and set the new text
                        new_item = QtWidgets.QTableWidgetItem(existing_item)
                        new_item.setText(value)
                        self.tableWidget.setItem(row, col, new_item)

                console.log(VerboseLevel.DEBUG, "Pixel Aspect changed")

            self.system_modified = False

    def name_changed(self):
        """Triggers when field in parameters (Not table) gets changed"""
        console.log(VerboseLevel.DEBUG, "name_changed")

        if not self.system_modified:
            self.system_modified = True
            value = self.txtName.text()

            selected_rows = self.tableWidget.table_get_selected(None)
            if selected_rows[0]:
                for selection in selected_rows:
                    row = selection[0]
                    col = columns.names.index("Name")
                    existing_item = self.tableWidget.item(row, col)

                    if existing_item is not None:
                        # Create a new item with the existing attributes and set the new text
                        new_item = QtWidgets.QTableWidgetItem(existing_item)
                        new_item.setText(value)

                        self.tableWidget.setItem(row, col, new_item)

                        if not self.outputPathOverride.checkState():
                            # Set Output Path
                            col = columns.names.index("Output Path")
                            if self.outputPathOverride.isChecked():
                                path_value = self.txtOutputPath.text()
                            else:
                                path_value = QtWidgets.QTableWidgetItem(
                                    Config.DEFAULT_PATH_TEXT
                                )
                            self.tableWidget.setItem(row, col, path_value)

                console.log(VerboseLevel.DEBUG, "Name changed")
                self.tableWidget.resize_column_to_contents()
                self.system_modified = False

    def output_path_changed(self):
        """Triggers when field in parameters (Not table) gets changed"""
        console.log(VerboseLevel.DEBUG, "output_path_changed")

        if not self.system_modified:
            self.system_modified = True

            if self.outputPathOverride.isChecked():
                value = self.txtOutputPath.text()
            else:
                value = Config.DEFAULT_PATH_TEXT

            selected = self.tableWidget.table_get_selected(None)
            if selected != (None, None):
                for selection in selected:
                    row = selection[0]
                    col = columns.names.index("Output Path")
                    existing_item = self.tableWidget.item(row, col)

                    if existing_item is not None:
                        # Create a new item with the existing attributes and set the new text
                        new_item = QtWidgets.QTableWidgetItem(existing_item)
                        new_item.setText(value)
                        self.tableWidget.setItem(row, col, new_item)

            self.system_modified = False

    def camera_changed(self):
        """Triggers when field in parameters (Not table) gets changed"""
        console.log(VerboseLevel.DEBUG, "camera_changed")

        if not self.system_modified:
            self.system_modified = True
            camera_Name = self.cmbCamera.currentText()
            current_index = self.cmbCamera.currentIndex()
            camera_ID = self.cmbCamera.itemData(current_index)
            # log(camera_ID)

            for selection in self.tableWidget.table_get_selected(None):
                row = selection[0]
                col = columns.names.index("Camera")
                existing_item = self.tableWidget.item(row, col)

                if existing_item is not None:
                    # Create a new item with the existing attributes and set the new text
                    new_item = QtWidgets.QTableWidgetItem(existing_item)
                    new_item.setText(camera_Name)
                    self.tableWidget.setCellData(row, col, new_item, camera_ID)

            console.log(VerboseLevel.INFO, f"Camera changed: {camera_Name} ({camera_ID})")

            self.tableWidget.resize_column_to_contents()
            self.system_modified = False

    def scene_state_changed(self):
        """Triggers when field in parameters (Not table) gets changed"""
        if not self.system_modified:
            self.system_modified = True
            value = self.cmbSceneState.currentText()
            if value == Config.DEFAULT_TEXT:
                value = ""

            for selection in self.tableWidget.table_get_selected(None):
                row = selection[0]
                col = columns.names.index("Scene State")
                existing_item = self.tableWidget.item(row, col)

                if existing_item is not None:
                    # Create a new item with the existing attributes and set the new text
                    new_item = QtWidgets.QTableWidgetItem(existing_item)
                    new_item.setText(value)
                    self.tableWidget.setItem(row, col, new_item)

            console.log(VerboseLevel.DEBUG, "Scene State changed")

            self.system_modified = False

    def preset_changed(self):
        """Triggers when field in parameters (Not table) gets changed"""
        if not self.system_modified:
            self.system_modified = True

            value = self.cmbPreset.currentText()
            if value == Config.DEFAULT_TEXT:
                value = ""

            for selection in self.tableWidget.table_get_selected(None):
                row = selection[0]
                col = columns.names.index("Render Preset")
                existing_item = self.tableWidget.item(row, col)

                if existing_item is not None:
                    # Create a new item with the existing attributes and set the new text
                    new_item = QtWidgets.QTableWidgetItem(existing_item)
                    new_item.setText(value)
                    self.tableWidget.setItem(row, col, new_item)

            console.log(VerboseLevel.DEBUG, "Preset changed")

            self.system_modified = False

    def layer_preset_changed(self):
        """Triggers when field in parameters (Not table) gets changed"""
        console.log(VerboseLevel.DEBUG, "layer_preset_changed")

        if not self.system_modified:
            self.system_modified = True

            value = self.cmbLayerPreset.currentText()
            if value == Config.DEFAULT_TEXT:
                value = ""

            for selection in self.tableWidget.table_get_selected(None):
                row = selection[0]
                col = columns.names.index("Layer Preset")
                existing_item = self.tableWidget.item(row, col)

                if existing_item is not None:
                    # Create a new item with the existing attributes and set the new text
                    new_item = QtWidgets.QTableWidgetItem(existing_item)
                    new_item.setText(value)
                    self.tableWidget.setItem(row, col, new_item)

            console.log(VerboseLevel.DEBUG, "Layer Preset changed")

            self.system_modified = False


if __name__ == "__main__":
    pass
