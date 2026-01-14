"""
Mock pymxs module for testing outside of 3ds Max.
This provides a minimal implementation of the pymxs runtime that can be used for testing.
"""

import sys
from typing import Any, Dict, Optional, List, Union, Callable


class PathConfig:
    """Mock implementation of pathConfig"""

    def getCurrentProjectFolder(self):
        """Mock implementation of GetCurrentProjectFolder"""
        return "C:\\Users\\Username\\Documents\\3ds Max 2024"


class MockRuntime:
    """Mock implementation of pymxs.runtime"""

    def __init__(self):
        self._max_objects = {}
        self._next_handle = 1
        self._listeners = {}

        # Add some common attributes that might be accessed
        self.selection = []
        self.root_scene = {}
        self.renderers = MockRenderers()
        self.vray = MockVray()
        self.pathConfig = PathConfig()  # Add pathConfig attribute

    def __getattr__(self, name):
        """Dynamically handle attribute access to simulate MAXScript objects"""
        if name.startswith("_"):
            raise AttributeError(name)

        # Return a callable that can be used as a function
        def method(*args, **kwargs):
            print(f"[MOCK] Called {name} with args: {args}, kwargs: {kwargs}")
            return None

        return method

    def Name(self, obj):
        """Mock implementation of Name function"""
        return getattr(obj, "name", "MockObject")

    def Execute(self, script: str):
        """Mock implementation of Execute"""
        print(f"[MOCK] Executing script:\n{script}")
        return None

    def GetQuietMode(self) -> bool:
        """Mock implementation of GetQuietMode"""
        return False

    def SetQuietMode(self, value: bool):
        """Mock implementation of SetQuietMode"""
        pass

    def AddRollout(self, *args, **kwargs):
        """Mock implementation of AddRollout"""
        print("[MOCK] Added rollout")
        return None

    def CreateDialog(self, *args, **kwargs):
        """Mock implementation of CreateDialog"""
        print("[MOCK] Created dialog")
        return None

    def GetMAXWindowPos(self):
        """Mock implementation of GetMAXWindowPos"""
        return [0, 0, 100, 100]

    def GetMAXIniFile(self):
        """Mock implementation of GetMAXIniFile"""
        return "C:\\path\\to\\3dsmax.ini"

    def maxVersion(self):
        """Mock implementation of maxVersion"""
        return ["", "2024", ""]

    def GetDir(self, which):
        """Mock implementation of GetDir"""
        dirs = {
            "maxroot": "C:\\Program Files\\Autodesk\\3ds Max 2024",
            "scripts": "C:\\Users\\Username\\Documents\\3ds Max 2024\\scripts",
            "userScripts": "C:\\Users\\Username\\Documents\\3ds Max 2024\\scripts\\Startup",
            "userMacros": "C:\\Users\\Username\\Documents\\3ds Max 2024\\usermacros",
            "export": "C:\\Users\\Username\\Documents\\3ds Max 2024\\export",
            "import": "C:\\Users\\Username\\Documents\\3ds Max 2024\\import",
            "renderOutput": "C:\\Users\\Username\\Documents\\3ds Max 2024\\renderoutput",
            "scene": "C:\\Users\\Username\\Documents\\3ds Max 2024\\scenes",
            "image": "C:\\Users\\Username\\Documents\\3ds Max 2024\\sceneassets\\images",
            "autoBackup": "C:\\Users\\Username\\Documents\\3ds Max 2024\\autoback",
            "preview": "C:\\Users\\Username\\Documents\\3ds Max 2024\\previews",
            "plugcfg": "C:\\Users\\Username\\Documents\\3ds Max 2024\\plugcfg",
            "vray": "C:\\Program Files\\Chaos Group\\V-Ray\\3dsmax 2024\\bin",
            "defaults": "C:\\Program Files\\Autodesk\\3ds Max 2024\\en-US\\defaults",
            "startupScripts": "C:\\Program Files\\Autodesk\\3ds Max 2024\\scripts\\Startup",
            "ui_ln": "C:\\Program Files\\Autodesk\\3ds Max 2024\\ui",
            "ui": "C:\\Program Files\\Autodesk\\3ds Max 2024\\ui",
            "userIcons": "C:\\Users\\Username\\Documents\\3ds Max 2024\\usermacros\\icons",
            "userIcons_ln": "C:\\Users\\Username\\Documents\\3ds Max 2024\\usermacros\\icons",
            "defaultIcons": "C:\\Program Files\\Autodesk\\3ds Max 2024\\UI_ln\\Icons",
            "defaultIcons_ln": "C:\\Program Files\\Autodesk\\3ds Max 2024\\UI_ln\\Icons",
            "icons": "C:\\Program Files\\Autodesk\\3ds Max 2024\\UI_ln\\Icons",
            "icons_ln": "C:\\Program Files\\Autodesk\\3ds Max 2024\\UI_ln\\Icons",
            "help": "C:\\Program Files\\Autodesk\\3ds Max 2024\\help",
            "help_ln": "C:\\Program Files\\Autodesk\\3ds Max 2024\\help",
            "tutorials": "C:\\Program Files\\Autodesk\\3ds Max 2024\\tutorials",
            "tutorials_ln": "C:\\Program Files\\Autodesk\\3ds Max 2024\\tutorials",
            "samples": "C:\\Program Files\\Autodesk\\3ds Max 2024\\samples",
            "samples_ln": "C:\\Program Files\\Autodesk\\3ds Max 2024\\samples",
            "sceneassets": "C:\\Users\\Username\\Documents\\3ds Max 2024\\sceneassets",
            "sceneassets_ln": "C:\\Users\\Username\\Documents\\3ds Max 2024\\sceneassets",
            "renderPresets": "C:\\Users\\Username\\Documents\\3ds Max 2024\\renderpresets",
            "renderPresets_ln": "C:\\Users\\Username\\Documents\\3ds Max 2024\\renderpresets",
            "renderElementPresets": "C:\\Users\\Username\\Documents\\3ds Max 2024\\renderpresets\\render_elementpresets",
            "renderElementPresets_ln": "C:\\Users\\Username\\Documents\\3ds Max 2024\\renderpresets\\render_elementpresets",
            "renderElementUserPresets": "C:\\Users\\Username\\Documents\\3ds Max 2024\\renderpresets\\render_elementuserpresets",
            "renderElementUserPresets_ln": "C:\\Users\\Username\\Documents\\3ds Max 2024\\renderpresets\\render_elementuserpresets",
            "renderElementIcons": "C:\\Program Files\\Autodesk\\3ds Max 2024\\UI_ln\\Icons\\render_elements",
            "renderElementIcons_ln": "C:\\Program Files\\Autodesk\\3ds Max 2024\\UI_ln\\Icons\\render_elements",
            "renderElementUserIcons": "C:\\Users\\Username\\Documents\\3ds Max 2024\\usermacros\\icons\\render_elements",
            "renderElementUserIcons_ln": "C:\\Users\\Username\\Documents\\3ds Max 2024\\usermacros\\icons\\render_elements",
        }
        return dirs.get(which.lower(), "C:\\")


class MockRenderers:
    """Mock implementation of renderers"""

    def __init__(self):
        self.current = MockRenderer()

    def __getitem__(self, item):
        return self.current

    def __setitem__(self, key, value):
        pass


class MockRenderer:
    """Mock implementation of a renderer"""

    def __init__(self):
        self.output_width = 1920
        self.output_height = 1080
        self.timeType = 1  # Single frame
        self.nthFrame = 1
        self.start = 0
        self.end = 100
        self.output_saveFile = False
        self.output_file = ""
        self.progressCallback = None
        self.silentMode = True

    def __setattr__(self, key, value):
        # Allow setting any attribute
        self.__dict__[key] = value


class MockVray:
    """Mock implementation of V-Ray specific functionality"""

    def __init__(self):
        self.vray = self
        self.adv_irradmap_mode = 0
        self.adv_irradmap_mode_type = 0
        self.adv_irradmap_autoSave = False
        self.adv_irradmap_autoSaveFile = ""
        self.adv_irradmap_loadFileName = ""
        self.adv_irradmap_autoSaveFileName = ""
        self.adv_irradmap_autoSaveOn = False
        self.adv_irradmap_loadFileName_type = 0
        self.adv_irradmap_autoSaveFileName_type = 0

    def __getattr__(self, name):
        # Return a callable that can be used as a function
        def method(*args, **kwargs):
            # Use the built-in print directly
            print(f"[MOCK] V-Ray: Called {name} with args: {args}, kwargs: {kwargs}")
            return None

        return method


# Create the runtime instance
runtime = MockRuntime()


def print_to_listener(*args, **kwargs):
    """Mock implementation of print function that would normally print to the 3ds Max listener"""
    print("[MOCK] ", end="")
    print(*args, **kwargs)


# Set up the module
sys.modules[__name__].runtime = runtime
sys.modules[__name__].rt = runtime

# Common MAXScript functions that might be used
Name = runtime.Name
Execute = runtime.Execute
GetQuietMode = runtime.GetQuietMode
SetQuietMode = runtime.SetQuietMode
AddRollout = runtime.AddRollout
CreateDialog = runtime.CreateDialog
GetMAXWindowPos = runtime.GetMAXWindowPos
GetDir = runtime.GetDir
