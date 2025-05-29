import sys
import os
import debugpy

def startup():
    sysexec = sys.executable
    (base, file) = os.path.split(sys.executable)
    if file.lower() == "3dsmax.exe":
        sys.executable = os.path.join(base, "python", "python.exe")
        host = "localhost"
        port = 5678
        debugpy.listen((host, port))
        print(f"-- now ready to receive debugging connections from vscode on (${host}, ${port})")
        sys.executable = sysexec

startup()