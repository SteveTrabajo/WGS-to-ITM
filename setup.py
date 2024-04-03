# import sys
# from cx_Freeze import setup, Executable

# base = None
# if sys.platform == "win32":
#     base = "Win32GUI"

# setup(
#     name="Natan's WGS84 TO ITM Converter",
#     version="1.0",
#     description="Converts WGS84 to ITM",
#     executables=[Executable("main.py", base=base)],
# )

from cx_Freeze import setup, Executable
import os
import sys

base = None
if sys.platform == "win32":
    base = "Win32GUI"
    
PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')

options = {
    'build_exe': {
        'include_files':[
            os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll'),
            os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll'),
         ],
    },
}

setup(name="Natan's WGS84 TO ITM Converter",
    version="1.0",
    description="Converts WGS84 to ITM",
      options=options,
      executables = [Executable("main.py", base=base, target_name="nwgti.exe")])