import sys

from cx_Freeze import setup, Executable

import os.path
PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')

build_exe_options = dict(
        includes = ["imaplib", "email", "re", "requests", "os", "datetime", "time","urllib","exifread","unicodedata","csv","openpyxl","hashlib", "gmplot","idna.idnadata","numpy","numpy.core._methods"],
        include_files = [os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll'),
                         os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll'),
]
)

base = None



setup(  name = "GmailParser",
        version = "1.0",
        description = "GmailParser",
        author = "H33ro",
        options = {"build_exe" : build_exe_options},
        executables = [Executable("gmail.py", base = base, targetName="Gmail.exe")]
)



