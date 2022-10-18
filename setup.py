import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
# "packages": ["os"] is used as example only
build_exe_options = {"packages": ["sqlalchemy"],
                    "include_files":
                        ['images', 'README.md', 'LICENSE.txt',
                        'manpower.ico', 'screenshots', 'index.html']}

# base="Win32GUI" should be used only for Windows GUI app
base = None

if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Manpower",
    version="2.0.1",
    description="A simple manpower management software",
    options={"build_exe": build_exe_options},
    executables=[Executable("manpower.py",
                           copyright="Copyright (C) 2022 Jesus Vedasto Olazo",
                           base=base,
                           icon="manpower.ico",
                           )],
)
