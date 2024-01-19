from cx_Freeze import setup, Executable

base = None

executables = [Executable("main.py", base=base)]

packages = ["idna", "openpyxl"]
options = {
    'build_exe': {
        'packages': packages,
    },
}

setup(
    name="ExcelSearchGUI",
    options=options,
    version="1.0",
    description="Search Excel GUI",
    executables=executables
)
