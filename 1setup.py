from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine-tuning.
build_options = {
    'packages': ['openpyxl'],
    'excludes': [],
    'include_files': [('C:/Users/Michael.Guilmette/OneDrive - UKG/Documents/Programming Projects/DB Tables Search/data', 'data')]
}

base = 'console'

executables = [
    Executable('main.py', base=base, target_name='dbsearch')
]

setup(name='DB Search',
      version='1.0',
      description='Database Search',
      options={'build_exe': build_options},
      executables=executables)
