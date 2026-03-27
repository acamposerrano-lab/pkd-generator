"""
Hook para PyInstaller: asegurar que tkinter empaquete TCL/TK correctamente
"""
from PyInstaller.utils.hooks import collect_data_files

# Recopilar todos los datos de tkinter
datas = collect_data_files('tkinter')

# Módulos de tkinter que necesitamos
hiddenimports = [
    'tkinter',
    'tkinter.ttk',
    'tkinter.filedialog',
    'tkinter.messagebox',
]
