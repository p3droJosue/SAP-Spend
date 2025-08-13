import os
from tkinter import Tk
from tkinter.filedialog import askdirectory

def get_path_dates():
    Tk().withdraw()
    folder_path = askdirectory(title="Selecciona la carpeta de descarga")
    folder_path = os.path.normpath(folder_path)  # <-- normaliza slashes
    Date_From = input("Desde qué fecha deseas descargar (dd.mm.yyyy): ")
    Date_To = input("Hasta qué fecha deseas descargar (dd.mm.yyyy): ")
    return folder_path, Date_From, Date_To
