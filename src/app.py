from pathlib import Path
import tkinter as tk

from core.enums import ExcelApp
from core.graphical_interface import ExcelTableFilterApp

BASEDIR_PROJECT = Path(__file__).resolve().parents[1]

root = tk.Tk()
app = ExcelTableFilterApp(
    root=root,
    title=ExcelApp.TITLE,
)

if __name__ == '__main__':
    app.run()
    root.mainloop()
