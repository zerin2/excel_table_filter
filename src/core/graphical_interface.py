import os
import tkinter as tk
from tkinter import filedialog, messagebox

from core.enums import ExcelApp, ExcelAppSettings
from core.exceptions import InvalidFilterValue, InvalidFilterColumnName
from core.excel_filter import excel_service



class ExcelTableFilterApp:
    """Графический интерфейс для фильтрации Excel-файлов с помощью tkinter."""
    def __init__(self, root, title: str):
        self.root = root
        self.root.title(title)
        self.service = excel_service

        self.file_path = None
        self.status_var = 'Статус: {status}'
        self.status_label = None
        self.filter_column = None
        self.filter_entry = None

    def run(self):
        """Запускает отображение всех элементов интерфейса."""
        self.init_buttons()
        self.init_status_labels()
        self.init_filter_column()
        self.init_filter_entry()

    def init_status_labels(self):
        """Создаёт и размещает метку для отображения статуса приложения."""
        self.status_label = tk.Label(
            self.root,
            text=self.status_var.format(status=ExcelApp.WAITING_FILE),
            fg='green'
        )
        self.status_label.grid(row=4, column=0, columnspan=2, pady=10)

    def set_status(self, text: str):
        """Обновляет текст статуса на экране."""
        self.status_label.config(text=self.status_var.format(status=text))

    def init_buttons(self):
        """Создаёт и размещает основные кнопки интерфейса."""
        tk.Button(
            self.root,
            text=ExcelApp.OPEN_FILE,
            command=self.open_file
        ).grid(row=0, column=0, padx=5, pady=5)
        tk.Button(
            self.root,
            text=ExcelApp.SAVE_AS,
            command=self.save_as
        ).grid(row=3, column=0, padx=5, pady=5)

        tk.Button(
            self.root,
            text=ExcelApp.START,
            command=self.start_processing
        ).grid(row=3, column=1, padx=5, pady=5)

    def init_filter_column(self):
        """Создаёт поле ввода для названия столбца фильтрации."""
        tk.Label(
            self.root,
            text=ExcelApp.COLUMN_FILTER,
        ).grid(row=1, column=0, sticky='w', padx=5)
        self.filter_column = tk.Entry(self.root)
        self.filter_column.grid(row=1, column=1, padx=5)

    def init_filter_entry(self):
        """Создаёт поле ввода для значения фильтрации."""
        tk.Label(
            self.root,
            text=ExcelApp.VALUE_FILTER,
        ).grid(row=2, column=0, sticky='w', padx=5)
        self.filter_entry = tk.Entry(self.root)
        self.filter_entry.grid(row=2, column=1, padx=5)

    def open_file(self):
        """Открывает диалог выбора файла."""
        self.file_path = filedialog.askopenfilename(
            title=ExcelApp.OPEN_EXCEL_FILE,
            filetypes=[ExcelAppSettings.EXCEL_TYPES.value]
        )
        self.service.upload_file_path = self.file_path
        if self.file_path:
            self.set_status(ExcelApp.DOWNLOAD_FILE + self.file_path)

    def save_as(self):
        """Открывает диалог выбора пути сохранения."""
        full_path = filedialog.asksaveasfilename(
            title=ExcelApp.SAVE_AS,
            defaultextension=ExcelAppSettings.DEFAULT_EXTENSION.value,
            filetypes=[ExcelAppSettings.EXCEL_TYPES.value]
        )
        if full_path:
            dir_path, file_name = os.path.split(full_path)
            self.service.save_file_path = dir_path
            self.service.save_file_name = file_name
            self.set_status(ExcelApp.SAVE_IN.format(save_path=full_path))

    def start_processing(self):
        """Запускает процесс фильтрации:
        - проверяет ввод;
        - передаёт параметры в сервис;
        - выполняет фильтрацию и сохранение;
        - обрабатывает возможные ошибки.
        """
        if not self.file_path:
            messagebox.showwarning(
                title=ExcelApp.ERROR,
                message=ExcelApp.ERROR_FILE_PATH_MSG,
            )
            return

        filter_column = self.filter_column.get()
        filter_value = self.filter_entry.get()

        if not filter_column or not filter_value:
            messagebox.showwarning(
                title=ExcelApp.ERROR,
                message=ExcelApp.ERROR_FILTER_VALUES,
            )
            return

        try:
            self.service.filter_column = filter_column
            self.service.filter_value = filter_value
            self.service.run_filter()
        except InvalidFilterColumnName as e:
            messagebox.showwarning(
                title=ExcelApp.ERROR_FILTER,
                message=str(e),
            )
            return
        except InvalidFilterValue as e:
            messagebox.showwarning(
                title=ExcelApp.ERROR_FILTER,
                message=str(e),
            )
            return
        except Exception as e:
            messagebox.showwarning(
                title=ExcelApp.ERROR,
                message=str(e),
            )
            return
        else:
            self.set_status(ExcelApp.FILTERING_BY.format(
                filter_column=filter_column,
                filter_value=filter_value
            ))
            messagebox.showinfo(
                title=ExcelApp.READY,
                message=ExcelApp.FILE_SUCCESS,
            )
            return
