from dataclasses import dataclass
from enum import StrEnum, Enum


class TableSettings(Enum):
    DEFAULT_RESULT_HEADERS = (
        ['ФИО', 'Должность', 'Отдел', 'Дата найма', 'Зарплата']
    )
    TABLE_HEADER_INDICATOR = '№ п/п'

@dataclass(frozen=True, slots=True)
class ExcelApp:
    TITLE = 'Фильтр Excel-файла'
    OPEN_FILE = 'Открыть файл'
    DOWNLOAD_FILE = 'Файл загружен: '
    OPEN_EXCEL_FILE = 'Открыть Excel файл'
    SAVE_AS = 'Сохранить как'
    SAVE_IN = 'Сохранить в {save_path}'
    COLUMN_FILTER = 'Столбец для фильтрации:'
    VALUE_FILTER = 'Значение для фильтрации:'
    START = 'start'
    WAITING_FILE = 'Ожидание выбора файла'
    ERROR = 'Ошибка'
    ERROR_FILTER = 'Ошибка фильтрации'
    READY = 'Готово'
    ERROR_FILE_PATH_MSG = 'Сначала выберите файл.'
    ERROR_FILTER_VALUES = 'Укажите столбец и значение.'
    FILE_SUCCESS = 'Файл успешно обработан.'
    FILTERING_BY = 'Фильтрация по: {filter_column} = {filter_value}'


class ExcelAppSettings(Enum):
    DEFAULT_EXTENSION = '.xlsx'
    EXCEL_TYPES = ('Excel файлы', '*.xlsx *.xls')
    STATUS_LENGTH = 45
