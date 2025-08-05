class CustomBaseException(Exception):
    def __init__(self, message):
        self.message = message
        super().__init__(message)

    def __repr__(self):
        return f'({self.__class__.__name__}): {self.message}'

    def __str__(self):
        return self.__repr__()


class InvalidFilterValue(CustomBaseException):
    def __init__(self, message='Указанное значение не найдено'):
        super().__init__(message)


class InvalidFilterColumnName(CustomBaseException):
    def __init__(self, message='Указанная колонка не найдена в заголовках'):
        super().__init__(message)
