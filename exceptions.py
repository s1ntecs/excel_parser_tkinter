class ExcelAppError(Exception):
    """Базовый класс для всех исключений приложения ExcelApp."""
    pass


class FileNotFoundError(ExcelAppError):
    """Исключение для отсутствующего файла."""
    pass


class InvalidFileException(ExcelAppError):
    """Исключение для неверного формата файла."""
    pass


class OutOfRangeError(ExcelAppError):
    """Исключение для адреса ячейки или строки вне диапазона."""
    pass


class ValueError(ExcelAppError):
    """Исключение для некорректного значения."""
    pass


class ProcessorNotLoadedError(ExcelAppError):
    """Исключение для случая, когда процессор не был загружен."""
    pass


ERROR_MESSAGES = {
    ProcessorNotLoadedError: "Пожалуйста, загрузите данные сначала.",
    FileNotFoundError: "Файл не найден. Пожалуйста, выберите правильный файл.",
    InvalidFileException: "Неверный формат файла. Пожалуйста, выберите файл Excel.",
    OutOfRangeError: "Адрес ячейки или номер строки вне диапазона. Пожалуйста, введите корректные данные.",
    ValueError: "Пожалуйста, введите числовое значение.",
    ExcelAppError: "Произошла ошибка: {error}",
}