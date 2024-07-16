import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import openpyxl
import openpyxl.utils
import openpyxl.utils.exceptions

from exceptions import (ERROR_MESSAGES,
                        ExcelAppError,
                        FileNotFoundError,
                        InvalidFileException,
                        OutOfRangeError,
                        ValueError,
                        ProcessorNotLoadedError)
from decorators import requires_processor
from table import ExcelRangeProcessor


class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Range Processor")
        self.entry_width = 40
        self.button_width = 25
        self.folder_path = None
        self.file_paths = ["./data/coordinates.xlsx"]
        self.create_widgets()

    def create_widgets(self):
        """Создание основных виджетов для приложения."""
        self.set_default_values()
        self.create_file_widgets()
        self.create_range_widgets()
        self.create_cell_widgets()
        self.create_row_widgets()
        self.create_column_widgets()
        self.create_search_widgets()
        self.create_result_widgets()

    def set_default_values(self):
        """Установка значений по умолчанию для переменных приложения."""
        self.selected_path = tk.StringVar(value=self.file_paths[0])
        self.range_str = tk.StringVar(value="A1:D35")
        self.cell_address = tk.StringVar(value="C1")
        self.row_number = tk.StringVar(value=1)
        self.col_letter = tk.StringVar(value="A")
        self.search_word = tk.StringVar(value="нет данных")
        self.copy_message = tk.StringVar()

    def create_file_widgets(self):
        """Создание виджетов для выбора файла или папки."""
        ttk.Label(
            self.root, text="Файлы или папка Excel:"
        ).grid(row=0, column=0, padx=10, pady=5, sticky="w")

        # Поле для отображения выбранного пути
        path_entry_width = self.entry_width - 10  # Уменьшаем ширину строки
        ttk.Entry(
            self.root, textvariable=self.selected_path, width=path_entry_width
        ).grid(row=0, column=1, padx=10, pady=5, sticky="we", columnspan=1)

        # Кнопки для выбора файлов и папок
        half_button_width = 10
        frame_buttons = ttk.Frame(self.root)
        frame_buttons.grid(row=0, column=2, padx=10, pady=5, sticky="w")
        ttk.Button(
            frame_buttons,
            text="Файл",
            command=self.browse_files,
            width=half_button_width
        ).grid(row=0, column=0, padx=(0, 5), pady=5)
        ttk.Button(
            frame_buttons,
            text="Папка",
            command=self.browse_folder,
            width=half_button_width
        ).grid(row=0, column=1, padx=(5, 0), pady=5)

    def create_range_widgets(self):
        """Создание виджетов для ввода диапазона ячеек."""
        ttk.Label(
            self.root, text="Диапазон ячеек:"
        ).grid(row=1, column=0, padx=10, pady=5, sticky="w")
        ttk.Entry(
            self.root,
            textvariable=self.range_str,
            width=self.entry_width
        ).grid(row=1, column=1, padx=10, pady=5)
        ttk.Button(
            self.root, text="Загрузить данные",
            command=self.load_data,
            width=self.button_width
        ).grid(row=1, column=2, padx=10, pady=5)

    def create_cell_widgets(self):
        """Создание виджетов для ввода адреса ячейки."""
        ttk.Label(
            self.root, text="Введите адрес ячейки:"
        ).grid(row=2, column=0, padx=10, pady=5, sticky="w")
        ttk.Entry(
            self.root,
            textvariable=self.cell_address,
            width=self.entry_width
        ).grid(row=2, column=1, padx=10, pady=5)
        ttk.Button(
            self.root, text="Показать значение ячейки",
            command=self.get_cell_value,
            width=self.button_width
        ).grid(row=2, column=2, padx=10, pady=5)

    def create_row_widgets(self):
        """Создание виджетов для ввода номера строки."""
        ttk.Label(
            self.root, text="Введите номер строки:"
        ).grid(row=3, column=0, padx=10, pady=5, sticky="w")
        ttk.Entry(
            self.root,
            textvariable=self.row_number,
            width=self.entry_width
        ).grid(row=3, column=1, padx=10, pady=5)
        ttk.Button(
            self.root,
            text="Показать данные со строки",
            command=self.get_row,
            width=self.button_width
        ).grid(row=3, column=2, padx=10, pady=5)

    def create_column_widgets(self):
        """Создание виджетов для ввода названия столбца."""
        ttk.Label(
            self.root, text="Введите название столбца:"
        ).grid(row=4, column=0, padx=10, pady=5, sticky="w")
        ttk.Entry(
            self.root,
            textvariable=self.col_letter,
            width=self.entry_width
        ).grid(row=4, column=1, padx=10, pady=5)
        ttk.Button(
            self.root,
            text="Показать данные со столбца",
            command=self.get_column,
            width=self.button_width
        ).grid(row=4, column=2, padx=10, pady=5)

    def create_search_widgets(self):
        """Создание виджетов для поиска."""
        ttk.Label(
            self.root, text="Введите слово для поиска:"
        ).grid(row=5, column=0, padx=10, pady=5, sticky="w")
        ttk.Entry(
            self.root,
            textvariable=self.search_word,
            width=self.entry_width
        ).grid(row=5, column=1, padx=10, pady=5)
        ttk.Button(
            self.root, text="Найти слово", command=self.find_word,
            width=self.button_width
        ).grid(row=5, column=2, padx=10, pady=5)

    def create_result_widgets(self):
        """Создание виджетов для отображения результатов."""
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_columnconfigure(2, weight=1)
        self.root.grid_columnconfigure(3, weight=1)
        self.root.grid_columnconfigure(4, weight=1)

        ttk.Button(
            self.root, text="Показать весь Массив",
            command=self.show_full_array,
            width=self.button_width
        ).grid(row=6, column=0, columnspan=5, padx=10, pady=10)

        ttk.Label(
            self.root, text="Результат:"
        ).grid(row=7, column=0, columnspan=5, padx=10, pady=10, sticky="w")

        self.result_text = tk.Text(
            self.root, wrap="word", height=15, width=30
        )
        self.result_text.grid(
            row=8, column=0, columnspan=5, padx=10, pady=10, sticky="nsew"
        )

        scrollbar = ttk.Scrollbar(
            self.root, orient="vertical", command=self.result_text.yview
        )
        scrollbar.grid(row=8, column=5, sticky="ns")
        self.result_text.config(yscrollcommand=scrollbar.set)

        self.root.grid_rowconfigure(8, weight=1)

    def browse_files(self):
        """Открыть диалоговое окно для выбора нескольких
        файлов Excel и установить пути к файлам."""
        file_paths = filedialog.askopenfilenames(
            filetypes=[("Excel files", "*.xlsx")])
        self.selected_path.set(file_paths)
        if file_paths:
            self.file_paths = list(file_paths)
            self.folder_path = None

    def browse_folder(self):
        """Открыть диалоговое окно для выбора
        папки и установить путь к папке."""
        folder_path = filedialog.askdirectory()
        self.selected_path.set(folder_path)
        if folder_path:
            self.folder_path = folder_path
            self.file_paths = []

    def load_data(self):
        """Загрузить данные из указанных файлов или из папки Excel."""
        try:
            if self.folder_path:
                self.processor = ExcelRangeProcessor(
                    folder_path=self.folder_path,
                    range_str=self.range_str.get())
            elif self.file_paths:
                self.processor = ExcelRangeProcessor(
                    file_paths=self.file_paths,
                    range_str=self.range_str.get()
                )
            else:
                raise FileNotFoundError(
                    "Не выбраны файлы или папка для загрузки.")

            self.show_info("Успех", "Данные успешно загружены!")
        except FileNotFoundError:
            self.handle_error(FileNotFoundError)
        except openpyxl.utils.exceptions.InvalidFileException:
            self.handle_error(InvalidFileException)
        except Exception as e:
            self.handle_error(ExcelAppError, e)

    @requires_processor
    def get_cell_value(self):
        """Получить значение ячейки по указанному адресу."""
        if not self.validate_input(self.cell_address.get(),
                                   "Введите адрес ячейки."):
            return
        if not hasattr(self, 'processor'):
            self.handle_error(ProcessorNotLoadedError)
            return
        try:
            self.result_text.delete(1.0, tk.END)
            value = self.processor.get_cell_value(self.cell_address.get())
            result_str = f"Значение в {self.cell_address.get()}: {value}\n"
            self.result_text.insert(tk.END, result_str)
        except IndexError:
            self.handle_error(OutOfRangeError)
        except Exception as e:
            self.handle_error(ExcelAppError, e)

    @requires_processor
    def get_row(self):
        """Получить данные из указанной строки."""
        if not self.validate_input(self.row_number.get(),
                                   "Введите номер строки."):
            return
        try:
            self.result_text.delete(1.0, tk.END)
            row = self.processor.get_row(int(self.row_number.get()))
            result_str = f"Строка {self.row_number.get()}: {row}\n"
            self.result_text.insert(tk.END, result_str)
        except IndexError:
            self.handle_error(OutOfRangeError)
        except ValueError:
            self.handle_error(ValueError)
        except Exception as e:
            self.handle_error(ExcelAppError, e)

    @requires_processor
    def get_column(self):
        """Получить данные из указанного столбца."""
        if not self.validate_input(self.col_letter.get(),
                                   "Введите букву столбца."):
            return
        try:
            self.result_text.delete(1.0, tk.END)
            column = self.processor.get_column(self.col_letter.get().upper())
            result_str = f"Столбец {self.col_letter.get()}: {column}\n"
            self.result_text.insert(tk.END, result_str)
        except IndexError:
            self.handle_error(OutOfRangeError)
        except Exception as e:
            self.handle_error(ExcelAppError, e)

    @requires_processor
    def find_word(self):
        """Найти указанное слово в диапазоне ячеек."""
        if not self.validate_input(self.search_word.get(),
                                   "Введите искомое слово."):
            return
        try:
            self.result_text.delete(1.0, tk.END)
            addresses = self.processor.find_word(self.search_word.get())
            result_str = \
                f"Адреса, содержащие {self.search_word.get()}:\n{addresses}\n"
            self.result_text.insert(tk.END, result_str)
        except Exception as e:
            self.handle_error(ExcelAppError, e)

    @requires_processor
    def show_full_array(self):
        """Показать полный массив данных."""
        if hasattr(self, 'processor'):
            self.result_text.delete(1.0, tk.END)
            table = self.processor.My_table
            result_str = str(table)
            self.result_text.insert(tk.END, result_str + "\n")
        else:
            self.show_warning("Предупреждение",
                              "Пожалуйста, загрузите данные сначала.")

    def handle_error(self, error_type, error=None):
        """Обработать ошибки и вывести соответствующее сообщение."""
        error_message = ERROR_MESSAGES.get(error_type,
                                           ERROR_MESSAGES[ExcelAppError])
        messagebox.showerror("Ошибка", error_message.format(error=error))

    def show_info(self, title, message):
        """Показать информационное сообщение."""
        messagebox.showinfo(title, message)

    def show_warning(self, title, message):
        """Показать предупреждающее сообщение."""
        messagebox.showwarning(title, message)

    def validate_input(self, value, message):
        """Проверить ввод пользователя и показать
        предупреждение, если значение пустое."""
        if not value:
            self.show_warning("Предупреждение", message)
            return False
        return True


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelApp(root)
    root.mainloop()
