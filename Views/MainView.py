import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

from Model.Constants import MIN_TRANSFER_TIME, MAX_TRANSFER_TIME
from Model.Enums.OperationType import OperationType
from Model.Enums.ProductComplexity import ProductComplexity

""" Класс представления """
class MainView:
    # Конструктор класса
    def __init__(self, controller):
        self.controller = controller
        self.window = tk.Tk()

        # Регистрация команды валидирования float данных
        self.validate_float_cmd = self.window.register(self.validate_float)

        # Настройка окна
        self.window.title("Вариант 19. Определение требуемого количества постов конвейерной линии.")
        self.window.geometry("700x400")
        self.window.minsize(700, 400)

        # Настройка сетки
        self.window.columnconfigure(0, weight=1)
        self.window.columnconfigure(1, weight=1)

        self.window.rowconfigure(0, weight=1)
        self.window.rowconfigure(1, weight=1)
        
        # Создание виджетов ввода данных

        # Виджет выбора типа операции
        self.operation_type_label = ttk.Label(self.window, text="Вид операций:")
        self.operation_type_combo = ttk.Combobox(self.window, values=[e.value for e in OperationType], state="readonly")

        # Виджет выбора сложности изделия
        self.product_complexity_label = ttk.Label(self.window, text="Сложность изделия:")
        self.product_complexity_combo = ttk.Combobox(self.window, values=[e.value for e in ProductComplexity], state="readonly")

        # Виджет ввода средней продолжительности операции на доформовочном участке(мин):
        self.preform_operation_time_label = ttk.Label(self.window, text="Средняя продолжительность операции на доформовочном участке(мин):")
        # Валидация entry происходит при нажатии клавиши(validate="key"), %P - подстановка после применения изменений
        self.preform_operation_time_entry = ttk.Entry(self.window, validate="key", validatecommand=(self.validate_float_cmd, '%P'))

        # Виджет ввода средней продолжительности операции на формовочном участке(мин):
        self.form_operation_time_label = ttk.Label(self.window, text="Средняя продолжительность операции на формовочном участке(мин):")
        self.form_operation_time_entry = ttk.Entry(self.window, validate="key", validatecommand=(self.validate_float_cmd, '%P'))

        # Виджет ввода средней продолжительности операции на послеформовочном участке(мин):
        self.postform_operation_time_label = ttk.Label(self.window, text="Средняя продолжительность операции на послеформовочном участке(мин):")
        self.postform_operation_time_entry = ttk.Entry(self.window, validate="key", validatecommand=(self.validate_float_cmd, '%P'))

        # Виджет ввода продолжительности цикла формования
        self.cycle_time_label = ttk.Label(self.window, text="Продолжительность цикла формования (мин):")
        self.cycle_time_entry = ttk.Entry(self.window, validate="key", validatecommand=(self.validate_float_cmd, '%P'))

        # Переменная для хранения значения ползунка
        self.transfer_time = tk.DoubleVar(value=1.5)
        # Виджет ввода продолжительности передвижения тележек
        self.transfer_time_label = ttk.Label(self.window, text="Продолжительность передвижения тележек (мин):")

        self.transfer_time_entry = tk.Scale(
            self.window,
            variable=self.transfer_time,
            from_=MIN_TRANSFER_TIME,
            to=MAX_TRANSFER_TIME,
            resolution=0.1,
            orient="horizontal",
            showvalue=True
            )

        # Виджет ввода коэффициента
        self.denominator_var = tk.BooleanVar()
        self.denomirator_label = ttk.Label(self.window, text="Использовать знаменатель коэффициента?")
        self.denomirator_check = ttk.Checkbutton(self.window, variable=self.denominator_var)

        # Кнопка расчета
        self.calculate_button = ttk.Button(self.window, text="Рассчитать", command=self.calculate)

        # Виджет вывода результата
        self.result_label = ttk.Label(self.window, text="Результат:")
        self.result_value = tk.StringVar()
        self.result_display = ttk.Label(self.window, textvariable=self.result_value)

        # Кнопка экспорта
        self.doc_button = ttk.Button(self.window, text="Экспортировать *.docx", command=self.export_docx)
        self.xls_button = ttk.Button(self.window, text="Экспортировать *.xlsx", command=self.export_xlsx)

        # Размещение виджетов

        self.operation_type_label.grid(row=0, column=0, sticky=tk.EW, padx=5, pady=5)
        self.operation_type_combo.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)

        self.product_complexity_label.grid(row=1, column=0, sticky=tk.EW, padx=5, pady=5)
        self.product_complexity_combo.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=5)

        self.preform_operation_time_label.grid(row=2, column=0, sticky=tk.EW, padx=5, pady=5)
        self.preform_operation_time_entry.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=5)

        self.form_operation_time_label.grid(row=3, column=0, sticky=tk.EW, padx=5, pady=5)
        self.form_operation_time_entry.grid(row=3, column=1, sticky=tk.EW, padx=5, pady=5)

        self.postform_operation_time_label.grid(row=4, column=0, sticky=tk.EW, padx=5, pady=5)
        self.postform_operation_time_entry.grid(row=4, column=1, sticky=tk.EW, padx=5, pady=5)

        self.cycle_time_label.grid(row=5, column=0, sticky=tk.EW, padx=5, pady=5)
        self.cycle_time_entry.grid(row=5, column=1, sticky=tk.EW, padx=5, pady=5)

        self.transfer_time_label.grid(row=6, column=0, sticky=tk.EW, padx=5, pady=5)
        self.transfer_time_entry.grid(row=6, column=1, sticky=tk.EW, padx=5, pady=5)

        self.denomirator_label.grid(row=7, column=0, sticky=tk.EW, padx=5, pady=5)
        self.denomirator_check.grid(row=7, column=1, sticky=tk.EW, padx=5, pady=5)

        self.calculate_button.grid(row=8, column=0, columnspan=2, sticky=tk.EW, padx=5, pady=5)

        self.result_label.grid(row=9, column=0, sticky=tk.EW, padx=5, pady=5)
        self.result_display.grid(row=9, column=1, sticky=tk.EW, padx=5, pady=5)

        self.doc_button.grid(row=10, column=0, columnspan=2, sticky=tk.EW, padx=5, pady=5)
        self.xls_button.grid(row=11, column=0, columnspan=2, sticky=tk.EW, padx=5, pady=5)

    """Метод валидации чисел с плавающей точкой"""
    def validate_float(self, new_value):
        if new_value == "":
            return True
        try:
            float(new_value)
            return True
        except ValueError:
            return False
        
    def calculate(self):
        try:
            # Извлечение значений из полей ввода
            operation_type = OperationType(self.operation_type_combo.get())
            product_complexity = ProductComplexity(self.product_complexity_combo.get())
            preform_operation_time = float(self.preform_operation_time_entry.get())
            form_operation_time = float(self.form_operation_time_entry.get())
            postform_operation_time = float(self.postform_operation_time_entry.get())
            cycle_time = float(self.cycle_time_entry.get())
            transfer_time = float(self.transfer_time_entry.get())
            use_denominator = bool(self.denominator_var.get())

            # Вызов метода расчета в контроллере
            self.controller.calculate_and_display(operation_type, product_complexity, preform_operation_time, form_operation_time, postform_operation_time, cycle_time, transfer_time, use_denominator)
        except ValueError as e:
            self.result_value.set("Ошибка: Введены некорректные данные.")
            messagebox.showerror("Ошибка", e)
        except ZeroDivisionError:
            self.result_value.set("Ошибка: В ходе расчета произошло деление на ноль. Проверьте корректность данных.")
        except Exception as e:
            self.result_value.set("Ошибка: Произошла непредвиденная ошибка.")
            messagebox.showerror("Ошибка", e)

    def export_docx(self):
        data = self.get_input_data()
        data["Результат"] = self.result_value.get()
        try:
            self.controller.export_to_docx("export.docx", data)
        except PermissionError:
            messagebox.showerror("Ошибка", "Недостаточно прав для доступа к файлу. Файл может быть открыт в текущий момент.")

    def export_xlsx(self):
        data = self.get_input_data()
        data["Результат"] = self.result_value.get()
        try:
            self.controller.export_to_xlsx("export.xlsx", data)
        except PermissionError:
            messagebox.showerror("Ошибка", "Недостаточно прав для доступа к файлу. Файл может быть открыт в текущий момент.")

    def get_input_data(self):
        return {
            "Вид операций": self.operation_type_combo.get(),
            "Сложность изделия": self.product_complexity_combo.get(),
            "Средняя продолжительность операции на доформовочном участке(мин)": self.preform_operation_time_entry.get(),
            "Средняя продолжительность операции на формовочном участке(мин)": self.form_operation_time_entry.get(),
            "Средняя продолжительность операции на послеформовочном участке(мин)": self.postform_operation_time_entry.get(),
            "Продолжительность цикла формования (мин)": self.cycle_time_entry.get(),
            "Продолжительность передвижения тележек (мин)": self.transfer_time.get(),
            "Использовать знаменатель коэффициента?": self.denominator_var.get()
        }

    # Метод для заполнения поля результата
    def set_result(self, result):
        self.result_value.set(str(result))

    # Метод запуска главного цикла обработки событий окна
    def run(self):
        self.window.mainloop()