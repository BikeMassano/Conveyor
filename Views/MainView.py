import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

from Model.Constants import MIN_TRANSFER_TIME, MAX_TRANSFER_TIME, FLOAT_ENTRY_LENTH
from Model.Enums.OperationType import OperationType
from Model.Enums.ProductComplexity import ProductComplexity

""" Класс представления """
class MainView:
    # Конструктор класса
    def __init__(self, controller):
        self.controller = controller
        self.__window = tk.Tk()

        # Регистрация команды валидирования float данных
        self.__validate_float_cmd = (self.__window.register(self.__validate_float), '%P', FLOAT_ENTRY_LENTH)

        # Настройка окна
        self.__window.title("Вариант 19. Определение требуемого количества постов конвейерной линии.")
        self.__window.geometry("700x400")
        self.__window.minsize(750, 400)

        # Настройка сетки
        self.__window.columnconfigure(0, weight=1)
        self.__window.columnconfigure(1, weight=1)

        self.__window.rowconfigure(0, weight=1)
        self.__window.rowconfigure(1, weight=1)
        
        # Создание виджетов ввода данных

        # Виджет выбора типа операции
        self.operation_type_label = ttk.Label(self.__window, text="Вид операций:")
        self.operation_type_combo = ttk.Combobox(self.__window, values=[e.value for e in OperationType], state="readonly")
        self.operation_type_combo.bind("<<ComboboxSelected>>", self.__on_element_change)

        # Виджет выбора сложности изделия
        self.product_complexity_label = ttk.Label(self.__window, text="Сложность изделия:")
        self.product_complexity_combo = ttk.Combobox(self.__window, values=[e.value for e in ProductComplexity], state="readonly")
        self.product_complexity_combo.bind("<<ComboboxSelected>>", self.__on_element_change)

        # Виджет ввода средней продолжительности операции на доформовочном участке(мин):
        self.preform_operation_time_label = ttk.Label(self.__window, text=f"Средняя продолжительность операции на доформовочном участке (мин), {FLOAT_ENTRY_LENTH} символов:")
        # Валидация entry происходит при нажатии клавиши(validate="key"), %P - подстановка после применения изменений
        self.preform_operation_time_entry = ttk.Entry(self.__window, validate="key", validatecommand=self.__validate_float_cmd)

        # Виджет ввода средней продолжительности операции на формовочном участке(мин):
        self.form_operation_time_label = ttk.Label(self.__window, text=f"Средняя продолжительность операции на формовочном участке (мин), {FLOAT_ENTRY_LENTH} символов:")
        self.form_operation_time_entry = ttk.Entry(self.__window, validate="key", validatecommand=self.__validate_float_cmd)

        # Виджет ввода средней продолжительности операции на послеформовочном участке(мин):
        self.postform_operation_time_label = ttk.Label(self.__window, text=f"Средняя продолжительность операции на послеформовочном участке, {FLOAT_ENTRY_LENTH} символов:")
        self.postform_operation_time_entry = ttk.Entry(self.__window, validate="key", validatecommand=self.__validate_float_cmd)

        # Виджет ввода продолжительности цикла формования
        self.cycle_time_label = ttk.Label(self.__window, text=f"Продолжительность цикла формования (мин), {FLOAT_ENTRY_LENTH} символов:")
        self.cycle_time_entry = ttk.Entry(self.__window, validate="key", validatecommand=self.__validate_float_cmd)

        # Переменная для хранения значения ползунка
        self.transfer_time = tk.DoubleVar(value=1.5)
        # Виджет ввода продолжительности передвижения тележек
        self.transfer_time_label = ttk.Label(self.__window, text="Продолжительность передвижения тележек (мин):")

        self.transfer_time_entry = tk.Scale(
            self.__window,
            variable=self.transfer_time,
            from_=MIN_TRANSFER_TIME,
            to=MAX_TRANSFER_TIME,
            resolution=0.1,
            orient="horizontal",
            showvalue=True
            )
        self.transfer_time_entry.bind("<ButtonRelease-1>", self.__on_element_change)

        # Виджет ввода коэффициента
        self.denominator_var = tk.BooleanVar()
        self.denomirator_label = ttk.Label(self.__window, text="Использовать знаменатель коэффициента?")
        self.denomirator_check = ttk.Checkbutton(self.__window, variable=self.denominator_var)
        self.denomirator_check.bind("<ButtonRelease-1>", self.__on_element_change)

        # Кнопка расчета
        self.calculate_button = ttk.Button(self.__window, text="Рассчитать", command=self.__calculate)

        # Виджет вывода результата
        self.result_label = ttk.Label(self.__window, text="Результат:")
        self.result_value = tk.StringVar()
        self.result_display = ttk.Label(self.__window, textvariable=self.result_value)

        # Кнопка экспорта
        self.doc_button = ttk.Button(self.__window, text="Экспортировать *.docx", command=self.export_docx, state="disabled")
        self.xls_button = ttk.Button(self.__window, text="Экспортировать *.xlsx", command=self.export_xlsx, state="disabled")

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
    def __validate_float(self, new_value, max_length):
        if new_value == "":
            self.__disable_export_buttons()
            return True
        try:
            float(new_value)
            return (len(new_value) <= int(max_length))
        except ValueError:
            return False
        finally:
            self.__disable_export_buttons()
        
    def __on_element_change(self, event):
        self.__disable_export_buttons()

    def __calculate(self):
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
        
            # Если расчет прошел успешно, разрешаем экспорт
            self.__enable_export_buttons()
        except ValueError as e:
            self.__disable_export_buttons()
            self.result_value.set("Ошибка: Введены некорректные данные.")
            messagebox.showerror("Ошибка", e)
        except ZeroDivisionError:
            self.__disable_export_buttons()
            self.result_value.set("Ошибка: В ходе расчета произошло деление на ноль. Проверьте корректность данных.")
        except Exception as e:
            self.__disable_export_buttons()
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

    def __disable_export_buttons(self):
        self.doc_button.config(state="disabled")
        self.xls_button.config(state="disabled")

    def __enable_export_buttons(self):
        self.doc_button.config(state="normal")
        self.xls_button.config(state="normal")

    # Метод для заполнения поля результата
    def set_result(self, result):
        self.result_value.set(str(result))

    # Метод запуска главного цикла обработки событий окна
    def run(self):
        self.__window.mainloop()

