from Model.Enums.OperationType import OperationType
from Model.Enums.ProductComplexity import ProductComplexity
from Model.Services.OperationCoefficients import OperationCoefficients
from Model.ConveyorLineSegment import ConveyorLineSegment

class CalculatorController:
    def __init__(self, view, calculator, export_service):
        # Текцущее представление
        self.view = view
        # Сервис для подсчёта количества постов
        self.calculator = calculator
        # Сервис экспорта
        self.export_service = export_service
        # Словарь коэффициентов
        self.coefficients = {
            (OperationType.AUTOMATED, ProductComplexity.SIMPLE): OperationCoefficients(OperationType.AUTOMATED, ProductComplexity.SIMPLE, 1.05),
            (OperationType.AUTOMATED, ProductComplexity.COMPLEX): OperationCoefficients(OperationType.AUTOMATED, ProductComplexity.COMPLEX, 1.05),
            (OperationType.MECHANIZED, ProductComplexity.SIMPLE): OperationCoefficients(OperationType.MECHANIZED, ProductComplexity.SIMPLE, 1.15, 1.10),
            (OperationType.MECHANIZED, ProductComplexity.COMPLEX): OperationCoefficients(OperationType.MECHANIZED, ProductComplexity.COMPLEX, 1.25, 1.15),
            (OperationType.MANUAL, ProductComplexity.SIMPLE): OperationCoefficients(OperationType.MANUAL, ProductComplexity.SIMPLE, 1.25, 1.15),
            (OperationType.MANUAL, ProductComplexity.COMPLEX): OperationCoefficients(OperationType.MANUAL, ProductComplexity.COMPLEX, 1.35, 1.20),
        }
    def calculate_and_display(self, operation_type, product_complexity, preform_operation_time, form_operation_time, postform_operation_time, cycle_time, transfer_time, use_denominator):
        segment = ConveyorLineSegment(operation_type, product_complexity, preform_operation_time, form_operation_time, postform_operation_time, cycle_time, transfer_time)

        # Поиск коэффициента для заданной пары значений
        coefficient = self.coefficients.get((operation_type, product_complexity))
        if not coefficient:
            self.view.set_result("Ошибка: Не найден коэффициент для указанных параметров")
            return

        # Вычисление количества постов
        result = self.calculator.calculate_posts(segment, coefficient, use_denominator)

        # Форматирование результата
        formatted_result = "{:.2f}".format(result)

        # Отображение результата
        self.view.set_result(formatted_result)

    def export_to_docx(self, filename, data):
        self.export_service.export_to_docx(filename, data)

    def export_to_xlsx(self, filename, data):
        self.export_service.export_to_xlsx(filename, data)